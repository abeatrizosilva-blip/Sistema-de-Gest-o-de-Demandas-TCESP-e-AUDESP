from datetime import date, datetime
import io
import html
import hashlib
import json
import os
import socket
import sqlite3
import tempfile
from threading import RLock
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen

import bcrypt
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from flask import Flask, jsonify, redirect, render_template_string, request, send_file, session, url_for


app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "TROQUE-ESTA-CHAVE-POR-UMA-CHAVE-SECRETA")
CAMINHO_ONEDRIVE_WINDOWS = r"C:\Users\ana.silva\OneDrive - PRODESP\SP_AGUAS\Sistema de Gestão de Demandas - SP Aguas.xlsx"
PLANILHA = os.environ.get("EXCEL_DATABASE", os.environ.get("DATABASE", CAMINHO_ONEDRIVE_WINDOWS if os.name == "nt" and os.path.exists(CAMINHO_ONEDRIVE_WINDOWS) else "/tmp/sp_aguas.xlsx" if os.environ.get("VERCEL") else "sp_aguas.xlsx"))
SQLITE_LEGADO = os.environ.get("SQLITE_DATABASE", "sp_aguas.db")
ARQUIVO_LOCK = RLock()
CONEXAO_COMPARTILHADA = None
ONEDRIVE_ENABLED = os.environ.get("SHAREPOINT_ENABLED", os.environ.get("ONEDRIVE_ENABLED", "")).lower() in {"1", "true", "sim", "yes"}
ONEDRIVE_TENANT_ID = os.environ.get("SHAREPOINT_TENANT_ID", os.environ.get("ONEDRIVE_TENANT_ID", ""))
ONEDRIVE_CLIENT_ID = os.environ.get("SHAREPOINT_CLIENT_ID", os.environ.get("ONEDRIVE_CLIENT_ID", ""))
ONEDRIVE_CLIENT_SECRET = os.environ.get("SHAREPOINT_CLIENT_SECRET", os.environ.get("ONEDRIVE_CLIENT_SECRET", ""))
ONEDRIVE_USER = os.environ.get("SHAREPOINT_USER", os.environ.get("ONEDRIVE_USER", ""))
SHAREPOINT_FILE_PATH = os.environ.get(
	"SHAREPOINT_FILE_PATH",
	"SP_AGUAS/Sistema de Gestão de Demandas - SP Aguas.xlsx"
)
ONEDRIVE_PATH = os.environ.get("ONEDRIVE_PATH", SHAREPOINT_FILE_PATH)
SHAREPOINT_ENABLED = ONEDRIVE_ENABLED
SHAREPOINT_TENANT_ID = ONEDRIVE_TENANT_ID
SHAREPOINT_CLIENT_ID = ONEDRIVE_CLIENT_ID
SHAREPOINT_CLIENT_SECRET = ONEDRIVE_CLIENT_SECRET
SHAREPOINT_SITE_ID = os.environ.get("SHAREPOINT_SITE_ID", "")
SHAREPOINT_DRIVE_ID = os.environ.get("SHAREPOINT_DRIVE_ID", "")

TABELAS_EXCEL = {
	"usuarios": ("id", "nome", "usuario", "email", "senha_hash", "perfil", "ativo", "aprovado", "criado_em"),
	"demandas": ("id", "numero_processo", "origem", "assunto", "area", "responsavel", "data_recebimento", "prazo_area", "prazo_fatal", "situacao", "prioridade", "observacoes", "criado_por", "criado_em", "atualizado_em"),
	"historico": ("id", "demanda_id", "usuario_id", "acao", "descricao", "data_hora"),
}


def _onedrive_configurado():
	# Suporta tanto SharePoint (site/drive) quanto OneDrive for Business do usuário.
	return SHAREPOINT_ENABLED and all((SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET, SHAREPOINT_FILE_PATH))


def _onedrive_token():
	if not _onedrive_configurado():
		raise RuntimeError("SharePoint não configurado. Defina SHAREPOINT_ENABLED, SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET, SHAREPOINT_USER e SHAREPOINT_FILE_PATH.")
	dados = urlencode({
		"client_id": SHAREPOINT_CLIENT_ID,
		"client_secret": SHAREPOINT_CLIENT_SECRET,
		"scope": "https://graph.microsoft.com/.default",
		"grant_type": "client_credentials",
	}).encode()
	url = f"https://login.microsoftonline.com/{quote(SHAREPOINT_TENANT_ID, safe='')}/oauth2/v2.0/token"
	try:
		with urlopen(Request(url, data=dados, headers={"Content-Type": "application/x-www-form-urlencoded"}), timeout=30) as resposta:
			payload = json.loads(resposta.read().decode())
			if not payload.get("access_token"):
				raise RuntimeError("Microsoft Graph nao retornou access_token.")
			return payload["access_token"]
	except HTTPError as erro:
		detalhe = erro.read().decode("utf-8", errors="replace")[:1000]
		raise RuntimeError(f"Falha na autenticacao Microsoft Graph (HTTP {erro.code}): {detalhe}") from erro
	except (URLError, KeyError, json.JSONDecodeError) as erro:
		raise RuntimeError(f"Nao foi possivel autenticar no Microsoft Graph: {erro}") from erro


def _onedrive_url():
	caminho = quote(SHAREPOINT_FILE_PATH.strip("/"), safe="/")
	# Se SITE_ID/DRIVE_ID estiverem definidos, mantém compatibilidade com SharePoint.
	if SHAREPOINT_SITE_ID and SHAREPOINT_DRIVE_ID:
		return f"https://graph.microsoft.com/v1.0/sites/{quote(SHAREPOINT_SITE_ID, safe='')}/drives/{quote(SHAREPOINT_DRIVE_ID, safe='')}/root:/{caminho}:/content"
	if not ONEDRIVE_USER:
		raise RuntimeError("SHAREPOINT_USER/ONEDRIVE_USER é obrigatório quando SITE_ID e DRIVE_ID não forem usados.")
	usuario = quote(ONEDRIVE_USER, safe="")
	return f"https://graph.microsoft.com/v1.0/users/{usuario}/drive/root:/{caminho}:/content"


def obter_planilha():
	"""Baixa a planilha remota para um arquivo temporario e retorna seu caminho."""
	if not _onedrive_configurado():
		return None
	try:
		with urlopen(Request(_onedrive_url(), headers={"Authorization": f"Bearer {_onedrive_token()}"}), timeout=60) as resposta:
			conteudo = resposta.read()
	except HTTPError as erro:
		detalhe = erro.read().decode("utf-8", errors="replace")[:1500]
		if erro.code == 404:
			raise RuntimeError(f"Arquivo nao encontrado no SharePoint. Verifique SHAREPOINT_FILE_PATH='{SHAREPOINT_FILE_PATH}'. Resposta Graph: {detalhe}") from erro
		if erro.code in (401, 403):
			raise RuntimeError("Microsoft Graph recusou o acesso. Verifique o consentimento administrativo e Files.ReadWrite.All.") from erro
		raise RuntimeError(f"Nao foi possivel baixar a planilha do SharePoint (HTTP {erro.code}): {detalhe}") from erro
	except URLError as erro:
		raise RuntimeError(f"Nao foi possivel acessar o SharePoint: {erro}") from erro
	with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temporario:
		temporario.write(conteudo)
		return temporario.name


def _baixar_planilha_one_drive():
	caminho_temporario = obter_planilha()
	if not caminho_temporario:
		return False
	os.makedirs(os.path.dirname(os.path.abspath(PLANILHA)), exist_ok=True)
	try:
		os.replace(caminho_temporario, PLANILHA)
	finally:
		if os.path.exists(caminho_temporario):
			os.unlink(caminho_temporario)
	return True


def salvar_planilha():
	"""Impedida de fazer PUT cego; toda escrita deve usar eTag/If-Match."""
	raise RuntimeError("Gravação direta desabilitada. Use executar_mutacao_atomica().")


def _enviar_planilha_one_drive():
	return salvar_planilha()


def _graph_metadata_diagnostico():
	"""Obtém os metadados do arquivo remoto no SharePoint."""
	token = _onedrive_token()
	url = _onedrive_url().replace("/content", "")
	requisicao = Request(url, headers={"Authorization": f"Bearer {token}"})
	with urlopen(requisicao, timeout=30) as resposta:
		return json.loads(resposta.read().decode("utf-8"))


def _graph_download_diagnostico():
	"""Baixa o arquivo remoto como bytes."""
	token = _onedrive_token()
	requisicao = Request(_onedrive_url(), headers={"Authorization": f"Bearer {token}"})
	with urlopen(requisicao, timeout=60) as resposta:
		return resposta.read()


def _graph_upload_diagnostico(conteudo, etag=None):
	"""PUT condicional no arquivo remoto; If-Match impede sobrescrita de versão concorrente."""
	token = _onedrive_token()
	headers = {
		"Authorization": f"Bearer {token}",
		"Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
	}
	if etag:
		headers["If-Match"] = etag
	requisicao = Request(
		_onedrive_url(),
		data=conteudo,
		method="PUT",
		headers=headers,
	)
	try:
		with urlopen(requisicao, timeout=120) as resposta:
			return resposta.status
	except HTTPError as erro:
		if erro.code == 412:
			raise ConcurrentUpdateError("O arquivo do SharePoint foi alterado por outra instância durante a gravação.") from erro
		if erro.code in (401, 403):
			raise RuntimeError("Microsoft Graph recusou a gravação. Verifique o consentimento administrativo e Files.ReadWrite.All.") from erro
		if erro.code == 404:
			raise RuntimeError(f"Arquivo não localizado para gravação. Verifique SHAREPOINT_FILE_PATH='{SHAREPOINT_FILE_PATH}'.") from erro
		detalhe = erro.read().decode("utf-8", errors="replace")[:1200]
		raise RuntimeError(f"Falha SharePoint HTTP {erro.code}: {detalhe}") from erro


class ConcurrentUpdateError(RuntimeError):
	pass


def _graph_snapshot():
	"""Lê metadados e bytes do mesmo arquivo remoto."""
	meta = _graph_metadata_diagnostico()
	etag = meta.get("eTag")
	if not etag:
		raise RuntimeError("O SharePoint não retornou o eTag do arquivo remoto.")
	conteudo = _graph_download_diagnostico()
	if not conteudo:
		raise RuntimeError("O arquivo remoto está vazio.")
	return meta, conteudo


def _gerar_xlsx_com_demanda(conteudo_remoto, demanda, historico):
	"""Atualiza o XLSX remoto preservando todas as abas e linhas existentes."""
	workbook = load_workbook(io.BytesIO(conteudo_remoto))
	try:
		if "demandas" not in workbook.sheetnames or "historico" not in workbook.sheetnames:
			raise RuntimeError("O arquivo remoto precisa conter as abas demandas e historico.")
		planilha = workbook["demandas"]
		cabecalho = [c.value for c in planilha[1]]
		indices = {nome: cabecalho.index(nome) + 1 for nome in TABELAS_EXCEL["demandas"] if nome in cabecalho}
		if len(indices) != len(TABELAS_EXCEL["demandas"]):
			raise RuntimeError("A aba demandas não possui todas as colunas esperadas.")
		planilha.append([demanda[coluna] for coluna in TABELAS_EXCEL["demandas"]])
		planilha_h = workbook["historico"]
		cab_h = [c.value for c in planilha_h[1]]
		if not all(c in cab_h for c in TABELAS_EXCEL["historico"]):
			raise RuntimeError("A aba historico não possui todas as colunas esperadas.")
		planilha_h.append([historico[coluna] for coluna in TABELAS_EXCEL["historico"]])
		buffer = io.BytesIO()
		workbook.save(buffer)
		return buffer.getvalue()
	finally:
		workbook.close()


def _confirmar_demanda_no_xlsx(conteudo, demanda_id):
	workbook = load_workbook(io.BytesIO(conteudo), read_only=True, data_only=True)
	try:
		if "demandas" not in workbook.sheetnames:
			return False
		planilha = workbook["demandas"]
		cabecalho = [c.value for c in next(planilha.iter_rows(min_row=1, max_row=1))]
		if "id" not in cabecalho:
			raise RuntimeError("A aba demandas não possui a coluna id.")
		idx = cabecalho.index("id")
		return any(linha[idx] == demanda_id for linha in planilha.iter_rows(min_row=2, values_only=True))
	finally:
		workbook.close()


def salvar_demanda_atomicamente(valores, usuario_id, tentativas=3):
	"""Cria uma demanda com controle otimista de concorrência no SharePoint."""
	if not _onedrive_configurado():
		raise RuntimeError("Integração SharePoint não está configurada.")
	for tentativa in range(1, tentativas + 1):
		with ARQUIVO_LOCK:
			meta, remoto_antes = _graph_snapshot()
			etag = meta["eTag"]
			conn = conectar()
			conn._conn.rollback()
			for tabela in reversed(tuple(TABELAS_EXCEL.keys())):
				conn._conn.execute(f"DELETE FROM {tabela}")
			_carregar_xlsx_bytes_na_conexao(conn, remoto_antes)
			agora_valor = agora()
			cur = conn._conn.execute(
				"INSERT INTO demandas (numero_processo, origem, assunto, area, responsavel, data_recebimento, prazo_area, prazo_fatal, situacao, prioridade, observacoes, criado_por, criado_em, atualizado_em) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
				(*valores, usuario_id, agora_valor, agora_valor),
			)
			demanda_id = cur.lastrowid
			historico_id = conn._conn.execute("SELECT COALESCE(MAX(id), 0) + 1 FROM historico").fetchone()[0]
			historico = {"id": historico_id, "demanda_id": demanda_id, "usuario_id": usuario_id, "acao": "CRIACAO", "descricao": "Demanda cadastrada.", "data_hora": agora_valor}
			conn._conn.execute(
				"INSERT INTO historico (id, demanda_id, usuario_id, acao, descricao, data_hora) VALUES (?, ?, ?, ?, ?, ?)",
				tuple(historico[c] for c in TABELAS_EXCEL["historico"]),
			)
			conn._conn.commit()
			demanda = {c: conn._conn.execute(f"SELECT {c} FROM demandas WHERE id = ?", (demanda_id,)).fetchone()[0] for c in TABELAS_EXCEL["demandas"]}
			novo_xlsx = _gerar_xlsx_com_demanda(remoto_antes, demanda, historico)
			hash_enviado = hashlib.sha256(novo_xlsx).hexdigest()
			try:
				status_put = _graph_upload_diagnostico(novo_xlsx, etag=etag)
			except ConcurrentUpdateError:
				if tentativa == tentativas:
					raise RuntimeError("Não foi possível gravar porque o Excel foi alterado por outra instância. Tente novamente.")
				continue
			if status_put not in (200, 201):
				raise RuntimeError(f"Microsoft Graph retornou HTTP {status_put}.")
			remoto_depois = _graph_download_diagnostico()
			hash_remoto = hashlib.sha256(remoto_depois).hexdigest()
			if hash_remoto != hash_enviado:
				raise RuntimeError("O arquivo remoto ficou diferente do XLSX enviado.")
			if not _confirmar_demanda_no_xlsx(remoto_depois, demanda_id):
				raise RuntimeError("A demanda não foi encontrada no arquivo remoto após a gravação.")
			return {"status": "OK", "demanda_id": demanda_id, "tentativa": tentativa, "etag_antes": etag, "sha256": hash_remoto, "mensagem": "Demanda cadastrada e confirmada no SharePoint."}
	raise RuntimeError("Falha de concorrência ao gravar a demanda.")


def _carregar_xlsx_bytes_na_conexao(conn, conteudo):
	"""Carrega um snapshot XLSX em uma conexão SQLite em memória já existente."""
	workbook = load_workbook(io.BytesIO(conteudo), read_only=True, data_only=True)
	try:
		for tabela, colunas in TABELAS_EXCEL.items():
			if tabela not in workbook.sheetnames:
				continue
			planilha = workbook[tabela]
			cabecalho = [c.value for c in next(planilha.iter_rows(min_row=1, max_row=1))]
			indices = {nome: cabecalho.index(nome) for nome in colunas if nome in cabecalho}
			if len(indices) != len(colunas):
				continue
			for linha in planilha.iter_rows(min_row=2, values_only=True):
				if not any(v is not None for v in linha):
					continue
				valores_linha = tuple(linha[indices[c]] for c in colunas)
				conn._conn.execute(f"INSERT INTO {tabela} ({', '.join(colunas)}) VALUES ({', '.join('?' for _ in colunas)})", valores_linha)
	finally:
		workbook.close()


def salvar_demanda_com_confirmacao(demanda):
	"""Sincroniza uma demanda e confirma sua presença no arquivo remoto."""
	arquivo_remoto = _graph_download_diagnostico()
	metadata = _graph_metadata_diagnostico()
	etag = metadata.get("eTag")
	if not etag:
		raise RuntimeError("O SharePoint não retornou o eTag do arquivo remoto.")
	workbook = load_workbook(io.BytesIO(arquivo_remoto))
	if "demandas" not in workbook.sheetnames:
		workbook.close()
		raise RuntimeError("O arquivo remoto não possui a aba demandas.")
	planilha = workbook["demandas"]
	nova_linha = [demanda[coluna] for coluna in TABELAS_EXCEL["demandas"]]
	cabecalho = [celula.value for celula in planilha[1]]
	try:
		indice_id = cabecalho.index("id") + 1
	except ValueError as erro:
		workbook.close()
		raise RuntimeError("A aba demandas não possui a coluna id.") from erro

	linha_existente = None
	for numero_linha in range(2, planilha.max_row + 1):
		if planilha.cell(numero_linha, indice_id).value == demanda["id"]:
			linha_existente = numero_linha
			break
	if linha_existente is None:
		planilha.append(nova_linha)
	else:
		for numero_coluna, valor in enumerate(nova_linha, 1):
			planilha.cell(linha_existente, numero_coluna).value = valor
	buffer = io.BytesIO()
	workbook.save(buffer)
	workbook.close()
	novo_xlsx = buffer.getvalue()
	novo_hash = hashlib.sha256(novo_xlsx).hexdigest()
	status = _graph_upload_diagnostico(novo_xlsx, etag=etag)
	if status not in (200, 201):
		raise RuntimeError(f"Falha ao gravar no SharePoint. HTTP={status}")

	remoto_confirmacao = _graph_download_diagnostico()
	workbook_confirmacao = load_workbook(io.BytesIO(remoto_confirmacao), read_only=True, data_only=True)
	try:
		if "demandas" not in workbook_confirmacao.sheetnames:
			raise RuntimeError("Arquivo remoto não possui a aba demandas após a gravação.")
		indice_id_confirmacao = [celula.value for celula in next(workbook_confirmacao["demandas"].iter_rows(min_row=1, max_row=1))].index("id")
		encontrada = any(
			linha[indice_id_confirmacao] == demanda["id"]
			for linha in workbook_confirmacao["demandas"].iter_rows(min_row=2, values_only=True)
		)
	finally:
		workbook_confirmacao.close()
	if not encontrada:
		raise RuntimeError("Arquivo remoto não contém a demanda recém gravada.")
	return {"status": "OK", "hash_enviado": novo_hash, "hash_remoto": hashlib.sha256(remoto_confirmacao).hexdigest(), "mensagem": "Demanda cadastrada e confirmada no SharePoint."}


def salvar_demanda_sharepoint_seguro(demanda_id):
	"""Atualiza uma demanda remota com eTag e confirma a persistência."""
	metadata = _graph_metadata_diagnostico()
	etag = metadata.get("eTag")
	if not etag:
		raise RuntimeError("SharePoint não retornou eTag.")

	arquivo_remoto = _graph_download_diagnostico()
	workbook = load_workbook(io.BytesIO(arquivo_remoto))
	try:
		if "demandas" not in workbook.sheetnames:
			raise RuntimeError("Aba demandas não encontrada.")

		conn = conectar()
		registro = conn.execute("SELECT * FROM demandas WHERE id = ?", (demanda_id,)).fetchone()
		conn.close()
		if not registro:
			raise RuntimeError("Demanda não encontrada.")

		planilha = workbook["demandas"]
		cabecalho = [celula.value for celula in planilha[1]]
		try:
			indice_id = cabecalho.index("id") + 1
		except ValueError as erro:
			raise RuntimeError("A aba demandas não possui a coluna id.") from erro

		valores = [registro[coluna] for coluna in TABELAS_EXCEL["demandas"]]
		linha_localizada = next((linha for linha in range(2, planilha.max_row + 1) if planilha.cell(linha, indice_id).value == registro["id"]), None)
		if linha_localizada:
			for coluna, valor in enumerate(valores, start=1):
				planilha.cell(linha_localizada, coluna).value = valor
		else:
			planilha.append(valores)

		buffer = io.BytesIO()
		workbook.save(buffer)
		bytes_xlsx = buffer.getvalue()
	finally:
		workbook.close()

	hash_enviado = hashlib.sha256(bytes_xlsx).hexdigest()
	status = _graph_upload_diagnostico(bytes_xlsx, etag)
	if status not in (200, 201):
		raise RuntimeError("Falha ao gravar no SharePoint.")

	remoto = _graph_download_diagnostico()
	hash_remoto = hashlib.sha256(remoto).hexdigest()
	if hash_enviado != hash_remoto:
		raise RuntimeError("Confirmação SHA-256 falhou.")

	workbook_confirmacao = load_workbook(io.BytesIO(remoto), read_only=True, data_only=True)
	try:
		planilha_confirmacao = workbook_confirmacao["demandas"]
		cabecalho_confirmacao = [celula.value for celula in next(planilha_confirmacao.iter_rows(min_row=1, max_row=1))]
		indice_confirmacao = cabecalho_confirmacao.index("id")
		encontrada = any(linha[indice_confirmacao] == demanda_id for linha in planilha_confirmacao.iter_rows(min_row=2, values_only=True))
	finally:
		workbook_confirmacao.close()
	if not encontrada:
		raise RuntimeError("Demanda não localizada no arquivo remoto.")
	return "Demanda cadastrada"


def _criar_esquema(conn):
	conn.executescript("""
		CREATE TABLE IF NOT EXISTS usuarios (
			id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT NOT NULL,
			usuario TEXT NOT NULL UNIQUE, email TEXT UNIQUE, senha_hash TEXT NOT NULL,
			perfil TEXT NOT NULL DEFAULT 'Usuario', ativo INTEGER NOT NULL DEFAULT 1,
			aprovado INTEGER NOT NULL DEFAULT 0, criado_em TEXT NOT NULL
		);
		CREATE TABLE IF NOT EXISTS demandas (
			id INTEGER PRIMARY KEY AUTOINCREMENT, numero_processo TEXT, origem TEXT NOT NULL,
			assunto TEXT NOT NULL, area TEXT, responsavel TEXT, data_recebimento TEXT,
			prazo_area TEXT, prazo_fatal TEXT, situacao TEXT, prioridade TEXT, observacoes TEXT,
			criado_por INTEGER, criado_em TEXT NOT NULL, atualizado_em TEXT
		);
		CREATE TABLE IF NOT EXISTS historico (
			id INTEGER PRIMARY KEY AUTOINCREMENT, demanda_id INTEGER, usuario_id INTEGER,
			acao TEXT, descricao TEXT, data_hora TEXT NOT NULL
		);
	""")


def _carregar_planilha(conn):
	if not os.path.exists(PLANILHA):
		return False
	workbook = load_workbook(PLANILHA, read_only=True, data_only=True)
	try:
		for tabela, colunas in TABELAS_EXCEL.items():
			if tabela not in workbook.sheetnames:
				continue
			planilha = workbook[tabela]
			cabecalho = [celula.value for celula in next(planilha.iter_rows(min_row=1, max_row=1))]
			indices = {nome: cabecalho.index(nome) for nome in colunas if nome in cabecalho}
			if len(indices) != len(colunas):
				continue
			for linha in planilha.iter_rows(min_row=2, values_only=True):
				if not any(valor is not None for valor in linha):
					continue
				valores = tuple(linha[indices[coluna]] for coluna in colunas)
				conn.execute(f"INSERT INTO {tabela} ({', '.join(colunas)}) VALUES ({', '.join('?' for _ in colunas)})", valores)
	finally:
		workbook.close()
	return True


def _carregar_sqlite_legado(conn):
	if not os.path.exists(SQLITE_LEGADO) or os.path.abspath(SQLITE_LEGADO) == os.path.abspath(PLANILHA):
		return
	legado = sqlite3.connect(SQLITE_LEGADO)
	legado.row_factory = sqlite3.Row
	try:
		for tabela, colunas in TABELAS_EXCEL.items():
			try:
				registros = legado.execute(f"SELECT {', '.join(colunas)} FROM {tabela}").fetchall()
			except sqlite3.OperationalError:
				continue
			for registro in registros:
				conn.execute(f"INSERT INTO {tabela} ({', '.join(colunas)}) VALUES ({', '.join('?' for _ in colunas)})", tuple(registro))
	finally:
		legado.close()


def _salvar_planilha(conn):
	workbook = Workbook()
	workbook.remove(workbook.active)
	for tabela, colunas in TABELAS_EXCEL.items():
		planilha = workbook.create_sheet(tabela)
		planilha.append(list(colunas))
		for celula in planilha[1]:
			celula.font = Font(bold=True)
		for registro in conn.execute(f"SELECT {', '.join(colunas)} FROM {tabela} ORDER BY id"):
			planilha.append([registro[coluna] for coluna in colunas])
		planilha.freeze_panes = "A2"
		planilha.auto_filter.ref = planilha.dimensions
		for coluna in planilha.columns:
			largura = min(max(len(str(celula.value or "")) for celula in coluna) + 2, 45)
			planilha.column_dimensions[coluna[0].column_letter].width = largura
	os.makedirs(os.path.dirname(os.path.abspath(PLANILHA)), exist_ok=True)
	diretorio = os.path.dirname(os.path.abspath(PLANILHA))
	with tempfile.NamedTemporaryFile(suffix=".xlsx", dir=diretorio, delete=False) as temporario:
		caminho_temporario = temporario.name
	try:
		workbook.save(caminho_temporario)
		os.replace(caminho_temporario, PLANILHA)
	finally:
		if os.path.exists(caminho_temporario):
			os.unlink(caminho_temporario)
	_enviar_planilha_one_drive()


class ConexaoExcel:
	def __init__(self):
		self._conn = sqlite3.connect(":memory:", check_same_thread=False)
		self._conn.row_factory = sqlite3.Row
		_baixar_planilha_one_drive()
		_criar_esquema(self._conn)
		if not _carregar_planilha(self._conn):
			_carregar_sqlite_legado(self._conn)

	def execute(self, consulta, parametros=()):
		return self._conn.execute(consulta, parametros)

	def executescript(self, consulta):
		return self._conn.executescript(consulta)

	def commit(self):
		# O XLSX remoto só é alterado pelos fluxos atômicos com eTag/If-Match.
		with ARQUIVO_LOCK:
			self._conn.commit()

	def close(self):
		pass


def conectar():
	global CONEXAO_COMPARTILHADA
	with ARQUIVO_LOCK:
		if CONEXAO_COMPARTILHADA is None:
			CONEXAO_COMPARTILHADA = ConexaoExcel()
		return CONEXAO_COMPARTILHADA


def sincronizar_planilha():
	"""Mantido por compatibilidade, mas sem PUT cego do banco em memória."""
	return executar_mutacao_atomica(lambda workbook: {"operacao": "SINCRONIZACAO"})


def agora():
	return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _cabecalho_planilha(planilha):
	return [celula.value for celula in next(planilha.iter_rows(min_row=1, max_row=1))]


def _linha_por_id(planilha, cabecalho, registro_id):
	if "id" not in cabecalho:
		raise RuntimeError(f"A aba {planilha.title} não possui a coluna id.")
	indice = cabecalho.index("id") + 1
	for numero_linha in range(2, planilha.max_row + 1):
		if planilha.cell(numero_linha, indice).value == registro_id:
			return numero_linha
	return None


def _proximo_id(planilha, cabecalho):
	if "id" not in cabecalho:
		raise RuntimeError(f"A aba {planilha.title} não possui a coluna id.")
	indice = cabecalho.index("id") + 1
	maior = 0
	for numero_linha in range(2, planilha.max_row + 1):
		valor = planilha.cell(numero_linha, indice).value
		try:
			maior = max(maior, int(valor))
		except (TypeError, ValueError):
			continue
	return maior + 1


def _salvar_workbook_preservando_arquivo(workbook):
	buffer = io.BytesIO()
	workbook.save(buffer)
	return buffer.getvalue()


def _atualizar_conexao_com_snapshot(conteudo):
	"""Depois de uma gravação confirmada, faz a conexão em memória refletir o XLSX remoto."""
	global CONEXAO_COMPARTILHADA
	with ARQUIVO_LOCK:
		if CONEXAO_COMPARTILHADA is None:
			CONEXAO_COMPARTILHADA = ConexaoExcel()
		conn = CONEXAO_COMPARTILHADA
		conn._conn.rollback()
		for tabela in reversed(tuple(TABELAS_EXCEL.keys())):
			conn._conn.execute(f"DELETE FROM {tabela}")
		_carregar_xlsx_bytes_na_conexao(conn, conteudo)
		conn._conn.commit()
	return conn


def executar_mutacao_atomica(mutator, confirmador=None, tentativas=3):
	"""Executa qualquer escrita do sistema contra a versão atual do Excel."""
	if not _onedrive_configurado():
		raise RuntimeError("Integração SharePoint não está configurada.")
	ultimo_conflito = None
	for tentativa in range(1, tentativas + 1):
		with ARQUIVO_LOCK:
			meta, remoto_antes = _graph_snapshot()
			etag = meta["eTag"]
			workbook = load_workbook(io.BytesIO(remoto_antes))
			try:
				resultado = mutator(workbook)
				novo_xlsx = _salvar_workbook_preservando_arquivo(workbook)
			finally:
				workbook.close()
			try:
				status_put = _graph_upload_diagnostico(novo_xlsx, etag=etag)
			except ConcurrentUpdateError as erro:
				ultimo_conflito = erro
				continue
			if status_put not in (200, 201):
				raise RuntimeError(f"Microsoft Graph retornou HTTP {status_put}.")
			remoto_depois = _graph_download_diagnostico()
			if hashlib.sha256(remoto_depois).hexdigest() != hashlib.sha256(novo_xlsx).hexdigest():
				raise RuntimeError("O arquivo remoto ficou diferente do XLSX enviado; a operação não foi confirmada.")
			if confirmador is not None and not confirmador(remoto_depois, resultado):
				raise RuntimeError("A operação foi enviada, mas não pôde ser confirmada no arquivo remoto.")
			_atualizar_conexao_com_snapshot(remoto_depois)
			return resultado
	if ultimo_conflito:
		raise RuntimeError("O Excel foi alterado simultaneamente por outra instância. Tente novamente.") from ultimo_conflito
	raise RuntimeError("Não foi possível concluir a gravação no SharePoint.")


def _append_registro(planilha, colunas, registro):
	cabecalho = _cabecalho_planilha(planilha)
	if not all(c in cabecalho for c in colunas):
		raise RuntimeError(f"A aba {planilha.title} não possui todas as colunas esperadas.")
	linha = planilha.max_row + 1
	for coluna, nome in enumerate(cabecalho, start=1):
		if nome in colunas:
			planilha.cell(linha, coluna).value = registro[nome]
	return linha


def _atualizar_registro(planilha, colunas, registro_id, registro):
	cabecalho = _cabecalho_planilha(planilha)
	linha = _linha_por_id(planilha, cabecalho, registro_id)
	if linha is None:
		raise RuntimeError(f"Registro {registro_id} não encontrado na aba {planilha.title}.")
	for coluna, nome in enumerate(cabecalho, start=1):
		if nome in colunas:
			planilha.cell(linha, coluna).value = registro[nome]
	return linha


def _confirmar_id_na_aba(conteudo, aba, registro_id):
	workbook = load_workbook(io.BytesIO(conteudo), read_only=True, data_only=True)
	try:
		if aba not in workbook.sheetnames:
			return False
		planilha = workbook[aba]
		cabecalho = _cabecalho_planilha(planilha)
		if "id" not in cabecalho:
			return False
		idx = cabecalho.index("id")
		return any(linha[idx] == registro_id for linha in planilha.iter_rows(min_row=2, values_only=True))
	finally:
		workbook.close()


def criar_banco():
	conn = conectar()
	conn.executescript("""
		CREATE TABLE IF NOT EXISTS usuarios (
			id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT NOT NULL,
			usuario TEXT NOT NULL UNIQUE, email TEXT UNIQUE, senha_hash TEXT NOT NULL,
			perfil TEXT NOT NULL DEFAULT 'Usuario', ativo INTEGER NOT NULL DEFAULT 1,
			aprovado INTEGER NOT NULL DEFAULT 0, criado_em TEXT NOT NULL
		);
		CREATE TABLE IF NOT EXISTS demandas (
			id INTEGER PRIMARY KEY AUTOINCREMENT, numero_processo TEXT, origem TEXT NOT NULL,
			assunto TEXT NOT NULL, area TEXT, responsavel TEXT, data_recebimento TEXT,
			prazo_area TEXT, prazo_fatal TEXT, situacao TEXT, prioridade TEXT, observacoes TEXT,
			criado_por INTEGER, criado_em TEXT NOT NULL, atualizado_em TEXT
		);
		CREATE TABLE IF NOT EXISTS historico (
			id INTEGER PRIMARY KEY AUTOINCREMENT, demanda_id INTEGER, usuario_id INTEGER,
			acao TEXT, descricao TEXT, data_hora TEXT NOT NULL
		);
	""")
	colunas_demandas = {linha[1] for linha in conn.execute("PRAGMA table_info(demandas)")}
	for coluna, definicao in (("criado_por", "INTEGER"), ("atualizado_em", "TEXT")):
		if coluna not in colunas_demandas:
			conn.execute(f"ALTER TABLE demandas ADD COLUMN {coluna} {definicao}")
	conn.commit()
	conn.close()


def data_br(valor):
	if not valor:
		return "-"
	try:
		return datetime.strptime(valor, "%Y-%m-%d").strftime("%d/%m/%Y")
	except (TypeError, ValueError):
		return valor


def calcular_prazos(prazo_area, prazo_fatal):
	resultado = {"dias_entre": None, "dias_area": None, "dias_fatal": None, "status": "Sem prazo"}
	try:
		fatal = datetime.strptime(prazo_fatal, "%Y-%m-%d").date()
		resultado["dias_fatal"] = (fatal - date.today()).days
	except (TypeError, ValueError):
		pass
	try:
		area = datetime.strptime(prazo_area, "%Y-%m-%d").date()
		resultado["dias_area"] = (area - date.today()).days
	except (TypeError, ValueError):
		pass
	try:
		area = datetime.strptime(prazo_area, "%Y-%m-%d").date()
		fatal = datetime.strptime(prazo_fatal, "%Y-%m-%d").date()
		resultado["dias_entre"] = (fatal - area).days
	except (TypeError, ValueError):
		pass
	dias = resultado["dias_fatal"]
	if dias is not None:
		resultado["status"] = "VENCIDO" if dias < 0 else "CRÍTICO" if dias <= 3 else "PRÓXIMO" if dias <= 7 else "NORMAL"
	return resultado


def registrar_historico(demanda_id, usuario_id, acao, descricao):
	"""Registra histórico no mesmo ciclo atômico do Excel remoto."""
	def mutator(workbook):
		planilha = workbook["historico"]
		cabecalho = _cabecalho_planilha(planilha)
		historico_id = _proximo_id(planilha, cabecalho)
		registro = {"id": historico_id, "demanda_id": demanda_id, "usuario_id": usuario_id, "acao": acao, "descricao": descricao, "data_hora": agora()}
		_append_registro(planilha, TABELAS_EXCEL["historico"], registro)
		return registro
	return executar_mutacao_atomica(mutator, lambda conteudo, r: _confirmar_id_na_aba(conteudo, "historico", r["id"]))


def usuario_logado():
	return "usuario_id" in session


def administrador():
	return session.get("perfil") == "Administrador"


def popup_alertas():
	if not usuario_logado():
		return ""
	conn = conectar()
	registros = conn.execute("SELECT * FROM demandas WHERE situacao != 'Concluído' ORDER BY prazo_fatal").fetchall()
	conn.close()
	alertas = []
	for demanda in registros:
		prazo = calcular_prazos(demanda["prazo_area"], demanda["prazo_fatal"])
		if prazo["status"] != "NORMAL" and prazo["dias_fatal"] is not None:
			alertas.append((demanda, prazo))
	if not alertas:
		return ""
	linhas = "".join(
		f'<li><strong>{html.escape(prazo["status"])}</strong> — '
		f'{html.escape(demanda["numero_processo"] or "Sem número")} — '
		f'{html.escape(demanda["assunto"])} '
		f'<span>({prazo["dias_fatal"]} dias restantes)</span></li>'
		for demanda, prazo in alertas[:5]
	)
	mais = f'<p class="popup-mais">E mais {len(alertas) - 5} alerta(s).</p>' if len(alertas) > 5 else ""
	return f'''
	<div id="popup-alertas" class="popup-alertas" role="alertdialog" aria-modal="true" aria-labelledby="popup-alertas-titulo">
		<div class="popup-alertas-conteudo">
			<button type="button" class="popup-fechar" aria-label="Fechar alertas" onclick="fecharPopupAlertas()">&times;</button>
			<div class="popup-icone">!</div>
			<h2 id="popup-alertas-titulo">Atenção aos prazos</h2>
			<p>Existem <strong>{len(alertas)}</strong> demanda(s) que precisam de acompanhamento.</p>
			<ul>{linhas}</ul>
			{mais}
			<a class="btn" href="/alertas">Ver todos os alertas</a>
		</div>
	</div>
	<script>
		(function () {{
			const chave = "sp-alertas-" + new Date().toISOString().slice(0, 10);
			if (!sessionStorage.getItem(chave)) document.getElementById("popup-alertas").classList.add("visivel");
		}})();
		function fecharPopupAlertas() {{
			document.getElementById("popup-alertas").classList.remove("visivel");
			sessionStorage.setItem("sp-alertas-" + new Date().toISOString().slice(0, 10), "1");
		}}
	</script>'''


HTML_BASE = """
<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>{{ titulo or 'SP ÁGUAS' }}</title><style>
*{box-sizing:border-box}body{margin:0;font-family:Arial,sans-serif;background:#f3f6f9;color:#263238}header{background:linear-gradient(135deg,#005b96,#0077b6);color:#fff;padding:18px 30px;box-shadow:0 2px 8px #0003}.logo{font-size:22px;font-weight:bold}.menu{margin-top:15px}.menu a{color:#fff;text-decoration:none;margin-right:18px;font-size:14px}.container{max-width:1400px;margin:auto;padding:25px}.card,.metrica{background:#fff;border-radius:12px;padding:22px;margin-bottom:22px;box-shadow:0 2px 8px #0001}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:15px}.metrica strong{display:block;font-size:30px;margin-top:10px;color:#005b96}.metrica h3{margin:0;color:#607d8b;font-size:14px}input,select,textarea{width:100%;padding:11px;border:1px solid #cfd8dc;border-radius:6px;margin:6px 0 15px;font-size:14px}textarea{min-height:100px}label{font-weight:bold;font-size:13px}button,.btn{background:#0077b6;color:#fff;border:0;border-radius:6px;padding:10px 16px;cursor:pointer;text-decoration:none;display:inline-block}.btn-verde{background:#2e7d32}.btn-vermelho{background:#c62828}.btn-cinza{background:#607d8b}table{width:100%;border-collapse:collapse;background:#fff}th{background:#005b96;color:#fff;padding:11px;text-align:left}td{padding:10px;border-bottom:1px solid #e0e0e0;font-size:13px}.alerta{padding:15px;border-radius:8px;margin-bottom:10px}.vencido{background:#ffebee;color:#b71c1c}.critico{background:#fff3e0;color:#e65100}.proximo{background:#fffde7;color:#827717}.normal{background:#e8f5e9;color:#1b5e20}.info{background:#e3f2fd;color:#0d47a1}.erro{background:#ffebee;color:#b71c1c;padding:12px;border-radius:6px}@media(max-width:800px){.container{padding:12px}table{display:block;overflow-x:auto}}
</style></head><body><header><div class="logo">SP ÁGUAS</div>{% if session.get('usuario_id') %}<div>Sistema de Gestão de Processos e Prazos</div><div class="menu"><a href="/sistema">Dashboard</a><a href="/demandas">Demandas</a><a href="/nova-demanda">Nova demanda</a><a href="/tce">TCE-SP</a><a href="/audesp">AUDESP</a><a href="/alertas">Alertas</a><a href="/calendario">Calendário</a><a href="/exportar">Exportar</a>{% if session.get('perfil') == 'Administrador' %}<a href="/usuarios">Usuários</a><a href="/status">Status</a>{% endif %}<a href="/logout">Sair</a></div>{% endif %}</header><main class="container">{{ conteudo|safe }}</main></body></html>
"""


def pagina(titulo, conteudo):
	estilo_moderno = """
	<style>
	:root { --azul-escuro: #073b5c; --azul: #087e9f; --fundo: #f4f7f9; --borda: #dce6eb; --texto: #20333f; --suave: #71828d; --sombra: 0 10px 30px rgba(21, 55, 75, .08); }
	body { background: var(--fundo); color: var(--texto); font-family: Inter, "Segoe UI", sans-serif; }
	header { position: fixed; inset: 0 auto 0 0; width: 250px; min-height: 100vh; padding: 28px 16px; background: linear-gradient(180deg, #073b5c 0%, #087e9f 100%); box-shadow: 6px 0 24px rgba(7, 59, 92, .14); z-index: 10; }
	.logo { padding: 8px 12px 28px; font-size: 21px; letter-spacing: .4px; }
	header > div:not(.logo) { padding: 0 12px; color: rgba(255,255,255,.7); font-size: 12px; line-height: 1.5; }
	.menu { display: flex; flex-direction: column; gap: 5px; margin-top: 24px; }
	.menu a { margin: 0; padding: 11px 12px; border-radius: 9px; color: rgba(255,255,255,.86); font-size: 13px; transition: background .2s, transform .2s; }
	.menu a:hover { background: rgba(255,255,255,.14); transform: translateX(3px); }
	.container { max-width: none; min-height: 100vh; margin-left: 250px; padding: 38px 42px; }
	.card, .metrica { border: 1px solid var(--borda); border-radius: 14px; box-shadow: var(--sombra); }
	.card h1 { margin-top: 0; color: var(--azul-escuro); font-size: 27px; }
	.card h2 { color: var(--azul-escuro); font-size: 18px; }
	.grid { gap: 18px; }
	.metrica { padding: 22px; }
	.metrica h3 { text-transform: uppercase; letter-spacing: .6px; font-size: 11px; }
	.metrica strong { color: var(--azul); font-size: 32px; }
	input, select, textarea { border-color: var(--borda); background: #fbfdfe; border-radius: 8px; font-family: inherit; }
	input:focus, select:focus, textarea:focus { outline: 0; border-color: var(--azul); box-shadow: 0 0 0 3px rgba(8,126,159,.14); background: #fff; }
	button, .btn { border-radius: 8px; font-weight: 600; background: var(--azul); transition: transform .2s, box-shadow .2s, background .2s; }
	button:hover, .btn:hover { background: #066b88; box-shadow: 0 5px 12px rgba(8,126,159,.2); transform: translateY(-1px); }
	th { background: #f5f8fa; color: #60727d; font-size: 11px; text-transform: uppercase; letter-spacing: .45px; border-bottom: 1px solid var(--borda); }
	td { padding: 13px 10px; }
	.alerta { border-left: 4px solid transparent; box-shadow: 0 2px 8px rgba(21,55,75,.04); }
	.vencido { border-left-color: #c62828; } .critico { border-left-color: #e66a00; } .proximo { border-left-color: #c18a05; } .normal { border-left-color: #2e7d32; }
	.popup-alertas { display: none; position: fixed; inset: 0; z-index: 30; align-items: center; justify-content: center; padding: 20px; background: rgba(7, 30, 45, .48); }
	.popup-alertas.visivel { display: flex; animation: aparecer .2s ease-out; }
	.popup-alertas-conteudo { position: relative; width: min(520px, 100%); padding: 30px; border: 1px solid #f0d6a6; border-radius: 16px; background: #fffdf8; box-shadow: 0 18px 50px rgba(7, 30, 45, .25); }
	.popup-alertas-conteudo h2 { margin: 0 0 8px; color: #7b3f00; }
	.popup-alertas-conteudo ul { max-height: 230px; margin: 18px 0; padding-left: 20px; color: #4b3b2a; }
	.popup-alertas-conteudo li { margin: 9px 0; }
	.popup-alertas-conteudo li strong { color: #b3261e; }
	.popup-alertas-conteudo li span { color: #786b5c; font-size: 13px; }
	.popup-fechar { position: absolute; top: 10px; right: 12px; padding: 2px 9px; background: transparent; color: #6d6258; font-size: 27px; line-height: 1; }
	.popup-fechar:hover { background: transparent; color: #2d2520; box-shadow: none; transform: none; }
	.popup-icone { display: grid; width: 34px; height: 34px; margin-bottom: 12px; place-items: center; border-radius: 50%; background: #c62828; color: white; font-size: 22px; font-weight: bold; }
	.popup-mais { margin-top: -8px; color: #786b5c; font-size: 13px; }
	@keyframes aparecer { from { opacity: 0; transform: scale(.97); } to { opacity: 1; transform: scale(1); } }
	@media (max-width: 760px) { header { position: relative; width: 100%; min-height: auto; padding: 16px; } header > div:not(.logo) { padding: 0; } .logo { padding: 4px 0 14px; } .menu { flex-direction: row; flex-wrap: wrap; margin-top: 14px; } .menu a { padding: 8px 9px; } .container { margin-left: 0; padding: 20px 12px; } }
	</style>
	"""
	return render_template_string(HTML_BASE, titulo=titulo, conteudo=estilo_moderno + conteudo + popup_alertas())


def acesso_login():
	return redirect(url_for("login")) if not usuario_logado() else None


@app.route("/configurar", methods=["GET", "POST"])
def configurar():
	if request.method == "POST":
		nome, usuario, senha = request.form["nome"].strip(), request.form["usuario"].strip(), request.form["senha"]
		if len(senha) < 8:
			return pagina("Configuração", '<div class="card"><div class="erro">A senha precisa ter pelo menos 8 caracteres.</div></div>')
		def mutator(workbook):
			planilha = workbook["usuarios"]
			cabecalho = _cabecalho_planilha(planilha)
			if planilha.max_row > 1:
				raise RuntimeError("O sistema já possui usuário cadastrado.")
			registro = {"id": _proximo_id(planilha, cabecalho), "nome": nome, "usuario": usuario, "email": "", "senha_hash": bcrypt.hashpw(senha.encode(), bcrypt.gensalt()).decode(), "perfil": "Administrador", "ativo": 1, "aprovado": 1, "criado_em": agora()}
			_append_registro(planilha, TABELAS_EXCEL["usuarios"], registro)
			return registro
		try:
			executar_mutacao_atomica(mutator, lambda conteudo, r: _confirmar_id_na_aba(conteudo, "usuarios", r["id"]))
			return redirect(url_for("login"))
		except Exception as erro:
			return pagina("Configuração", f'<div class="card"><div class="erro">Não foi possível concluir a configuração: {html.escape(str(erro))}</div></div>')
	try:
		_, remoto = _graph_snapshot()
		wb = load_workbook(io.BytesIO(remoto), read_only=True, data_only=True)
		total = max(wb["usuarios"].max_row - 1, 0) if "usuarios" in wb.sheetnames else 0
		wb.close()
		if total:
			return redirect(url_for("login"))
	except Exception:
		pass
	return pagina("Configuração inicial", '<div class="card" style="max-width:500px;margin:auto"><h1>SP ÁGUAS</h1><h2>Configuração inicial</h2><p>Crie o primeiro usuário administrador.</p><form method="post"><label>Nome completo</label><input name="nome" required><label>Usuário</label><input name="usuario" required><label>Senha</label><input type="password" name="senha" minlength="8" required><button>Criar administrador</button></form></div>')


@app.route("/", methods=["GET", "POST"])
def login():
	conn = conectar(); total = conn.execute("SELECT COUNT(*) total FROM usuarios").fetchone()["total"]
	if not total:
		conn.close(); return redirect(url_for("configurar"))
	if request.method == "POST":
		user = conn.execute("SELECT * FROM usuarios WHERE usuario = ?", (request.form["usuario"].strip(),)).fetchone(); conn.close()
		if not user or not user["ativo"] or not user["aprovado"] or not bcrypt.checkpw(request.form["senha"].encode(), user["senha_hash"].encode()):
			return pagina("Login", '<div class="card" style="max-width:420px;margin:auto"><div class="erro">Usuário ou senha inválidos, ou acesso ainda não liberado.</div><br><a class="btn" href="/">Voltar</a></div>')
		session.update(usuario_id=user["id"], nome=user["nome"], usuario=user["usuario"], perfil=user["perfil"])
		return redirect(url_for("sistema"))
	conn.close()
	return pagina("Login", '<div class="card" style="max-width:420px;margin:70px auto"><h1>SP ÁGUAS</h1><p>Gestão de Processos e Prazos</p><form method="post"><label>Usuário</label><input name="usuario" required><label>Senha</label><input type="password" name="senha" required><button style="width:100%">ENTRAR</button></form><hr><a href="/cadastro" class="btn">Criar minha conta</a></div>')


@app.route("/cadastro", methods=["GET", "POST"])
def cadastro():
	if request.method == "POST":
		nome, usuario, email = request.form["nome"].strip(), request.form["usuario"].strip(), request.form["email"].strip()
		senha = request.form["senha"]
		if senha != request.form["confirmar"]: erro = "As senhas não coincidem."
		elif len(senha) < 8: erro = "A senha precisa ter pelo menos 8 caracteres."
		else:
			def mutator(workbook):
				planilha = workbook["usuarios"]
				cabecalho = _cabecalho_planilha(planilha)
				idx_usuario, idx_email = cabecalho.index("usuario"), cabecalho.index("email")
				for linha in planilha.iter_rows(min_row=2, values_only=True):
					if linha[idx_usuario] == usuario or (email and linha[idx_email] == email):
						raise sqlite3.IntegrityError("Usuário ou e-mail já cadastrado.")
				registro = {"id": _proximo_id(planilha, cabecalho), "nome": nome, "usuario": usuario, "email": email, "senha_hash": bcrypt.hashpw(senha.encode(), bcrypt.gensalt()).decode(), "perfil": "Usuario", "ativo": 1, "aprovado": 0, "criado_em": agora()}
				_append_registro(planilha, TABELAS_EXCEL["usuarios"], registro)
				return registro
			try:
				executar_mutacao_atomica(mutator, lambda conteudo, r: _confirmar_id_na_aba(conteudo, "usuarios", r["id"]))
				return pagina("Cadastro realizado", '<div class="card"><h1>Cadastro realizado</h1><p>Aguarde a aprovação do administrador.</p><a class="btn" href="/">Voltar ao login</a></div>')
			except sqlite3.IntegrityError:
				erro = "Usuário ou e-mail já cadastrado."
			except Exception as exc:
				erro = f"Não foi possível concluir o cadastro: {exc}"
		return pagina("Cadastro", f'<div class="card"><div class="erro">{html.escape(erro)}</div><a href="/cadastro" class="btn">Voltar</a></div>')
	return pagina("Cadastro", '<div class="card" style="max-width:550px;margin:auto"><h2>Criar acesso ao sistema</h2><form method="post"><label>Nome completo</label><input name="nome" required><label>Usuário</label><input name="usuario" required><label>E-mail</label><input type="email" name="email" required><label>Senha</label><input type="password" name="senha" minlength="8" required><label>Confirmar senha</label><input type="password" name="confirmar" minlength="8" required><button>Criar minha conta</button></form></div>')


def ler_demanda_form():
	return tuple(request.form.get(c, "").strip() for c in ("numero", "origem", "assunto", "area", "responsavel", "data_recebimento", "prazo_area", "prazo_fatal", "situacao", "prioridade", "observacoes"))


def validar_demanda(valores):
	campos = ("numero do processo", "origem", "assunto", "área", "responsável", "data de recebimento", "prazo da área", "prazo fatal", "situação", "prioridade", "observações")
	dados = dict(zip(campos, valores))
	if not dados["origem"] or not dados["assunto"] or not dados["prazo fatal"]:
		return "Preencha a origem, o assunto e o prazo fatal."
	if dados["origem"] not in {"TCE-SP", "AUDESP", "Outro"}:
		return "A origem informada é inválida."
	for nome in ("data de recebimento", "prazo da área", "prazo fatal"):
		if dados[nome]:
			try:
				datetime.strptime(dados[nome], "%Y-%m-%d")
			except ValueError:
				return f"A {nome} deve ser uma data válida."
	return None


@app.route("/sistema")
def sistema():
	if (resposta := acesso_login()): return resposta
	conn = conectar(); registros = conn.execute("SELECT * FROM demandas ORDER BY prazo_fatal").fetchall(); conn.close()
	contagens = {"total": len(registros), "tce": sum(d["origem"] == "TCE-SP" for d in registros), "audesp": sum(d["origem"] == "AUDESP" for d in registros), "concluidas": sum(d["situacao"] == "Concluído" for d in registros), "VENCIDO": 0, "CRÍTICO": 0, "PRÓXIMO": 0, "NORMAL": 0}
	for d in registros:
		if d["situacao"] != "Concluído": contagens[calcular_prazos(d["prazo_area"], d["prazo_fatal"])["status"]] += 1
	cards = [("Total de demandas", "total"), ("TCE-SP", "tce"), ("AUDESP", "audesp"), ("Vencidas", "VENCIDO"), ("Críticas", "CRÍTICO"), ("Próximas", "PRÓXIMO"), ("Normais", "NORMAL"), ("Concluídas", "concluidas")]
	html = f'<h1>Dashboard</h1><p>Bem-vindo, <strong>{session["nome"]}</strong>. Perfil: <strong>{session["perfil"]}</strong></p><div class="grid">' + ''.join(f'<div class="metrica"><h3>{nome}</h3><strong>{contagens[chave]}</strong></div>' for nome, chave in cards) + '</div><div class="card"><h2>Próximos prazos</h2>'
	proximos = sorted(((d, calcular_prazos(d["prazo_area"], d["prazo_fatal"])) for d in registros if d["situacao"] != "Concluído" and calcular_prazos(d["prazo_area"], d["prazo_fatal"])["dias_fatal"] is not None), key=lambda item: item[1]["dias_fatal"])
	html += ''.join(f'<div class="alerta {p["status"].lower()}"><strong>{d["numero_processo"] or "Sem número"}</strong> — {d["assunto"]}<br>Prazo fatal: <strong>{data_br(d["prazo_fatal"])}</strong> | Dias restantes: <strong>{p["dias_fatal"]}</strong> | <a href="/editar/{d["id"]}">Editar</a></div>' for d, p in proximos[:10]) + '</div>'
	return pagina("Dashboard", html)


@app.route("/nova-demanda", methods=["GET", "POST"])
def nova_demanda():
	if (resposta := acesso_login()): return resposta
	if request.method == "POST":
		valores = ler_demanda_form(); erro = validar_demanda(valores)
		if erro:
			return pagina("Nova demanda", f'<div class="card"><div class="erro">{erro}</div></div>' + formulario())
		def mutator(workbook):
			planilha = workbook["demandas"]
			cabecalho = _cabecalho_planilha(planilha)
			demanda_id = _proximo_id(planilha, cabecalho)
			agora_valor = agora()
			registro = dict(zip(TABELAS_EXCEL["demandas"], (demanda_id, *valores, session["usuario_id"], agora_valor, agora_valor)))
			_append_registro(planilha, TABELAS_EXCEL["demandas"], registro)
			planilha_h = workbook["historico"]
			historico_id = _proximo_id(planilha_h, _cabecalho_planilha(planilha_h))
			historico = {"id": historico_id, "demanda_id": demanda_id, "usuario_id": session["usuario_id"], "acao": "CRIACAO", "descricao": "Demanda cadastrada.", "data_hora": agora_valor}
			_append_registro(planilha_h, TABELAS_EXCEL["historico"], historico)
			return {"demanda_id": demanda_id, "historico_id": historico_id}
		try:
			resultado = executar_mutacao_atomica(
				mutator,
				lambda conteudo, r: _confirmar_id_na_aba(conteudo, "demandas", r["demanda_id"]) and _confirmar_id_na_aba(conteudo, "historico", r["historico_id"]),
			)
		except Exception as erro:
			return pagina("Nova demanda", f'<div class="card"><div class="erro">Demanda NÃO confirmada no SharePoint. {html.escape(str(erro))}</div><a class="btn" href="/nova-demanda">Tentar novamente</a> <a class="btn btn-cinza" href="/demandas">Voltar</a></div>')
		return redirect(url_for("demandas"))
	return pagina("Nova demanda", formulario())


def formulario(demanda=None):
	valor = lambda nome: (demanda[nome] or "") if demanda else ""
	return f'<div class="card"><h1>{"Editar demanda" if demanda else "Nova demanda"}</h1><form method="post"><div class="grid"><div><label>Número do processo</label><input name="numero" value="{valor("numero_processo")}"></div><div><label>Origem</label><select name="origem"><option>TCE-SP</option><option>AUDESP</option><option>Outro</option></select></div><div><label>Data de recebimento</label><input type="date" name="data_recebimento" value="{valor("data_recebimento")}"></div><div><label>Área</label><input name="area" value="{valor("area")}"></div><div><label>Responsável</label><input name="responsavel" value="{valor("responsavel")}"></div><div><label>Prioridade</label><select name="prioridade"><option>Normal</option><option>Alta</option><option>Urgente</option></select></div><div><label>Prazo da área</label><input type="date" name="prazo_area" value="{valor("prazo_area")}"></div><div><label>Prazo fatal</label><input type="date" name="prazo_fatal" value="{valor("prazo_fatal")}" required></div><div><label>Situação</label><select name="situacao"><option>Aberto</option><option>Em análise</option><option>Aguardando área</option><option>Aguardando documento</option><option>Respondido</option><option>Concluído</option></select></div></div><label>Assunto</label><textarea name="assunto" required>{valor("assunto")}</textarea><label>Observações</label><textarea name="observacoes">{valor("observacoes")}</textarea><button>Salvar</button> <a class="btn btn-cinza" href="/demandas">Cancelar</a></form></div>'


@app.route("/demandas")
def demandas():
	if (resposta := acesso_login()): return resposta
	busca, origem = request.args.get("busca", "").strip(), request.args.get("origem", "")
	conn = conectar(); sql = "SELECT * FROM demandas WHERE 1=1"; params = []
	if busca: sql += " AND (numero_processo LIKE ? OR assunto LIKE ? OR area LIKE ? OR responsavel LIKE ?)"; params += [f"%{busca}%"] * 4
	if origem: sql += " AND origem = ?"; params.append(origem)
	registros = conn.execute(sql + " ORDER BY prazo_fatal", params).fetchall(); conn.close()
	linhas = ''.join(f'<tr><td>{d["id"]}</td><td>{d["numero_processo"] or "-"}</td><td>{d["origem"]}</td><td>{d["assunto"]}</td><td>{d["area"] or "-"}</td><td>{d["responsavel"] or "-"}</td><td>{data_br(d["prazo_fatal"])}</td><td>{d["situacao"]}</td><td><a class="btn" href="/editar/{d["id"]}">Editar</a></td></tr>' for d in registros)
	return pagina("Demandas", f'<div class="card"><h1>Demandas</h1><form method="get"><input name="busca" value="{busca}" placeholder="Processo, assunto, área..."><button>Pesquisar</button></form></div><div class="card" style="overflow-x:auto"><table><tr><th>ID</th><th>Processo</th><th>Origem</th><th>Assunto</th><th>Área</th><th>Responsável</th><th>Prazo fatal</th><th>Situação</th><th>Ações</th></tr>{linhas}</table></div>')


@app.route("/editar/<int:id>", methods=["GET", "POST"])
def editar(id):
	if (resposta := acesso_login()): return resposta
	conn = conectar(); demanda = conn.execute("SELECT * FROM demandas WHERE id = ?", (id,)).fetchone()
	if not demanda: conn.close(); return "Demanda não encontrada.", 404
	if request.method == "POST":
		valores = ler_demanda_form(); erro = validar_demanda(valores)
		if erro:
			conn.close(); return pagina("Editar demanda", f'<div class="card"><div class="erro">{erro}</div></div>' + formulario(demanda))
		usuario_id = session["usuario_id"]
		def mutator(workbook):
			planilha = workbook["demandas"]
			registro = dict(zip(TABELAS_EXCEL["demandas"], (id, *valores, demanda["criado_por"], demanda["criado_em"], agora())))
			_atualizar_registro(planilha, TABELAS_EXCEL["demandas"], id, registro)
			ph = workbook["historico"]
			hid = _proximo_id(ph, _cabecalho_planilha(ph))
			historico = {"id": hid, "demanda_id": id, "usuario_id": usuario_id, "acao": "EDICAO", "descricao": "Demanda alterada.", "data_hora": agora()}
			_append_registro(ph, TABELAS_EXCEL["historico"], historico)
			return {"demanda_id": id, "historico_id": hid}
		try:
			executar_mutacao_atomica(mutator, lambda conteudo, r: _confirmar_id_na_aba(conteudo, "demandas", id) and _confirmar_id_na_aba(conteudo, "historico", r["historico_id"]))
		except Exception as exc:
			conn.close()
			return pagina("Editar demanda", f'<div class="card"><div class="erro">Alteração NÃO confirmada no SharePoint: {html.escape(str(exc))}</div><a class="btn" href="/editar/{id}">Tentar novamente</a></div>')
		conn.close(); return redirect(url_for("demandas"))
	conn.close(); return pagina("Editar demanda", formulario(demanda))


@app.route("/excluir/<int:id>", methods=["POST"])
def excluir(id):
	if not administrador(): return "Acesso negado.", 403
	def mutator(workbook):
		planilha = workbook["demandas"]
		cab = _cabecalho_planilha(planilha)
		linha = _linha_por_id(planilha, cab, id)
		if linha is None:
			raise RuntimeError("Demanda não encontrada.")
		planilha.delete_rows(linha, 1)
		hist = workbook["historico"]
		cab_h = _cabecalho_planilha(hist)
		idx_demanda = cab_h.index("demanda_id") + 1
		for linha_h in range(hist.max_row, 1, -1):
			if hist.cell(linha_h, idx_demanda).value == id:
				hist.delete_rows(linha_h, 1)
		return {"demanda_id": id}
	try:
		executar_mutacao_atomica(mutator, lambda conteudo, r: not _confirmar_id_na_aba(conteudo, "demandas", id))
	except Exception as exc:
		return pagina("Excluir demanda", f'<div class="card"><div class="erro">Exclusão NÃO confirmada no SharePoint: {html.escape(str(exc))}</div><a class="btn" href="/demandas">Voltar</a></div>'), 500
	return redirect(url_for("demandas"))


@app.route("/alertas")
def alertas():
	if (resposta := acesso_login()): return resposta
	conn = conectar(); registros = conn.execute("SELECT * FROM demandas WHERE situacao != 'Concluído' ORDER BY prazo_fatal").fetchall(); conn.close()
	html = '<div class="card"><h1>Alertas de prazos</h1>'
	for d in registros:
		p = calcular_prazos(d["prazo_area"], d["prazo_fatal"])
		if p["status"] != "NORMAL": html += f'<div class="alerta {p["status"].lower()}"><strong>{p["status"]}</strong> — {d["numero_processo"] or "-"} — {d["assunto"]}<br>Prazo fatal: {data_br(d["prazo_fatal"])} | Dias restantes: {p["dias_fatal"]} <a class="btn" href="/editar/{d["id"]}">Abrir</a></div>'
	return pagina("Alertas", html + '</div>')


@app.route("/usuarios", methods=["GET", "POST"])
def usuarios():
	if not administrador(): return "Acesso negado.", 403
	if request.method == "POST":
		acao, usuario_id = request.form["acao"], int(request.form["usuario_id"])
		if acao not in {"aprovar", "bloquear", "ativar"}: return "Ação inválida.", 400
		def mutator(workbook):
			planilha = workbook["usuarios"]
			cab = _cabecalho_planilha(planilha)
			linha = _linha_por_id(planilha, cab, usuario_id)
			if linha is None: raise RuntimeError("Usuário não encontrado.")
			if acao == "aprovar":
				planilha.cell(linha, cab.index("aprovado") + 1).value = 1
				planilha.cell(linha, cab.index("ativo") + 1).value = 1
			elif acao == "bloquear":
				planilha.cell(linha, cab.index("ativo") + 1).value = 0
			else:
				planilha.cell(linha, cab.index("ativo") + 1).value = 1
				planilha.cell(linha, cab.index("aprovado") + 1).value = 1
			return {"usuario_id": usuario_id, "acao": acao}
		def confirmar(conteudo, resultado):
			wb = load_workbook(io.BytesIO(conteudo), read_only=True, data_only=True)
			try:
				ws = wb["usuarios"]; cab = _cabecalho_planilha(ws); idx = cab.index("id"); ia = cab.index("ativo"); ip = cab.index("aprovado")
				for linha in ws.iter_rows(min_row=2, values_only=True):
					if linha[idx] == resultado["usuario_id"]:
						if resultado["acao"] == "bloquear": return linha[ia] == 0
						return linha[ia] == 1 and linha[ip] == 1
				return False
			finally: wb.close()
		try:
			executar_mutacao_atomica(mutator, confirmar)
		except Exception as exc:
			return pagina("Usuários", f'<div class="card"><div class="erro">Alteração de usuário NÃO confirmada no SharePoint: {html.escape(str(exc))}</div></div>')
	conn = conectar(); lista = conn.execute("SELECT * FROM usuarios ORDER BY nome").fetchall(); conn.close()
	linhas = ''.join(f'<tr><td>{u["nome"]}</td><td>{u["usuario"]}</td><td>{u["email"] or "-"}</td><td>{u["perfil"]}</td><td>{"Aprovado" if u["aprovado"] else "Pendente"}</td><td>{"Ativo" if u["ativo"] else "Bloqueado"}</td><td><form method="post"><input type="hidden" name="usuario_id" value="{u["id"]}"><input type="hidden" name="acao" value="{"bloquear" if u["ativo"] else "ativar"}"><button>{"Bloquear" if u["ativo"] else "Ativar"}</button></form></td></tr>' for u in lista)
	return pagina("Usuários", f'<div class="card"><h1>Usuários</h1><table><tr><th>Nome</th><th>Usuário</th><th>E-mail</th><th>Perfil</th><th>Aprovação</th><th>Status</th><th>Ação</th></tr>{linhas}</table></div>')


def tabela_por_origem(titulo, origem):
	conn = conectar()
	registros = conn.execute("SELECT * FROM demandas WHERE origem = ? ORDER BY prazo_fatal", (origem,)).fetchall()
	conn.close()
	linhas = ""
	for demanda in registros:
		prazo = calcular_prazos(demanda["prazo_area"], demanda["prazo_fatal"])
		classe = {"VENCIDO": "vencido", "CRÍTICO": "critico", "PRÓXIMO": "proximo", "NORMAL": "normal"}.get(prazo["status"], "info")
		linhas += f'<tr><td>{demanda["numero_processo"] or "-"}</td><td>{demanda["assunto"]}</td><td>{demanda["area"] or "-"}</td><td>{demanda["responsavel"] or "-"}</td><td>{data_br(demanda["prazo_fatal"])}</td><td><span class="alerta {classe}">{prazo["status"]}</span></td><td><a class="btn" href="/editar/{demanda["id"]}">Abrir</a></td></tr>'
	conteudo = f'<div class="card"><h1>{titulo}</h1><p>Demandas classificadas como {origem}.</p><div style="overflow-x:auto"><table><tr><th>Processo</th><th>Assunto</th><th>Área</th><th>Responsável</th><th>Prazo fatal</th><th>Situação</th><th>Ação</th></tr>{linhas}</table></div></div>'
	return pagina(titulo, conteudo)


@app.route("/tce")
def tce():
	if (resposta := acesso_login()):
		return resposta
	return tabela_por_origem("TCE-SP", "TCE-SP")


@app.route("/audesp")
def audesp():
	if (resposta := acesso_login()):
		return resposta
	return tabela_por_origem("AUDESP", "AUDESP")


@app.route("/calendario")
def calendario():
	if (resposta := acesso_login()):
		return resposta
	hoje = date.today()
	conn = conectar()
	registros = conn.execute("SELECT * FROM demandas WHERE prazo_fatal IS NOT NULL ORDER BY prazo_fatal").fetchall()
	conn.close()
	eventos = {}
	for demanda in registros:
		try:
			prazo = datetime.strptime(demanda["prazo_fatal"], "%Y-%m-%d").date()
			if prazo.year == hoje.year and prazo.month == hoje.month:
				eventos.setdefault(prazo.day, []).append(demanda)
		except (TypeError, ValueError):
			continue
	import calendar as calendario_lib
	primeiro_dia, total_dias = calendario_lib.monthrange(hoje.year, hoje.month)
	cabecalho = "".join(f"<div><strong>{nome}</strong></div>" for nome in ("Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"))
	celulas = "<div></div>" * primeiro_dia
	for dia in range(1, total_dias + 1):
		itens = "".join(f'<div style="font-size:11px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap"><a href="/editar/{item["id"]}">{item["numero_processo"] or item["assunto"][:20]}</a></div>' for item in eventos.get(dia, [])[:3])
		celulas += f'<div class="card" style="min-height:100px;margin:0"><strong>{dia}</strong>{itens}</div>'
	conteudo = f'<div class="card"><h1>Calendário</h1><h2>{hoje.strftime("%m/%Y")}</h2><div style="display:grid;grid-template-columns:repeat(7,1fr);gap:8px">{cabecalho}{celulas}</div></div>'
	return pagina("Calendário", conteudo)


@app.route("/exportar")
def exportar():
	if (resposta := acesso_login()): return resposta
	from openpyxl import Workbook
	conn = conectar(); dados = conn.execute("SELECT * FROM demandas ORDER BY prazo_fatal").fetchall(); conn.close(); wb = Workbook(); ws = wb.active; ws.title = "Demandas"
	ws.append(["ID", "Processo", "Origem", "Assunto", "Área", "Responsável", "Prazo área", "Prazo fatal", "Dias entre", "Dias restantes", "Situação", "Prioridade", "Observações"])
	for d in dados:
		p = calcular_prazos(d["prazo_area"], d["prazo_fatal"]); ws.append([d["id"], d["numero_processo"], d["origem"], d["assunto"], d["area"], d["responsavel"], data_br(d["prazo_area"]), data_br(d["prazo_fatal"]), p["dias_entre"], p["dias_fatal"], d["situacao"], d["prioridade"], d["observacoes"]])
	arquivo = io.BytesIO(); wb.save(arquivo); arquivo.seek(0); return send_file(arquivo, as_attachment=True, download_name="SP_AGUAS_Demandas.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/status")
def status():
	if (resposta := acesso_login()): return resposta
	resultado = {"status": "OK", "sharepoint_configurado": _onedrive_configurado(), "planilha_local": PLANILHA, "vercel": bool(os.environ.get("VERCEL"))}
	if not _onedrive_configurado():
		resultado.update(status="ERRO", erro="Variáveis do SharePoint não estão completas.")
		return jsonify(resultado), 500
	try:
		with urlopen(Request(_onedrive_url(), headers={"Authorization": f"Bearer {_onedrive_token()}"}), timeout=30) as resposta:
			conteudo = resposta.read()
		resultado.update(token_graph=True, arquivo_sharepoint_acessivel=True, tamanho_bytes=len(conteudo), mensagem="CONEXÃO COM SHAREPOINT OK")
		return jsonify(resultado), 200
	except Exception as erro:
		resultado.update(status="ERRO", token_graph=False, arquivo_sharepoint_acessivel=False, erro=str(erro))
		return jsonify(resultado), 500


@app.route("/teste-integracao")
def teste_integracao():
	if not usuario_logado():
		return jsonify({"status": "ERRO", "mensagem": "Usuário não autenticado."}), 401
	if not administrador():
		return jsonify({"status": "ERRO", "mensagem": "Somente administradores podem executar o teste."}), 403

	resultado = {"teste": "Integração SP ÁGUAS + Microsoft Graph + SharePoint", "status_final": "INICIANDO", "arquivo": SHAREPOINT_FILE_PATH, "etapas": []}

	def etapa(nome, status, mensagem, **dados):
		item = {"etapa": nome, "status": status, "mensagem": mensagem}
		item.update(dados)
		resultado["etapas"].append(item)

	if not _onedrive_configurado():
		etapa("1. Configuração", "ERRO", "Integração SharePoint desabilitada ou incompleta.")
		resultado["status_final"] = "FALHA"
		return jsonify(resultado), 500
	etapa("1. Configuração", "OK", "Configuração SharePoint habilitada.", arquivo=SHAREPOINT_FILE_PATH)

	try:
		token = _onedrive_token()
		etapa("2. Autenticação Microsoft Graph", "OK", "Token obtido com sucesso.")
	except Exception as erro:
		etapa("2. Autenticação Microsoft Graph", "ERRO", str(erro))
		resultado["status_final"] = "FALHA"
		return jsonify(resultado), 500

	try:
		with urlopen(Request(_onedrive_url(), headers={"Authorization": f"Bearer {token}"}), timeout=60) as resposta:
			arquivo_original = resposta.read()
			status_http = resposta.status
		if status_http != 200 or not arquivo_original:
			raise RuntimeError(f"Microsoft Graph retornou HTTP {status_http} ou arquivo vazio.")
		hash_original = hashlib.sha256(arquivo_original).hexdigest()
		etapa("3. Download do Excel", "OK", "Arquivo baixado com sucesso.", tamanho_bytes=len(arquivo_original), sha256=hash_original)
	except Exception as erro:
		etapa("3. Download do Excel", "ERRO", str(erro))
		resultado["status_final"] = "FALHA"
		return jsonify(resultado), 500

	try:
		workbook = load_workbook(io.BytesIO(arquivo_original), read_only=True, data_only=True)
		abas = workbook.sheetnames
		faltantes = {"usuarios", "demandas", "historico"} - set(abas)
		workbook.close()
		if faltantes:
			raise RuntimeError("Faltam as abas: " + ", ".join(sorted(faltantes)))
		etapa("4. Validação do Excel", "OK", "Arquivo XLSX válido e estrutura esperada encontrada.", abas=abas)
	except Exception as erro:
		etapa("4. Validação do Excel", "ERRO", str(erro))
		resultado["status_final"] = "FALHA"
		return jsonify(resultado), 500

	try:
		conn = conectar()
		contagens = {tabela: conn.execute(f"SELECT COUNT(*) FROM {tabela}").fetchone()[0] for tabela in ("usuarios", "demandas", "historico")}
		conn.close()
		etapa("5. Leitura dos dados", "OK", "Dados carregados no sistema.", **contagens)
	except Exception as erro:
		etapa("5. Leitura dos dados", "ERRO", str(erro))
		resultado["status_final"] = "FALHA"
		return jsonify(resultado), 500

	try:
		meta_teste = _graph_metadata_diagnostico()
		status_put = _graph_upload_diagnostico(arquivo_original, etag=meta_teste.get("eTag"))
		if status_put not in (200, 201):
			raise RuntimeError(f"Microsoft Graph retornou HTTP {status_put}.")
		etapa("6. Gravação no SharePoint", "OK", "O Graph aceitou os mesmos bytes do arquivo.")
	except Exception as erro:
		etapa("6. Gravação no SharePoint", "ERRO", str(erro))
		resultado["status_final"] = "FALHA"
		return jsonify(resultado), 500

	try:
		with urlopen(Request(_onedrive_url(), headers={"Authorization": f"Bearer {token}"}), timeout=60) as resposta:
			arquivo_depois = resposta.read()
		hash_depois = hashlib.sha256(arquivo_depois).hexdigest()
		if hash_original != hash_depois:
			raise RuntimeError("O arquivo recuperado é diferente do arquivo enviado.")
		etapa("7. Validação da persistência", "OK", "O arquivo recuperado é idêntico ao enviado.", sha256=hash_depois)
	except Exception as erro:
		etapa("7. Validação da persistência", "ERRO", str(erro))
		resultado["status_final"] = "FALHA"
		return jsonify(resultado), 500

	resultado["status_final"] = "OK"
	resultado["diagnostico"] = "A integração SharePoint está funcionando e a persistência foi confirmada por SHA-256."
	return jsonify(resultado), 200


@app.route("/health")
def health():
	return jsonify({"status": "ok", "service": "sp-aguas"}), 200


@app.route("/diagnostico-sync")
def diagnostico_sync():
	"""Diagnostica a cadeia banco em memória -> XLSX -> SharePoint."""
	if (resposta := acesso_login()):
		return resposta
	if not administrador():
		return jsonify({"status": "FALHA", "erro": "Somente administradores podem executar o diagnóstico."}), 403

	inicio = datetime.now()
	resultado = {"status": "INICIANDO", "arquivo_configurado": SHAREPOINT_FILE_PATH, "etapas": []}

	def etapa(numero, nome, status, mensagem, **dados):
		resultado["etapas"].append({"numero": numero, "etapa": nome, "status": status, "mensagem": mensagem, **dados})

	def falha(numero, nome, erro):
		etapa(numero, nome, "ERRO", str(erro))
		resultado["status"] = "FALHA"
		return jsonify(resultado), 500

	try:
		if not _onedrive_configurado():
			raise RuntimeError("Integração SharePoint desabilitada ou incompleta.")
		etapa(1, "Configuração", "OK", "Configuração do Microsoft Graph está preenchida.", caminho=SHAREPOINT_FILE_PATH)
	except Exception as erro:
		return falha(1, "Configuração", erro)

	try:
		token = _onedrive_token()
		etapa(2, "Autenticação Microsoft Graph", "OK", "Access token obtido com sucesso.")
	except Exception as erro:
		return falha(2, "Autenticação Microsoft Graph", erro)

	try:
		meta_antes = _graph_metadata_diagnostico()
		etapa(3, "Arquivo remoto", "OK", "Microsoft Graph encontrou o arquivo.", nome=meta_antes.get("name"), id=meta_antes.get("id"), tamanho_bytes=meta_antes.get("size"), etag=meta_antes.get("eTag"))
	except Exception as erro:
		return falha(3, "Arquivo remoto", erro)

	try:
		remoto_antes = _graph_download_diagnostico()
		hash_antes = hashlib.sha256(remoto_antes).hexdigest()
		etapa(4, "Download do Excel remoto", "OK", "Arquivo remoto baixado.", tamanho_bytes=len(remoto_antes), sha256=hash_antes)
	except Exception as erro:
		return falha(4, "Download do Excel remoto", erro)

	try:
		workbook = load_workbook(io.BytesIO(remoto_antes), read_only=True, data_only=True)
		abas = list(workbook.sheetnames)
		faltantes = set(TABELAS_EXCEL) - set(abas)
		contagens_remotas = {tabela: max(workbook[tabela].max_row - 1, 0) if tabela in abas else None for tabela in TABELAS_EXCEL}
		workbook.close()
		if faltantes:
			raise RuntimeError("Faltam as abas: " + ", ".join(sorted(faltantes)))
		etapa(5, "Estrutura do Excel remoto", "OK", "XLSX válido.", abas=abas, contagens=contagens_remotas)
	except Exception as erro:
		return falha(5, "Estrutura do Excel remoto", erro)

	try:
		conn = conectar()
		contagens = {tabela: conn.execute(f"SELECT COUNT(*) FROM {tabela}").fetchone()[0] for tabela in TABELAS_EXCEL}
		etapa(6, "Dados da aplicação", "OK", "Contagens obtidas do banco em memória.", contagens=contagens, contagens_remotas=contagens_remotas)
	except Exception as erro:
		return falha(6, "Dados da aplicação", erro)

	try:
		workbook = Workbook()
		workbook.remove(workbook.active)
		for tabela, colunas in TABELAS_EXCEL.items():
			planilha = workbook.create_sheet(tabela)
			planilha.append(list(colunas))
			for registro in conn.execute(f"SELECT {', '.join(colunas)} FROM {tabela} ORDER BY id"):
				planilha.append([registro[coluna] for coluna in colunas])
		memoria = io.BytesIO()
		workbook.save(memoria)
		arquivo_gerado = memoria.getvalue()
		hash_gerado = hashlib.sha256(arquivo_gerado).hexdigest()
		etapa(7, "Geração do XLSX", "OK", "XLSX reconstruído pela aplicação.", tamanho_bytes=len(arquivo_gerado), sha256=hash_gerado)
	except Exception as erro:
		return falha(7, "Geração do XLSX", erro)

	try:
		status_put = _graph_upload_diagnostico(arquivo_gerado)
		if status_put not in (200, 201):
			raise RuntimeError(f"Microsoft Graph retornou HTTP {status_put}.")
		etapa(8, "Gravação no SharePoint", "OK", "XLSX gerado enviado ao arquivo remoto.", http_status=status_put)
	except Exception as erro:
		return falha(8, "Gravação no SharePoint", erro)

	try:
		remoto_depois = _graph_download_diagnostico()
		hash_depois = hashlib.sha256(remoto_depois).hexdigest()
		if hash_depois != hash_gerado:
			raise RuntimeError("O arquivo remoto ficou diferente do XLSX enviado.")
		etapa(9, "Confirmação da persistência", "OK", "SHA-256 do arquivo remoto coincide com o arquivo enviado.", sha256=hash_depois)
	except Exception as erro:
		return falha(9, "Confirmação da persistência", erro)

	try:
		meta_depois = _graph_metadata_diagnostico()
		etapa(10, "Metadados finais", "OK", "Metadados consultados após a gravação.", tamanho_bytes=meta_depois.get("size"), etag=meta_depois.get("eTag"))
	except Exception as erro:
		etapa(10, "Metadados finais", "AVISO", str(erro))

	resultado.update(status="OK", duracao_segundos=round((datetime.now() - inicio).total_seconds(), 2), conclusao="A persistência do XLSX no SharePoint foi confirmada.")
	return jsonify(resultado), 200


@app.route("/logout")
def logout():
	session.clear(); return redirect(url_for("login"))


criar_banco()

if __name__ == "__main__":
	app.run(host="0.0.0.0", port=5000, debug=False)

