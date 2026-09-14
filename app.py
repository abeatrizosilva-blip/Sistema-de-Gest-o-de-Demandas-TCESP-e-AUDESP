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
def _env(*nomes, default=""):
	"""Lê a primeira variável disponível e remove espaços acidentais."""
	for nome in nomes:
		valor = os.environ.get(nome)
		if valor is not None and str(valor).strip() != "":
			return str(valor).strip()
	return default


# Credenciais continuam sendo obrigatoriamente configuradas no Vercel.
# O usuário/pasta/arquivo possuem valores padrão conhecidos do projeto para
# evitar que uma variável opcional ausente seja confundida com falha de integração.
_raw_enabled = _env("SHAREPOINT_ENABLED", "ONEDRIVE_ENABLED")
ONEDRIVE_TENANT_ID = _env("SHAREPOINT_TENANT_ID", "ONEDRIVE_TENANT_ID", "MICROSOFT_TENANT_ID")
ONEDRIVE_CLIENT_ID = _env("SHAREPOINT_CLIENT_ID", "ONEDRIVE_CLIENT_ID", "MICROSOFT_CLIENT_ID")
ONEDRIVE_CLIENT_SECRET = _env("SHAREPOINT_CLIENT_SECRET", "ONEDRIVE_CLIENT_SECRET", "MICROSOFT_CLIENT_SECRET")
ONEDRIVE_USER = _env(
	"SHAREPOINT_USER", "ONEDRIVE_USER", "MICROSOFT_USER",
	default="anabeatriz.silva@spaguas.sp.gov.br"
)
SHAREPOINT_FOLDER_PATH = _env("SHAREPOINT_FOLDER_PATH", "ONEDRIVE_FOLDER_PATH", default="SP_AGUAS")
SHAREPOINT_FILE_NAME = _env(
	"SHAREPOINT_FILE_NAME", "ONEDRIVE_FILE_NAME",
	default="Sistema de Gestão de Demandas - SP Aguas.xlsx"
)
SHAREPOINT_FILE_PATH = _env(
	"SHAREPOINT_FILE_PATH", "ONEDRIVE_PATH",
	default=f"{SHAREPOINT_FOLDER_PATH.strip('/')}/{SHAREPOINT_FILE_NAME}"
)
ONEDRIVE_PATH = SHAREPOINT_FILE_PATH

# Se a flag não existir, habilita automaticamente quando as três credenciais
# essenciais estiverem presentes. Se SHAREPOINT_ENABLED=false for informado,
# a integração permanece explicitamente desabilitada.
if _raw_enabled:
	ONEDRIVE_ENABLED = _raw_enabled.lower() in {"1", "true", "sim", "yes", "on"}
else:
	ONEDRIVE_ENABLED = bool(ONEDRIVE_TENANT_ID and ONEDRIVE_CLIENT_ID and ONEDRIVE_CLIENT_SECRET)
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


def _diagnostico_configuracao_sharepoint():
	"""Retorna a situação da configuração sem expor segredos."""
	usa_site_drive = bool(SHAREPOINT_SITE_ID and SHAREPOINT_DRIVE_ID)
	itens = {
		"SHAREPOINT_ENABLED": {"configurado": bool(SHAREPOINT_ENABLED), "obrigatorio": True, "valor_seguro": "true" if SHAREPOINT_ENABLED else "false"},
		"SHAREPOINT_TENANT_ID": {"configurado": bool(SHAREPOINT_TENANT_ID), "obrigatorio": True, "valor_seguro": "preenchido" if SHAREPOINT_TENANT_ID else "ausente"},
		"SHAREPOINT_CLIENT_ID": {"configurado": bool(SHAREPOINT_CLIENT_ID), "obrigatorio": True, "valor_seguro": "preenchido" if SHAREPOINT_CLIENT_ID else "ausente"},
		"SHAREPOINT_CLIENT_SECRET": {"configurado": bool(SHAREPOINT_CLIENT_SECRET), "obrigatorio": True, "valor_seguro": "preenchido" if SHAREPOINT_CLIENT_SECRET else "ausente"},
		"SHAREPOINT_USER": {"configurado": bool(ONEDRIVE_USER), "obrigatorio": not usa_site_drive, "valor_seguro": ONEDRIVE_USER if ONEDRIVE_USER else "ausente"},
		"SHAREPOINT_FOLDER_PATH": {"configurado": bool(SHAREPOINT_FOLDER_PATH), "obrigatorio": not usa_site_drive, "valor_seguro": SHAREPOINT_FOLDER_PATH},
		"SHAREPOINT_FILE_NAME": {"configurado": bool(SHAREPOINT_FILE_NAME), "obrigatorio": not usa_site_drive, "valor_seguro": SHAREPOINT_FILE_NAME},
		"SHAREPOINT_FILE_PATH": {"configurado": bool(SHAREPOINT_FILE_PATH), "obrigatorio": not usa_site_drive, "valor_seguro": SHAREPOINT_FILE_PATH},
		"SHAREPOINT_SITE_ID": {"configurado": bool(SHAREPOINT_SITE_ID), "obrigatorio": False, "valor_seguro": "preenchido" if SHAREPOINT_SITE_ID else "não utilizado"},
		"SHAREPOINT_DRIVE_ID": {"configurado": bool(SHAREPOINT_DRIVE_ID), "obrigatorio": False, "valor_seguro": "preenchido" if SHAREPOINT_DRIVE_ID else "não utilizado"},
	}
	faltantes = [nome for nome, item in itens.items() if item["obrigatorio"] and not item["configurado"]]
	if SHAREPOINT_SITE_ID and not SHAREPOINT_DRIVE_ID:
		faltantes.append("SHAREPOINT_DRIVE_ID (necessário se SHAREPOINT_SITE_ID for usado)")
	if SHAREPOINT_DRIVE_ID and not SHAREPOINT_SITE_ID:
		faltantes.append("SHAREPOINT_SITE_ID (necessário se SHAREPOINT_DRIVE_ID for usado)")
	return {
		"habilitada": bool(SHAREPOINT_ENABLED),
		"usa_site_drive": usa_site_drive,
		"usa_onedrive_usuario": not usa_site_drive,
		"itens": itens,
		"faltantes": faltantes,
		"configurada": bool(SHAREPOINT_ENABLED) and not faltantes,
	}


def _onedrive_configurado():
	# Suporta tanto SharePoint (site/drive) quanto OneDrive for Business do usuário.
	return _diagnostico_configuracao_sharepoint()["configurada"]


def _onedrive_token():
	if not _onedrive_configurado():
		d = _diagnostico_configuracao_sharepoint()
		faltantes = ", ".join(d["faltantes"]) or "SHAREPOINT_ENABLED está desabilitado"
		raise RuntimeError(f"Integração SharePoint não configurada. Variável(is) ausente(s): {faltantes}.")
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


def _onedrive_path_metadata_url():
	"""URL de metadados para localizar a pasta/arquivo por caminho no Graph."""
	if SHAREPOINT_SITE_ID and SHAREPOINT_DRIVE_ID:
		caminho = quote(SHAREPOINT_FILE_PATH.strip("/"), safe="/")
		return (
			f"https://graph.microsoft.com/v1.0/sites/{quote(SHAREPOINT_SITE_ID, safe='')}"
			f"/drives/{quote(SHAREPOINT_DRIVE_ID, safe='')}/root:/{caminho}"
		)
	if not ONEDRIVE_USER:
		raise RuntimeError("SHAREPOINT_USER/ONEDRIVE_USER é obrigatório para localizar o arquivo no OneDrive.")
	usuario = quote(ONEDRIVE_USER, safe="")
	caminho = quote(SHAREPOINT_FILE_PATH.strip("/"), safe="/")
	return f"https://graph.microsoft.com/v1.0/users/{usuario}/drive/root:/{caminho}"


def _onedrive_folder_metadata_url():
	"""URL de metadados da pasta do projeto no OneDrive."""
	if SHAREPOINT_SITE_ID and SHAREPOINT_DRIVE_ID:
		caminho = quote(SHAREPOINT_FOLDER_PATH.strip("/"), safe="/")
		return (
			f"https://graph.microsoft.com/v1.0/sites/{quote(SHAREPOINT_SITE_ID, safe='')}"
			f"/drives/{quote(SHAREPOINT_DRIVE_ID, safe='')}/root:/{caminho}"
		)
	if not ONEDRIVE_USER:
		raise RuntimeError("SHAREPOINT_USER/ONEDRIVE_USER é obrigatório para localizar a pasta.")
	usuario = quote(ONEDRIVE_USER, safe="")
	caminho = quote(SHAREPOINT_FOLDER_PATH.strip("/"), safe="/")
	return f"https://graph.microsoft.com/v1.0/users/{usuario}/drive/root:/{caminho}"


def _resolver_arquivo_graph(token):
	"""Localiza primeiro a pasta e depois o arquivo, retornando seus metadados."""
	base_headers = {"Authorization": f"Bearer {token}"}
	try:
		with urlopen(Request(_onedrive_folder_metadata_url(), headers=base_headers), timeout=30) as resposta:
			pasta = json.loads(resposta.read().decode("utf-8"))
	except HTTPError as erro:
		detalhe = erro.read().decode("utf-8", errors="replace")[:1000]
		if erro.code == 404:
			raise RuntimeError(
				f"Pasta '{SHAREPOINT_FOLDER_PATH}' não foi localizada no OneDrive de '{ONEDRIVE_USER}'."
			) from erro
		if erro.code in (401, 403):
			raise RuntimeError(
				f"Microsoft Graph recusou o acesso à pasta '{SHAREPOINT_FOLDER_PATH}' (HTTP {erro.code}). "
				"Verifique Files.ReadWrite.All e o consentimento administrativo."
			) from erro
		raise RuntimeError(f"Falha ao localizar a pasta no Microsoft Graph (HTTP {erro.code}): {detalhe}") from erro

	pasta_id = pasta.get("id")
	if not pasta_id:
		raise RuntimeError(f"O Graph localizou a pasta '{SHAREPOINT_FOLDER_PATH}', mas não retornou o ID dela.")

	# O endpoint por caminho abaixo resolve o arquivo dentro da pasta encontrada.
	try:
		with urlopen(Request(_onedrive_path_metadata_url(), headers=base_headers), timeout=30) as resposta:
			arquivo = json.loads(resposta.read().decode("utf-8"))
	except HTTPError as erro:
		detalhe = erro.read().decode("utf-8", errors="replace")[:1000]
		if erro.code == 404:
			raise RuntimeError(
				f"Arquivo '{SHAREPOINT_FILE_NAME}' não foi localizado dentro da pasta "
				f"'{SHAREPOINT_FOLDER_PATH}'."
			) from erro
		if erro.code in (401, 403):
			raise RuntimeError(
				f"Microsoft Graph recusou o acesso ao arquivo '{SHAREPOINT_FILE_NAME}' (HTTP {erro.code})."
			) from erro
		raise RuntimeError(f"Falha ao localizar o arquivo no Microsoft Graph (HTTP {erro.code}): {detalhe}") from erro

	arquivo_id = arquivo.get("id")
	if not arquivo_id:
		raise RuntimeError(f"O Graph localizou '{SHAREPOINT_FILE_NAME}', mas não retornou o ID do arquivo.")
	return pasta, arquivo


def _onedrive_url():
	"""Retorna a URL de conteúdo do arquivo, resolvendo pasta e arquivo pelo Graph."""
	if SHAREPOINT_SITE_ID and SHAREPOINT_DRIVE_ID:
		caminho = quote(SHAREPOINT_FILE_PATH.strip("/"), safe="/")
		return (
			f"https://graph.microsoft.com/v1.0/sites/{quote(SHAREPOINT_SITE_ID, safe='')}"
			f"/drives/{quote(SHAREPOINT_DRIVE_ID, safe='')}/root:/{caminho}:/content"
		)
	token = _onedrive_token()
	_, arquivo = _resolver_arquivo_graph(token)
	usuario = quote(ONEDRIVE_USER, safe="")
	return f"https://graph.microsoft.com/v1.0/users/{usuario}/drive/items/{quote(arquivo['id'], safe='')}/content"


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
	"""Obtém os metadados do arquivo remoto no SharePoint/OneDrive."""
	token = _onedrive_token()
	if SHAREPOINT_SITE_ID and SHAREPOINT_DRIVE_ID:
		url = _onedrive_path_metadata_url()
		with urlopen(Request(url, headers={"Authorization": f"Bearer {token}"}), timeout=30) as resposta:
			return json.loads(resposta.read().decode("utf-8"))
	_, arquivo = _resolver_arquivo_graph(token)
	return arquivo


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
	requisicao = Request(_onedrive_url(), data=conteudo, method="PUT", headers=headers)
	try:
		with urlopen(requisicao, timeout=120) as resposta:
			return resposta.status
	except HTTPError as erro:
		detalhe = erro.read().decode("utf-8", errors="replace")[:1200]
		if erro.code == 412:
			raise ConcurrentUpdateError("O arquivo do SharePoint foi alterado por outra instância durante a gravação.") from erro
		if erro.code in (401, 403):
			raise RuntimeError("Microsoft Graph recusou a gravação. Verifique o consentimento administrativo e Files.ReadWrite.All.") from erro
		if erro.code == 404:
			raise RuntimeError(
				f"Arquivo '{SHAREPOINT_FILE_NAME}' não localizado para gravação dentro de "
				f"'{SHAREPOINT_FOLDER_PATH}'."
			) from erro
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
	
		# Demanda
		planilha = workbook["demandas"]
		cabecalho = [c.value for c in planilha[1]]
		indices = {nome: cabecalho.index(nome) + 1 for nome in TABELAS_EXCEL["demandas"] if nome in cabecalho}
		if len(indices) != len(TABELAS_EXCEL["demandas"]):
			raise RuntimeError("A aba demandas não possui todas as colunas esperadas.")
		planilha.append([demanda[coluna] for coluna in TABELAS_EXCEL["demandas"]])

		# Histórico
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
	"""
	Fluxo atômico/otimista para Vercel + Excel no SharePoint:
	1) lê a versão atual; 2) altera os dados; 3) gera XLSX;
	4) PUT com If-Match/eTag; 5) baixa novamente; 6) confirma a demanda.
	Se outra instância alterar o arquivo, o PUT retorna 412 e o processo
	recomeça sobre a versão mais nova, evitando lost update.
	"""
	if not _onedrive_configurado():
		raise RuntimeError("Integração SharePoint não está configurada.")

	for tentativa in range(1, tentativas + 1):
		with ARQUIVO_LOCK:
			meta, remoto_antes = _graph_snapshot()
			etag = meta["eTag"]
			# Reconstrói a conexão em memória a partir da versão que acabamos de ler.
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

			# Confirmação pós-PUT: lê exatamente o que ficou no SharePoint.
			remoto_depois = _graph_download_diagnostico()
			hash_remoto = hashlib.sha256(remoto_depois).hexdigest()
			if hash_remoto != hash_enviado:
				raise RuntimeError("O arquivo remoto ficou diferente do XLSX enviado.")
			if not _confirmar_demanda_no_xlsx(remoto_depois, demanda_id):
				raise RuntimeError("A demanda não foi encontrada no arquivo remoto após a gravação.")

			return {
				"status": "OK",
				"demanda_id": demanda_id,
				"tentativa": tentativa,
				"etag_antes": etag,
				"sha256": hash_remoto,
				"mensagem": "Demanda cadastrada e confirmada no SharePoint."
			}

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
		# Commit aqui confirma apenas a transação da conexão em memória.
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
	"""Executa qualquer escrita do sistema contra a versão atual do Excel.

	Cada tentativa lê o XLSX + eTag, aplica a mutação, envia com If-Match,
	e só atualiza o banco em memória depois de confirmar o conteúdo remoto.
	Em conflito 412, relê a versão nova e repete a operação.
	"""
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
		cabecalho = [c.value for c in next(planilha.iter_rows(min_row=1, max_row=1))]
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
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{{ titulo or 'SP ÁGUAS' }}</title>
<style>
:root{--navy:#063b57;--petrol:#087f9d;--cyan:#12a4c2;--bg:#f5f8fa;--card:#fff;--text:#20313b;--muted:#71808a;--line:#e2e9ed;--shadow:0 12px 32px rgba(14,49,67,.08);--radius:16px}
*{box-sizing:border-box}html{scroll-behavior:smooth}body{margin:0;background:var(--bg);color:var(--text);font-family:Inter,"Segoe UI",Arial,sans-serif;font-size:14px;line-height:1.5}
a{color:inherit}.app-shell{min-height:100vh}.sidebar{position:fixed;z-index:50;inset:0 auto 0 0;width:258px;padding:24px 14px;background:linear-gradient(180deg,#063b57 0%,#07516e 58%,#087f9d 100%);color:#fff;box-shadow:8px 0 28px rgba(5,42,62,.14)}
.brand{display:flex;align-items:center;gap:12px;padding:4px 12px 22px;border-bottom:1px solid rgba(255,255,255,.13)}.brand-mark{width:38px;height:38px;border-radius:12px;display:grid;place-items:center;background:rgba(255,255,255,.13);font-weight:800;letter-spacing:-1px}.brand-title{font-weight:800;font-size:18px;letter-spacing:.3px}.brand-sub{display:block;margin-top:2px;color:rgba(255,255,255,.68);font-size:10px;text-transform:uppercase;letter-spacing:1px}
.user-box{margin:18px 6px 12px;padding:12px;border-radius:12px;background:rgba(255,255,255,.08);font-size:12px;color:rgba(255,255,255,.82)}.user-box strong{display:block;color:#fff;font-size:13px;margin-bottom:2px}.menu{display:flex;flex-direction:column;gap:4px;margin-top:10px}.menu-label{padding:12px 12px 5px;color:rgba(255,255,255,.42);font-size:10px;text-transform:uppercase;letter-spacing:1.2px}.menu a{display:flex;align-items:center;gap:10px;margin:0;padding:10px 12px;border-radius:10px;color:rgba(255,255,255,.82);text-decoration:none;font-size:13px;transition:.18s ease}.menu a:hover,.menu a.ativo{background:rgba(255,255,255,.12);color:#fff;transform:translateX(2px)}.menu a.sair{margin-top:10px;color:#ffd9d5}.sidebar-footer{position:absolute;left:20px;right:20px;bottom:18px;color:rgba(255,255,255,.42);font-size:10px}
.main{margin-left:258px;min-height:100vh}.topbar{height:72px;padding:0 38px;background:rgba(255,255,255,.92);border-bottom:1px solid var(--line);display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:20;backdrop-filter:blur(10px)}.breadcrumb{color:var(--muted);font-size:12px}.page-title{font-size:16px;font-weight:700;color:var(--navy)}.content{max-width:1500px;margin:0 auto;padding:32px 38px 50px}
.page-head{display:flex;align-items:flex-end;justify-content:space-between;gap:20px;margin-bottom:24px}.page-head h1{margin:0;color:var(--navy);font-size:28px;letter-spacing:-.5px}.page-head p{margin:6px 0 0;color:var(--muted)}.actions{display:flex;gap:9px;flex-wrap:wrap}
.card,.metrica{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);box-shadow:var(--shadow);padding:22px;margin-bottom:20px}.card h1{margin:0 0 7px;color:var(--navy);font-size:25px}.card h2{margin:0 0 15px;color:var(--navy);font-size:18px}.card h3{color:var(--navy)}
.grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:16px}.metrica{position:relative;overflow:hidden;padding:20px}.metrica:after{content:"";position:absolute;right:-22px;top:-22px;width:78px;height:78px;border-radius:50%;background:rgba(8,127,157,.06)}.metrica h3{margin:0;color:var(--muted);font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px}.metrica strong{display:block;margin-top:8px;color:var(--navy);font-size:30px;line-height:1.1}
.stat-strip{display:flex;gap:10px;flex-wrap:wrap;margin-top:8px}.pill,.badge{display:inline-flex;align-items:center;gap:5px;border-radius:999px;padding:5px 9px;font-size:11px;font-weight:700}.pill{background:#eef5f7;color:#526a75}.badge-vencido{background:#fdebec;color:#b42318}.badge-critico{background:#fff0df;color:#a65300}.badge-proximo{background:#fff8d9;color:#856404}.badge-normal{background:#e8f6ed;color:#23743a}.badge-concluido{background:#e9eef2;color:#52636d}.badge-info{background:#e6f4f8;color:#08657d}
.btn,button{appearance:none;border:0;border-radius:9px;background:var(--petrol);color:#fff;text-decoration:none;padding:10px 15px;font:600 13px inherit;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;gap:7px;transition:.18s ease}.btn:hover,button:hover{background:#066b85;transform:translateY(-1px);box-shadow:0 7px 16px rgba(8,127,157,.18)}.btn-verde{background:#2e7d4b}.btn-verde:hover{background:#25683d}.btn-vermelho{background:#c43d38}.btn-vermelho:hover{background:#a9322e}.btn-cinza{background:#60727c}.btn-cinza:hover{background:#4e606a}.btn-outline{background:#fff;color:var(--petrol);border:1px solid #b9d5dc}.btn-outline:hover{background:#f2fafc;color:var(--petrol);box-shadow:none}
.toolbar{display:flex;gap:10px;align-items:center;justify-content:space-between;flex-wrap:wrap;margin-bottom:16px}.toolbar form{display:flex;gap:9px;flex:1;min-width:260px}.toolbar form input{margin:0}.toolbar-actions{display:flex;gap:8px;flex-wrap:wrap}
input,select,textarea{width:100%;padding:11px 12px;border:1px solid #d4e0e5;border-radius:9px;margin:6px 0 14px;background:#fbfdfe;color:var(--text);font:14px inherit;transition:.15s}input:focus,select:focus,textarea:focus{outline:0;border-color:var(--cyan);box-shadow:0 0 0 3px rgba(18,164,194,.12);background:#fff}textarea{min-height:105px;resize:vertical}label{display:block;margin-top:4px;color:#40525c;font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:.45px}.form-section{padding:18px 0;border-top:1px solid var(--line)}.form-section:first-child{padding-top:0;border-top:0}.form-section h3{margin:0 0 13px;font-size:14px}.form-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:0 18px}.form-grid.three{grid-template-columns:repeat(3,minmax(0,1fr))}.full{grid-column:1/-1}.form-actions{display:flex;gap:9px;align-items:center;margin-top:6px;padding-top:18px;border-top:1px solid var(--line)}
.table-wrap{overflow:auto;border:1px solid var(--line);border-radius:12px}table{width:100%;border-collapse:separate;border-spacing:0;background:#fff;min-width:850px}th{position:sticky;top:0;background:#f7fafb;color:#61737d;padding:12px 11px;text-align:left;font-size:10px;text-transform:uppercase;letter-spacing:.55px;border-bottom:1px solid var(--line)}td{padding:13px 11px;border-bottom:1px solid #edf1f3;font-size:12px;vertical-align:middle}tr:last-child td{border-bottom:0}tbody tr:hover td{background:#fbfdfe}.table-title{font-weight:700;color:var(--navy)}
.alerta{padding:15px 17px;border:1px solid var(--line);border-left:4px solid #9aa9b0;border-radius:11px;margin:10px 0;background:#fff;box-shadow:0 3px 10px rgba(20,50,65,.04)}.vencido{border-left-color:#c43d38;background:#fff9f9}.critico{border-left-color:#e36c08;background:#fffaf4}.proximo{border-left-color:#c18a05;background:#fffdf3}.normal{border-left-color:#2e7d4b;background:#f9fcfa}.info{background:#eef8fb;color:#0a637a}.erro{background:#fff0f0;color:#a3221d;padding:12px 14px;border:1px solid #f2c9c7;border-radius:9px;margin-bottom:14px}
.empty{padding:42px 20px;text-align:center;color:var(--muted)}.empty strong{display:block;color:var(--navy);font-size:16px;margin-bottom:4px}.muted{color:var(--muted)}.success{background:#edf9f1;color:#216b39;border:1px solid #cbe9d4;padding:12px 14px;border-radius:9px}.danger-zone{border-color:#f0d4d2}.danger-zone h3{color:#a9322e}
.login-page{min-height:100vh;display:grid;place-items:center;padding:24px;background:radial-gradient(circle at 15% 10%,rgba(18,164,194,.15),transparent 35%),linear-gradient(135deg,#eef6f8,#f8fafb)}.login-card{width:min(460px,100%);background:#fff;border:1px solid var(--line);border-radius:22px;box-shadow:0 22px 60px rgba(10,48,65,.12);padding:36px}.login-brand{text-align:center;margin-bottom:28px}.login-brand .brand-mark{margin:0 auto 12px;background:#e8f5f8;color:var(--navy)}.login-brand h1{margin:0;color:var(--navy);font-size:26px}.login-brand p{margin:5px 0;color:var(--muted);font-size:12px}
.popup-alertas{display:none;position:fixed;inset:0;z-index:80;align-items:center;justify-content:center;padding:20px;background:rgba(4,26,39,.52)}.popup-alertas.visivel{display:flex;animation:aparecer .2s ease-out}.popup-alertas-conteudo{position:relative;width:min(520px,100%);padding:30px;border:1px solid #f0d6a6;border-radius:18px;background:#fffdf8;box-shadow:0 18px 50px rgba(7,30,45,.25)}.popup-alertas-conteudo h2{margin:0 0 8px;color:#7b3f00}.popup-alertas-conteudo ul{max-height:230px;margin:18px 0;padding-left:20px;color:#4b3b2a}.popup-alertas-conteudo li{margin:9px 0}.popup-fechar{position:absolute;top:10px;right:12px;padding:2px 9px;background:transparent;color:#6d6258;font-size:27px}.popup-icone{display:grid;width:34px;height:34px;margin-bottom:12px;place-items:center;border-radius:50%;background:#c62828;color:#fff;font-size:22px;font-weight:700}@keyframes aparecer{from{opacity:0;transform:scale(.97)}to{opacity:1;transform:scale(1)}}
@media(max-width:1050px){.grid{grid-template-columns:repeat(2,minmax(0,1fr))}.form-grid.three{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media(max-width:760px){.sidebar{position:relative;width:100%;min-height:auto;padding:15px}.brand{padding-bottom:15px}.sidebar-footer{display:none}.user-box{display:none}.menu{flex-direction:row;flex-wrap:wrap}.menu-label{display:none}.menu a{padding:8px 9px}.main{margin-left:0}.topbar{height:58px;padding:0 16px}.content{padding:22px 14px 40px}.page-head{align-items:flex-start;flex-direction:column}.page-head h1{font-size:24px}.grid{grid-template-columns:1fr 1fr}.form-grid,.form-grid.three{grid-template-columns:1fr}.full{grid-column:auto}.toolbar form{min-width:100%}.login-card{padding:28px 22px}}
@media(max-width:460px){.grid{grid-template-columns:1fr}.content{padding-left:10px;padding-right:10px}.card,.metrica{padding:17px}.topbar .breadcrumb{display:none}}
</style>
</head>
<body>
{% if session.get('usuario_id') %}
<div class="app-shell">
<aside class="sidebar">
  <div class="brand"><div class="brand-mark">SA</div><div><div class="brand-title">SP ÁGUAS</div><span class="brand-sub">Gestão de demandas</span></div></div>
  <div class="user-box"><strong>{{ session.get('nome','Usuário') }}</strong>{{ session.get('perfil','Usuário') }}</div>
  <nav class="menu">
    <div class="menu-label">Principal</div>
    <a class="{% if request.path == '/sistema' %}ativo{% endif %}" href="/sistema">▦ &nbsp;Dashboard</a>
    <a class="{% if request.path == '/demandas' %}ativo{% endif %}" href="/demandas">☷ &nbsp;Demandas</a>
    <a class="{% if request.path == '/nova-demanda' %}ativo{% endif %}" href="/nova-demanda">＋ &nbsp;Nova demanda</a>
    <div class="menu-label">Controle</div>
    <a class="{% if request.path == '/tce' %}ativo{% endif %}" href="/tce">▣ &nbsp;TCE-SP</a>
    <a class="{% if request.path == '/audesp' %}ativo{% endif %}" href="/audesp">◫ &nbsp;AUDESP</a>
    <a class="{% if request.path == '/alertas' %}ativo{% endif %}" href="/alertas">⚠ &nbsp;Alertas</a>
    <a class="{% if request.path == '/calendario' %}ativo{% endif %}" href="/calendario">□ &nbsp;Calendário</a>
    <div class="menu-label">Sistema</div>
    <a href="/exportar">⇩ &nbsp;Exportar</a>
    {% if session.get('perfil') == 'Administrador' %}<a class="{% if request.path == '/usuarios' %}ativo{% endif %}" href="/usuarios">♙ &nbsp;Usuários</a><a href="/status">◉ &nbsp;Status</a>{% endif %}
    <a class="sair" href="/logout">↪ &nbsp;Sair</a>
  </nav>
  <div class="sidebar-footer">SP ÁGUAS • Sistema interno</div>
</aside>
<div class="main"><div class="topbar"><div class="breadcrumb">SP ÁGUAS / <span class="page-title">{{ titulo }}</span></div><div class="muted">{{ session.get('perfil','') }}</div></div><main class="content">{{ conteudo|safe }}</main></div>
</div>
{% else %}
<div class="login-page">{{ conteudo|safe }}</div>
{% endif %}
</body></html>
"""


def pagina(titulo, conteudo):
    return render_template_string(HTML_BASE, titulo=titulo, conteudo=conteudo + popup_alertas())


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
	# Verificação somente leitura da versão remota atual.
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
			return pagina("Login", '<div class="login-card"><div class="login-brand"><div class="brand-mark">SA</div><h1>SP ÁGUAS</h1><p>Gestão de Processos e Prazos</p></div><div class="erro">Usuário ou senha inválidos, ou acesso ainda não liberado.</div><a class="btn" style="width:100%" href="/">Voltar ao login</a></div>')
		session.update(usuario_id=user["id"], nome=user["nome"], usuario=user["usuario"], perfil=user["perfil"])
		return redirect(url_for("sistema"))
	conn.close()
	return pagina("Login", '<div class="login-card"><div class="login-brand"><div class="brand-mark">SA</div><h1>SP ÁGUAS</h1><p>Gestão de Processos e Prazos</p></div><form method="post"><label>Usuário</label><input name="usuario" autocomplete="username" required><label>Senha</label><input type="password" name="senha" autocomplete="current-password" required><button style="width:100%;margin-top:4px">Entrar</button></form><div style="text-align:center;margin-top:18px"><span class="muted">Ainda não possui acesso?</span><br><a href="/cadastro" class="btn btn-outline" style="margin-top:9px">Criar minha conta</a></div></div>')


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
	return pagina("Cadastro", '<div class="login-card"><div class="login-brand"><div class="brand-mark">SA</div><h1>Criar acesso</h1><p>Cadastre seus dados para solicitar acesso ao SP ÁGUAS.</p></div><form method="post"><label>Nome completo</label><input name="nome" required><label>Usuário</label><input name="usuario" required><label>E-mail</label><input type="email" name="email" required><label>Senha</label><input type="password" name="senha" minlength="8" required><label>Confirmar senha</label><input type="password" name="confirmar" minlength="8" required><button style="width:100%">Criar minha conta</button></form><div style="text-align:center;margin-top:16px"><a href="/" class="btn btn-outline">Voltar ao login</a></div></div>')


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
    contagens = {"total": len(registros), "tce": 0, "audesp": 0, "concluidas": 0, "VENCIDO": 0, "CRÍTICO": 0, "PRÓXIMO": 0, "NORMAL": 0}
    calculados = []
    for d in registros:
        if d["origem"] == "TCE-SP": contagens["tce"] += 1
        if d["origem"] == "AUDESP": contagens["audesp"] += 1
        if d["situacao"] == "Concluído": contagens["concluidas"] += 1
        p = calcular_prazos(d["prazo_area"], d["prazo_fatal"]); calculados.append((d, p))
        if d["situacao"] != "Concluído": contagens[p["status"]] += 1
    contagens["ativos"] = sum(d["situacao"] != "Concluído" for d in registros)
    ativos = [item for item in calculados if item[0]["situacao"] != "Concluído" and item[1]["dias_fatal"] is not None]; ativos.sort(key=lambda item: item[1]["dias_fatal"])
    cards = [("Total de demandas", "total"), ("Em andamento", "ativos"), ("Vencidas", "VENCIDO"), ("Próximas", "PRÓXIMO")]
    cards_html = ''.join(f'<div class="metrica"><h3>{nome}</h3><strong>{contagens[chave]}</strong></div>' for nome, chave in cards)
    linhas = ''
    for d, p in ativos[:8]:
        classe = {"VENCIDO":"badge-vencido", "CRÍTICO":"badge-critico", "PRÓXIMO":"badge-proximo", "NORMAL":"badge-normal"}.get(p["status"], "badge-info")
        dias = p["dias_fatal"]; texto_dias = "vencido" if dias < 0 else "hoje" if dias == 0 else f"{dias} dia(s)"
        linhas += f'<div class="alerta {p["status"].lower()}"><div style="display:flex;justify-content:space-between;gap:12px;align-items:flex-start;flex-wrap:wrap"><div><span class="badge {classe}">{p["status"]}</span><div class="table-title" style="margin-top:7px">{html.escape(d["numero_processo"] or "Sem número")} · {html.escape(d["assunto"][:100])}</div><div class="muted" style="margin-top:3px">Prazo fatal: <strong>{data_br(d["prazo_fatal"])}</strong> · {texto_dias}</div></div><a class="btn btn-outline" href="/editar/{d['id']}">Abrir demanda</a></div></div>'
    if not linhas: linhas = '<div class="empty"><strong>Nenhum prazo pendente</strong>Todas as demandas estão concluídas ou sem prazo fatal.</div>'
    html_dashboard = f'<div class="page-head"><div><h1>Dashboard</h1><p>Visão geral das demandas e dos prazos do SP ÁGUAS.</p></div><div class="actions"><a class="btn" href="/nova-demanda">＋ Nova demanda</a><a class="btn btn-outline" href="/demandas">Ver todas</a></div></div><div class="grid">{cards_html}</div><div class="card"><div style="display:flex;justify-content:space-between;gap:12px;align-items:center;flex-wrap:wrap"><div><h2>Próximos prazos</h2><p class="muted" style="margin:0">Demandas que exigem acompanhamento prioritário.</p></div><div class="stat-strip"><span class="pill">TCE-SP: {contagens["tce"]}</span><span class="pill">AUDESP: {contagens["audesp"]}</span><span class="pill">Concluídas: {contagens["concluidas"]}</span></div></div>{linhas}</div>'
    return pagina("Dashboard", html_dashboard)


@app.route("/nova-demanda", methods=["GET", "POST"])
def nova_demanda():
	if (resposta := acesso_login()): return resposta
	if request.method == "POST":
		valores = ler_demanda_form(); erro = validar_demanda(valores)
		if erro:
			return pagina("Nova demanda", f'<div class="card"><div class="erro">{erro}</div></div>' + formulario())
		try:
			resultado = salvar_demanda_atomicamente(valores, session["usuario_id"], tentativas=3)
			if resultado.get("status") == "OK":
				return redirect(url_for("demandas"))
			raise RuntimeError("O cadastro não foi confirmado no SharePoint.")
		except Exception as erro:
			return pagina("Nova demanda", f'<div class="card"><div class="erro">Demanda NÃO confirmada no SharePoint. {html.escape(str(erro))}</div><a class="btn" href="/nova-demanda">Tentar novamente</a> <a class="btn btn-cinza" href="/demandas">Voltar</a></div>')
	return pagina("Nova demanda", formulario())


def formulario(demanda=None):
    valor = lambda nome: html.escape(str((demanda[nome] or "") if demanda else ""), quote=True)
    escolhido = lambda nome, opcao: "selected" if valor(nome) == opcao else ""
    titulo = "Editar demanda" if demanda else "Nova demanda"
    subtitulo = "Atualize os dados e mantenha o histórico da demanda." if demanda else "Cadastre uma nova demanda e acompanhe seus prazos em um único lugar."
    situacoes = ("Aberto","Em análise","Aguardando área","Aguardando documento","Respondido","Concluído")
    prioridades = ("Normal","Alta","Urgente")
    situacao_options = ''.join(f'<option {escolhido("situacao", o)}>{o}</option>' for o in situacoes)
    prioridade_options = ''.join(f'<option {escolhido("prioridade", o)}>{o}</option>' for o in prioridades)
    historico_link = f'<a class="btn btn-outline" href="/historico/{demanda["id"]}">Ver histórico</a>' if demanda else ''
    return f'<div class="page-head"><div><h1>{titulo}</h1><p>{subtitulo}</p></div><a class="btn btn-outline" href="/demandas">← Voltar</a></div><div class="card"><form method="post"><div class="form-section"><h3>Identificação</h3><div class="form-grid"><div><label>Número do processo</label><input name="numero" value="{valor("numero_processo")}" placeholder="Ex.: TC-000000/000/00"></div><div><label>Origem</label><select name="origem"><option {escolhido("origem","TCE-SP")}>TCE-SP</option><option {escolhido("origem","AUDESP")}>AUDESP</option><option {escolhido("origem","Outro")}>Outro</option></select></div><div class="full"><label>Assunto</label><textarea name="assunto" required placeholder="Descreva de forma objetiva o assunto da demanda">{valor("assunto")}</textarea></div></div></div><div class="form-section"><h3>Responsabilidade</h3><div class="form-grid"><div><label>Área</label><input name="area" value="{valor("area")}" placeholder="Área responsável"></div><div><label>Responsável</label><input name="responsavel" value="{valor("responsavel")}" placeholder="Nome do responsável"></div></div></div><div class="form-section"><h3>Prazos e classificação</h3><div class="form-grid three"><div><label>Data de recebimento</label><input type="date" name="data_recebimento" value="{valor("data_recebimento")}"></div><div><label>Prazo da área</label><input type="date" name="prazo_area" value="{valor("prazo_area")}"></div><div><label>Prazo fatal</label><input type="date" name="prazo_fatal" value="{valor("prazo_fatal") }" required></div><div><label>Situação</label><select name="situacao">{situacao_options}</select></div><div><label>Prioridade</label><select name="prioridade">{prioridade_options}</select></div></div></div><div class="form-section"><h3>Observações</h3><textarea name="observacoes" placeholder="Informações complementares, providências ou observações internas">{valor("observacoes")}</textarea></div><div class="form-actions"><button type="submit">✓ {"Salvar alterações" if demanda else "Cadastrar demanda"}</button><a class="btn btn-cinza" href="/demandas">Cancelar</a>{historico_link}</div></form></div>'


@app.route("/demandas")
def demandas():
    if (resposta := acesso_login()): return resposta
    busca, origem = request.args.get("busca", "").strip(), request.args.get("origem", "")
    conn = conectar(); sql = "SELECT * FROM demandas WHERE 1=1"; params = []
    if busca: sql += " AND (numero_processo LIKE ? OR assunto LIKE ? OR area LIKE ? OR responsavel LIKE ?)"; params += [f"%{busca}%"] * 4
    if origem: sql += " AND origem = ?"; params.append(origem)
    registros = conn.execute(sql + " ORDER BY prazo_fatal", params).fetchall(); conn.close()
    linhas = ''
    for d in registros:
        p = calcular_prazos(d["prazo_area"], d["prazo_fatal"]); classe = {"VENCIDO":"badge-vencido","CRÍTICO":"badge-critico","PRÓXIMO":"badge-proximo","NORMAL":"badge-normal"}.get(p["status"],"badge-concluido") if d["situacao"] != "Concluído" else "badge-concluido"; status = "Concluído" if d["situacao"] == "Concluído" else p["status"]
        linhas += f'<tr><td><span class="table-title">{html.escape(d["numero_processo"] or "Sem número")}</span></td><td>{html.escape(d["origem"] or "-")}</td><td>{html.escape(d["assunto"] or "-")}</td><td>{html.escape(d["area"] or "-")}</td><td>{html.escape(d["responsavel"] or "-")}</td><td>{data_br(d["prazo_fatal"])}</td><td><span class="badge {classe}">{status}</span></td><td><a class="btn btn-outline" href="/editar/{d['id']}">Abrir</a></td></tr>'
    if not linhas: linhas = '<tr><td colspan="8"><div class="empty"><strong>Nenhuma demanda encontrada</strong>Ajuste os filtros ou cadastre uma nova demanda.</div></td></tr>'
    origem_opts = ''.join(f'<option value="{o}" {"selected" if origem==o else ""}>{o}</option>' for o in ("TCE-SP","AUDESP","Outro"))
    conteudo = f'<div class="page-head"><div><h1>Demandas</h1><p>Pesquise, filtre e acompanhe todas as demandas cadastradas.</p></div><a class="btn" href="/nova-demanda">＋ Nova demanda</a></div><div class="card"><div class="toolbar"><form method="get"><input name="busca" value="{html.escape(busca, quote=True)}" placeholder="Pesquisar por processo, assunto, área ou responsável"><button>Pesquisar</button></form><div class="toolbar-actions"><select name="origem"><option value="">Todas as origens</option>{origem_opts}</select><a class="btn btn-outline" href="/demandas">Limpar</a></div></div></div><div class="card"><div class="table-wrap"><table><thead><tr><th>Processo</th><th>Origem</th><th>Assunto</th><th>Área</th><th>Responsável</th><th>Prazo fatal</th><th>Status</th><th>Ações</th></tr></thead><tbody>{linhas}</tbody></table></div></div>'
    return pagina("Demandas", conteudo)


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
    itens = []
    for d in registros:
        p = calcular_prazos(d["prazo_area"], d["prazo_fatal"])
        if p["status"] != "NORMAL":
            classe = {"VENCIDO":"badge-vencido","CRÍTICO":"badge-critico","PRÓXIMO":"badge-proximo"}.get(p["status"],"badge-info"); dias=p["dias_fatal"]; texto="prazo vencido" if dias is not None and dias<0 else "vence hoje" if dias==0 else f"vence em {dias} dia(s)"
            itens.append(f'<div class="alerta {p["status"].lower()}"><div style="display:flex;justify-content:space-between;gap:15px;align-items:center;flex-wrap:wrap"><div><span class="badge {classe}">{p["status"]}</span><div class="table-title" style="margin-top:7px">{html.escape(d["numero_processo"] or "Sem número")} · {html.escape(d["assunto"])}</div><div class="muted">Prazo fatal: <strong>{data_br(d["prazo_fatal"])}</strong> · {texto}</div></div><a class="btn" href="/editar/{d['id']}">Abrir demanda</a></div></div>')
    if not itens: itens=['<div class="empty"><strong>Sem alertas no momento</strong>Não há demandas vencidas, críticas ou próximas do prazo.</div>']
    return pagina("Alertas", f'<div class="page-head"><div><h1>Alertas de prazos</h1><p>Priorize as demandas que exigem ação ou acompanhamento imediato.</p></div></div><div class="card">{"".join(itens)}</div>')


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
				ws = wb["usuarios"]; cab = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]; idx = cab.index("id"); ia = cab.index("ativo"); ip = cab.index("aprovado")
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


@app.route("/historico/<int:id>")
def historico(id):
    if (resposta := acesso_login()): return resposta
    conn = conectar(); demanda = conn.execute("SELECT * FROM demandas WHERE id = ?", (id,)).fetchone()
    if not demanda: conn.close(); return "Demanda não encontrada.", 404
    registros = conn.execute("SELECT h.*, u.nome AS usuario_nome FROM historico h LEFT JOIN usuarios u ON u.id = h.usuario_id WHERE h.demanda_id = ? ORDER BY h.data_hora DESC, h.id DESC", (id,)).fetchall(); conn.close()
    itens=''.join(f'<div class="alerta normal"><div style="display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap"><strong>{html.escape(r["acao"] or "ATUALIZAÇÃO")}</strong><span class="muted">{html.escape(str(r["data_hora"] or ""))}</span></div><div style="margin-top:6px">{html.escape(r["descricao"] or "")}</div><div class="muted" style="margin-top:5px">Por: {html.escape(r["usuario_nome"] or "Usuário")}</div></div>' for r in registros)
    if not itens: itens='<div class="empty"><strong>Sem movimentações</strong>Esta demanda ainda não possui registros no histórico.</div>'
    return pagina("Histórico", f'<div class="page-head"><div><h1>Histórico da demanda</h1><p>{html.escape(demanda["numero_processo"] or "Sem número")} · {html.escape(demanda["assunto"])}</p></div><a class="btn btn-outline" href="/editar/{id}">← Voltar à demanda</a></div><div class="card">{itens}</div>')


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
		diagnostico = _diagnostico_configuracao_sharepoint()
		resultado.update(
			status="ERRO",
			erro=f"Integração Microsoft Graph não configurada. Faltantes: {', '.join(diagnostico['faltantes']) or 'SHAREPOINT_ENABLED está desabilitado'}.",
			configuracao=diagnostico,
		)
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
	"""Teste somente leitura da configuração, autenticação Graph e Excel remoto."""
	if not usuario_logado():
		return jsonify({"status": "ERRO", "mensagem": "Usuário não autenticado."}), 401
	if not administrador():
		return jsonify({"status": "ERRO", "mensagem": "Somente administradores podem executar o teste."}), 403

	config = _diagnostico_configuracao_sharepoint()
	resultado = {
		"teste": "Integração SP ÁGUAS + Microsoft Graph + OneDrive/SharePoint",
		"status_final": "INICIANDO",
		"pasta": SHAREPOINT_FOLDER_PATH,
		"arquivo": SHAREPOINT_FILE_NAME,
		"caminho_completo": SHAREPOINT_FILE_PATH,
		"endpoint_graph": _onedrive_path_metadata_url() if config["configurada"] else None,
		"observacao": "Este teste é somente leitura e não altera o Excel.",
		"etapas": [],
	}

	def etapa(numero, nome, status, mensagem, **dados):
		item = {"numero": numero, "etapa": nome, "status": status, "mensagem": mensagem}
		item.update(dados)
		resultado["etapas"].append(item)

	etapa(1, "Configuração", "OK" if config["configurada"] else "ERRO",
		"Todas as variáveis obrigatórias estão preenchidas." if config["configurada"] else "A configuração está incompleta.",
		habilitada=config["habilitada"], faltantes=config["faltantes"],
		modo="Site/Drive" if config["usa_site_drive"] else "OneDrive do usuário",
		variaveis={k: v["valor_seguro"] for k, v in config["itens"].items()})

	if not config["configurada"]:
		resultado["status_final"] = "FALHA_CONFIGURACAO"
		resultado["mensagem_final"] = "Corrija as variáveis indicadas na etapa 1 no Vercel e faça um novo Deploy."
		return jsonify(resultado), 500

	try:
		token = _onedrive_token()
		etapa(2, "Autenticação Microsoft Graph", "OK", "Token de aplicação obtido com sucesso. O segredo não é exibido.")
	except Exception as erro:
		etapa(2, "Autenticação Microsoft Graph", "ERRO", str(erro))
		resultado["status_final"] = "FALHA_AUTENTICACAO"
		resultado["mensagem_final"] = "Verifique TENANT_ID, CLIENT_ID, CLIENT_SECRET e o consentimento administrativo das permissões Graph."
		return jsonify(resultado), 500

	try:
		pasta, meta_arquivo = _resolver_arquivo_graph(token)
		etapa(
			3, "Localização da pasta e do arquivo", "OK",
			"O Graph localizou a pasta e o arquivo pelo caminho informado.",
			pasta_id=pasta.get("id"),
			arquivo_id=meta_arquivo.get("id"),
			nome_arquivo=meta_arquivo.get("name"),
			caminho_graph=meta_arquivo.get("parentReference", {}).get("path"),
		)
		usuario = quote(ONEDRIVE_USER, safe="")
		content_url = (
			f"https://graph.microsoft.com/v1.0/users/{usuario}/drive/items/"
			f"{quote(meta_arquivo['id'], safe='')}/content"
			if not (SHAREPOINT_SITE_ID and SHAREPOINT_DRIVE_ID)
			else _onedrive_url()
		)
		with urlopen(Request(content_url, headers={"Authorization": f"Bearer {token}"}), timeout=60) as resposta:
			arquivo = resposta.read()
			status_http = resposta.status
		etapa(4, "Download e validação do Excel", "OK", "O arquivo remoto foi localizado e baixado.",
			status_http=status_http, tamanho_bytes=len(arquivo),
			sha256=hashlib.sha256(arquivo).hexdigest())
	except HTTPError as erro:
		detalhe = erro.read().decode("utf-8", errors="replace")[:1200]
		etapa(3, "Localização da pasta e do arquivo", "ERRO", f"Graph retornou HTTP {erro.code}: {detalhe}")
		resultado["status_final"] = "FALHA_ACESSO_GRAPH"
		resultado["mensagem_final"] = "Verifique credenciais/permissões e o nome exato da pasta e do arquivo."
		return jsonify(resultado), 500
	except Exception as erro:
		etapa(3, "Localização da pasta e do arquivo", "ERRO", str(erro))
		resultado["status_final"] = "FALHA_ACESSO_GRAPH"
		resultado["mensagem_final"] = "Verifique se a pasta SP_AGUAS e o arquivo Excel existem no OneDrive corporativo."
		return jsonify(resultado), 500

	try:
		workbook = load_workbook(io.BytesIO(arquivo), read_only=True, data_only=True)
		abas = workbook.sheetnames
		workbook.close()
		exigidas = ["usuarios", "demandas", "historico"]
		faltantes_abas = [aba for aba in exigidas if aba not in abas]
		if faltantes_abas:
			etapa(5, "Validação do Excel", "ERRO", "O arquivo foi acessado, mas faltam abas obrigatórias.", abas=abas, faltantes=faltantes_abas)
			resultado["status_final"] = "FALHA_ESTRUTURA_EXCEL"
			return jsonify(resultado), 500
		etapa(5, "Validação do Excel", "OK", "XLSX válido e abas obrigatórias encontradas.", abas=abas)
	except Exception as erro:
		etapa(5, "Validação do Excel", "ERRO", str(erro))
		resultado["status_final"] = "FALHA_VALIDACAO_EXCEL"
		return jsonify(resultado), 500

	resultado["status_final"] = "OK"
	resultado["mensagem_final"] = "CONEXÃO MICROSOFT GRAPH + PASTA + ARQUIVO EXCEL OK. O teste foi concluído sem alterar o arquivo remoto."
	return jsonify(resultado), 200


@app.route("/health")
def health():
	return jsonify({"status": "ok", "service": "sp-aguas"}), 200


@app.route("/diagnostico-sync")
def diagnostico_sync():
    """Diagnóstico somente leitura da cadeia aplicação -> XLSX remoto.

    Esta rota nunca reconstrói nem envia o workbook ao SharePoint/OneDrive.
    """
    if (resposta := acesso_login()):
        return resposta
    if not administrador():
        return jsonify({"status": "FALHA", "erro": "Somente administradores podem executar o diagnóstico."}), 403

    inicio = datetime.now()
    config = _diagnostico_configuracao_sharepoint()
    resultado = {
        "status": "INICIANDO",
        "somente_leitura": True,
        "arquivo_configurado": SHAREPOINT_FILE_PATH,
        "etapas": [],
    }

    def etapa(numero, nome, status, mensagem, **dados):
        item = {"numero": numero, "etapa": nome, "status": status, "mensagem": mensagem}
        item.update(dados)
        resultado["etapas"].append(item)

    if not config["configurada"]:
        etapa(1, "Configuração", "ERRO", "Integração Microsoft Graph desabilitada ou incompleta.", faltantes=config["faltantes"])
        resultado.update(status="FALHA_CONFIGURACAO", mensagem_final="Preencha as variáveis do Microsoft Graph no Vercel e faça um novo Deploy.")
        return jsonify(resultado), 500
    etapa(1, "Configuração", "OK", "Configuração do Microsoft Graph está preenchida.", caminho=SHAREPOINT_FILE_PATH)

    try:
        token = _onedrive_token()
        etapa(2, "Autenticação Microsoft Graph", "OK", "Access token obtido. O segredo não é exibido.")
    except Exception as erro:
        etapa(2, "Autenticação Microsoft Graph", "ERRO", str(erro))
        resultado.update(status="FALHA_AUTENTICACAO", mensagem_final="Verifique credenciais, permissões Graph e consentimento administrativo.")
        return jsonify(resultado), 500

    try:
        pasta, arquivo = _resolver_arquivo_graph(token)
        etapa(3, "Localização", "OK", "Pasta e arquivo localizados pelo caminho configurado.", pasta_id=pasta.get("id"), arquivo_id=arquivo.get("id"), nome=arquivo.get("name"))
    except Exception as erro:
        etapa(3, "Localização", "ERRO", str(erro))
        resultado.update(status="FALHA_LOCALIZACAO", mensagem_final="Verifique a pasta SP_AGUAS e o nome exato do arquivo no OneDrive corporativo.")
        return jsonify(resultado), 500

    try:
        usuario = quote(ONEDRIVE_USER, safe="")
        if SHAREPOINT_SITE_ID and SHAREPOINT_DRIVE_ID:
            url_conteudo = _onedrive_url()
        else:
            url_conteudo = f"https://graph.microsoft.com/v1.0/users/{usuario}/drive/items/{quote(arquivo['id'], safe='')}/content"
        with urlopen(Request(url_conteudo, headers={"Authorization": f"Bearer {token}"}), timeout=60) as resposta:
            conteudo = resposta.read()
            status_http = resposta.status
        etapa(4, "Download", "OK", "Arquivo remoto baixado sem qualquer alteração.", status_http=status_http, tamanho_bytes=len(conteudo), sha256=hashlib.sha256(conteudo).hexdigest())
    except Exception as erro:
        etapa(4, "Download", "ERRO", str(erro))
        resultado.update(status="FALHA_DOWNLOAD", mensagem_final="O Graph encontrou o arquivo, mas não conseguiu baixá-lo.")
        return jsonify(resultado), 500

    try:
        workbook = load_workbook(io.BytesIO(conteudo), read_only=True, data_only=True)
        abas = list(workbook.sheetnames)
        contagens = {tabela: max(workbook[tabela].max_row - 1, 0) if tabela in abas else None for tabela in TABELAS_EXCEL}
        workbook.close()
        faltantes = [aba for aba in TABELAS_EXCEL if aba not in abas]
        if faltantes:
            raise RuntimeError("Faltam as abas obrigatórias: " + ", ".join(faltantes))
        etapa(5, "Estrutura do Excel", "OK", "XLSX válido e abas obrigatórias encontradas.", abas=abas, contagens_remotas=contagens)
    except Exception as erro:
        etapa(5, "Estrutura do Excel", "ERRO", str(erro))
        resultado.update(status="FALHA_ESTRUTURA_EXCEL", mensagem_final="O arquivo remoto não possui a estrutura esperada pelo sistema.")
        return jsonify(resultado), 500

    try:
        conn = conectar()
        contagens_locais = {tabela: conn.execute(f"SELECT COUNT(*) FROM {tabela}").fetchone()[0] for tabela in TABELAS_EXCEL}
        conn.close()
        etapa(6, "Dados carregados", "OK", "Dados da aplicação disponíveis em memória.", contagens_locais=contagens_locais, contagens_remotas=contagens)
    except Exception as erro:
        etapa(6, "Dados carregados", "ERRO", str(erro))
        resultado.update(status="FALHA_DADOS", mensagem_final="Não foi possível validar os dados carregados pela aplicação.")
        return jsonify(resultado), 500

    resultado.update(status="OK", duracao_segundos=round((datetime.now() - inicio).total_seconds(), 2), conclusao="Diagnóstico concluído em modo somente leitura. Nenhum upload ou alteração foi executado.")
    return jsonify(resultado), 200


@app.route("/logout")
def logout():
	session.clear(); return redirect(url_for("login"))


criar_banco()

if __name__ == "__main__":
	app.run(host="0.0.0.0", port=5000, debug=False)

