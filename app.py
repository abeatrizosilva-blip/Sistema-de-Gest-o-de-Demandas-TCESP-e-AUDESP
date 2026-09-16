from datetime import date, datetime, timedelta
import io
import html
import hashlib
import json
import os
import socket
import sqlite3
import tempfile
import time
import uuid
from threading import RLock
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode
from unicodedata import normalize as unicode_normalize
from urllib.request import Request, urlopen

import bcrypt
import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
try:
	from pypdf import PdfReader
except ImportError:
	PdfReader = None
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


# Integração: Vercel chama um fluxo Power Automate; o fluxo é quem acessa o Excel.
# Não são necessárias credenciais Microsoft Graph no Vercel.
POWER_AUTOMATE_URL = _env("POWER_AUTOMATE_URL")
POWER_AUTOMATE_SECRET = _env("POWER_AUTOMATE_SECRET")
try:
	POWER_AUTOMATE_TIMEOUT = int(_env("POWER_AUTOMATE_TIMEOUT", default="90"))
except ValueError:
	POWER_AUTOMATE_TIMEOUT = 90
_raw_pa_enabled = _env("POWER_AUTOMATE_ENABLED")
POWER_AUTOMATE_ENABLED = (_raw_pa_enabled.lower() in {"1", "true", "sim", "yes", "on"}) if _raw_pa_enabled else bool(POWER_AUTOMATE_URL and POWER_AUTOMATE_SECRET)

# Mantidos somente como metadados do arquivo, sem autenticação no Vercel.
SHAREPOINT_FOLDER_PATH = _env("SHAREPOINT_FOLDER_PATH", "ONEDRIVE_FOLDER_PATH", default="SP_AGUAS")
SHAREPOINT_FILE_NAME = _env("SHAREPOINT_FILE_NAME", "ONEDRIVE_FILE_NAME", default="Sistema de Gestão de Demandas - SP Aguas.xlsx")
SHAREPOINT_FILE_PATH = _env("SHAREPOINT_FILE_PATH", "ONEDRIVE_PATH", default=f"{SHAREPOINT_FOLDER_PATH.strip('/')}/{SHAREPOINT_FILE_NAME}")
ONEDRIVE_PATH = SHAREPOINT_FILE_PATH

TABELAS_EXCEL = {
	"usuarios": ("id", "nome", "usuario", "email", "senha_hash", "perfil", "ativo", "aprovado", "criado_em"),
	"demandas": ("id", "numero_processo", "numero_etc", "origem", "assunto", "area", "responsavel", "data_recebimento", "prazo_area", "prazo_fatal", "situacao", "prioridade", "observacoes", "doe_data", "doe_edicao", "doe_secao", "doe_palavra_chave", "doe_publicacao", "doe_url", "criado_por", "criado_em", "atualizado_em"),
	"historico": ("id", "demanda_id", "usuario_id", "acao", "descricao", "data_hora"),
}


def _diagnostico_configuracao_sharepoint():
	"""Compatibilidade: o diagnóstico agora descreve a integração Power Automate."""
	return _diagnostico_configuracao_power_automate()


def _power_automate_configurado():
	"""Indica se o Vercel está configurado para falar somente com o Power Automate."""
	return bool(POWER_AUTOMATE_ENABLED and POWER_AUTOMATE_URL)


def _diagnostico_configuracao_power_automate():
	itens = {
		"POWER_AUTOMATE_ENABLED": {"configurado": bool(POWER_AUTOMATE_ENABLED), "obrigatorio": True, "valor_seguro": "true" if POWER_AUTOMATE_ENABLED else "false"},
		"POWER_AUTOMATE_URL": {"configurado": bool(POWER_AUTOMATE_URL), "obrigatorio": True, "valor_seguro": "preenchido" if POWER_AUTOMATE_URL else "ausente"},
		"POWER_AUTOMATE_SECRET": {"configurado": bool(POWER_AUTOMATE_SECRET), "obrigatorio": True, "valor_seguro": "preenchido" if POWER_AUTOMATE_SECRET else "ausente"},
		"SHAREPOINT_FOLDER_PATH": {"configurado": bool(SHAREPOINT_FOLDER_PATH), "obrigatorio": False, "valor_seguro": SHAREPOINT_FOLDER_PATH},
		"SHAREPOINT_FILE_NAME": {"configurado": bool(SHAREPOINT_FILE_NAME), "obrigatorio": False, "valor_seguro": SHAREPOINT_FILE_NAME},
	}
	faltantes = [nome for nome, item in itens.items() if item["obrigatorio"] and not item["configurado"]]
	return {
		"habilitada": bool(POWER_AUTOMATE_ENABLED),
		"configurada": bool(POWER_AUTOMATE_ENABLED and not faltantes),
		"faltantes": faltantes,
		"itens": itens,
		"modo": "Power Automate → Excel Online (Business) → SharePoint/OneDrive",
	}


def _onedrive_configurado():
	# Compatibilidade com o restante do código: agora significa Power Automate configurado.
	return _power_automate_configurado()


def _normalizar_valor_excel(valor):
	if isinstance(valor, datetime):
		return valor.strftime("%Y-%m-%d %H:%M:%S")
	if isinstance(valor, date):
		return valor.strftime("%Y-%m-%d")
	if isinstance(valor, bool):
		return 1 if valor else 0
	return valor


def _workbook_para_snapshot(conteudo):
	"""Converte o XLSX em um JSON compacto, sem enviar o arquivo para o Vercel."""
	workbook = load_workbook(io.BytesIO(conteudo), read_only=True, data_only=True)
	try:
		resultado = {"usuarios": [], "demandas": [], "historico": []}
		for tabela, colunas in TABELAS_EXCEL.items():
			if tabela not in workbook.sheetnames:
				continue
			ws = workbook[tabela]
			linhas = list(ws.iter_rows(values_only=True))
			if not linhas:
				continue
			cabecalho = [str(v).strip() if v is not None else "" for v in linhas[0]]
			indices = {nome: cabecalho.index(nome) for nome in colunas if nome in cabecalho}
			for linha in linhas[1:]:
				if not any(v is not None for v in linha):
					continue
				resultado[tabela].append({nome: _normalizar_valor_excel(linha[indices[nome]]) if nome in indices and indices[nome] < len(linha) else "" for nome in colunas})
		return resultado
	finally:
		workbook.close()


def _snapshot_para_xlsx(snapshot):
	workbook = Workbook()
	workbook.remove(workbook.active)
	for tabela, colunas in TABELAS_EXCEL.items():
		ws = workbook.create_sheet(tabela)
		ws.append(list(colunas))
		for celula in ws[1]:
			celula.font = Font(bold=True)
		for registro in snapshot.get(tabela, []):
			ws.append([registro.get(coluna, "") for coluna in colunas])
		ws.freeze_panes = "A2"
		ws.auto_filter.ref = ws.dimensions
		for coluna in ws.columns:
			largura = min(max(len(str(celula.value or "")) for celula in coluna) + 2, 45)
			ws.column_dimensions[coluna[0].column_letter].width = largura
	buffer = io.BytesIO()
	workbook.save(buffer)
	workbook.close()
	return buffer.getvalue()


def _snapshot_do_xlsx(conteudo):
	return _workbook_para_snapshot(conteudo)


def _power_automate_request(action, snapshot=None, extra=None):
	if not _power_automate_configurado():
		d = _diagnostico_configuracao_power_automate()
		faltantes = ", ".join(d["faltantes"]) or "POWER_AUTOMATE_ENABLED está desabilitado"
		raise RuntimeError(f"Integração Power Automate não configurada. Variável(is) ausente(s): {faltantes}.")
	payload = {
		"action": action,
		"requestId": str(uuid.uuid4()),
		"source": "SP_AGUAS",
		"timestamp": agora(),
	}
	if snapshot is not None:
		payload["snapshot"] = snapshot
	if extra:
		payload.update(extra)
	headers = {
		"Content-Type": "application/json",
		"Accept": "application/json",
		"X-SP-AGUAS-SECRET": POWER_AUTOMATE_SECRET,
	}
	try:
		with urlopen(Request(POWER_AUTOMATE_URL, data=json.dumps(payload, ensure_ascii=False).encode("utf-8"), method="POST", headers=headers), timeout=POWER_AUTOMATE_TIMEOUT) as resposta:
			texto = resposta.read().decode("utf-8", errors="replace")
			if not texto:
				return {"status": "OK"}
			resultado = json.loads(texto)
			if isinstance(resultado, dict) and resultado.get("status") in {"ERRO", "ERROR"}:
				if int(resultado.get("httpStatus") or resultado.get("statusCode") or 0) in (409, 412):
					raise ConcurrentUpdateError(resultado.get("mensagem") or "O Excel foi alterado durante a operação.")
				raise RuntimeError(resultado.get("mensagem") or resultado.get("error") or "Power Automate retornou erro.")
			return resultado if isinstance(resultado, dict) else {"status": "OK", "result": resultado}
	except HTTPError as erro:
		detalhe = erro.read().decode("utf-8", errors="replace")[:1800]
		raise RuntimeError(f"Power Automate retornou HTTP {erro.code}: {detalhe}") from erro
	except (URLError, TimeoutError, json.JSONDecodeError) as erro:
		raise RuntimeError(f"Não foi possível acessar o fluxo do Power Automate: {erro}") from erro


def _pa_snapshot_remoto():
	resposta = _power_automate_request("GET_SNAPSHOT")
	snapshot = resposta.get("snapshot") or resposta.get("data") or resposta.get("result")
	if isinstance(snapshot, str):
		try:
			snapshot = json.loads(snapshot)
		except json.JSONDecodeError as erro:
			raise RuntimeError("O Power Automate retornou um snapshot inválido.") from erro
	if not isinstance(snapshot, dict):
		raise RuntimeError("O Power Automate não retornou o snapshot do Excel.")
	return snapshot, resposta.get("version") or resposta.get("etag") or resposta.get("lastModified") or "power-automate"


def _graph_metadata_diagnostico():
	_, versao = _pa_snapshot_remoto()
	return {"eTag": str(versao), "version": str(versao), "name": SHAREPOINT_FILE_NAME, "source": "Power Automate"}


def _graph_download_diagnostico():
	snapshot, _ = _pa_snapshot_remoto()
	return _snapshot_para_xlsx(snapshot)


def _graph_upload_diagnostico(conteudo, etag=None):
	snapshot = _snapshot_do_xlsx(conteudo)
	resposta = _power_automate_request("REPLACE_SNAPSHOT", snapshot=snapshot, extra={"expectedVersion": etag or ""})
	status = resposta.get("httpStatus") or resposta.get("statusCode") or 200
	if isinstance(status, str) and status.isdigit():
		status = int(status)
	if status in (409, 412):
		raise ConcurrentUpdateError("O Excel foi alterado durante a operação pelo Power Automate.")
	if status not in (200, 201):
		raise RuntimeError(f"Power Automate não confirmou a gravação (status {status}).")
	return int(status)


def _graph_snapshot():
	meta = _graph_metadata_diagnostico()
	conteudo = _graph_download_diagnostico()
	if not conteudo:
		raise RuntimeError("O Power Automate retornou o Excel vazio.")
	return meta, conteudo


def obter_planilha():
	if not _power_automate_configurado():
		return None
	conteudo = _graph_download_diagnostico()
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
	raise RuntimeError("Gravação direta desabilitada. Use o fluxo do Power Automate.")


def _enviar_planilha_one_drive():
	return salvar_planilha()


def _onedrive_token():
	# Compatibilidade com funções legadas; nenhum token Microsoft é usado pelo Vercel.
	return "POWER_AUTOMATE"


def _onedrive_url():
	return POWER_AUTOMATE_URL


def _onedrive_path_metadata_url():
	return POWER_AUTOMATE_URL


def _onedrive_folder_metadata_url():
	return POWER_AUTOMATE_URL


def _resolver_arquivo_graph(token=None):
	return {"name": SHAREPOINT_FOLDER_PATH}, {"name": SHAREPOINT_FILE_NAME, "id": "power-automate"}


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
	"""Cadastra uma demanda usando o Power Automate como camada de persistência."""
	def mutator(workbook):
		planilha = workbook["demandas"]
		cabecalho = _cabecalho_planilha(planilha)
		demanda_id = _proximo_id(planilha, cabecalho)
		agora_valor = agora()
		registro = dict(zip(TABELAS_EXCEL["demandas"], (demanda_id, *valores, usuario_id, agora_valor, agora_valor)))
		_append_registro(planilha, TABELAS_EXCEL["demandas"], registro)
		ph = workbook["historico"]
		hid = _proximo_id(ph, _cabecalho_planilha(ph))
		historico = {"id": hid, "demanda_id": demanda_id, "usuario_id": usuario_id, "acao": "CRIACAO", "descricao": "Demanda cadastrada.", "data_hora": agora_valor}
		_append_registro(ph, TABELAS_EXCEL["historico"], historico)
		return {"demanda_id": demanda_id, "historico_id": hid}
	resultado = executar_mutacao_atomica(mutator, lambda conteudo, r: _confirmar_id_na_aba(conteudo, "demandas", r["demanda_id"]) and _confirmar_id_na_aba(conteudo, "historico", r["historico_id"]), tentativas=tentativas)
	return {"status": "OK", **resultado, "mensagem": "Demanda cadastrada e confirmada pelo Power Automate."}


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
			for linha in planilha.iter_rows(min_row=2, values_only=True):
				if not any(v is not None for v in linha):
					continue
				valores_linha = tuple(linha[indices[c]] if c in indices and indices[c] < len(linha) else "" for c in colunas)
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
			id INTEGER PRIMARY KEY AUTOINCREMENT, numero_processo TEXT, numero_etc TEXT,
			origem TEXT NOT NULL, assunto TEXT NOT NULL, area TEXT, responsavel TEXT,
			data_recebimento TEXT, prazo_area TEXT, prazo_fatal TEXT, situacao TEXT,
			prioridade TEXT, observacoes TEXT, doe_data TEXT, doe_edicao TEXT,
			doe_secao TEXT, doe_palavra_chave TEXT, doe_publicacao TEXT, doe_url TEXT,
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
	"""Executa a mutação localmente sobre um snapshot e devolve o XLSX ao Power Automate.

	O Vercel nunca autentica no Microsoft Graph e nunca envia o arquivo diretamente ao
	SharePoint. O fluxo Power Automate recebe o snapshot e grava o Excel por meio do
	Excel Online (Business)/Office Scripts.
	"""
	if not _power_automate_configurado():
		d = _diagnostico_configuracao_power_automate()
		faltantes = ", ".join(d["faltantes"]) or "POWER_AUTOMATE_ENABLED está desabilitado"
		raise RuntimeError(f"Integração Power Automate não configurada. Faltantes: {faltantes}.")
	ultimo_conflito = None
	for tentativa in range(1, tentativas + 1):
		with ARQUIVO_LOCK:
			meta, remoto_antes = _graph_snapshot()
			versao = meta.get("eTag")
			workbook = load_workbook(io.BytesIO(remoto_antes))
			try:
				resultado = mutator(workbook)
				novo_xlsx = _salvar_workbook_preservando_arquivo(workbook)
			finally:
				workbook.close()
			try:
				status_put = _graph_upload_diagnostico(novo_xlsx, etag=versao)
			except ConcurrentUpdateError as erro:
				ultimo_conflito = erro
				continue
			if status_put not in (200, 201):
				raise RuntimeError(f"Power Automate retornou status {status_put}.")
			remoto_depois = _graph_download_diagnostico()
			if hashlib.sha256(remoto_depois).hexdigest() != hashlib.sha256(novo_xlsx).hexdigest():
				raise RuntimeError("O Excel confirmado pelo Power Automate ficou diferente do conteúdo enviado.")
			if confirmador is not None and not confirmador(remoto_depois, resultado):
				raise RuntimeError("A operação foi enviada, mas não pôde ser confirmada no Excel remoto.")
			_atualizar_conexao_com_snapshot(remoto_depois)
			return {"status": "OK", **(resultado if isinstance(resultado, dict) else {"resultado": resultado}), "tentativa": tentativa}
	if ultimo_conflito:
		raise RuntimeError("O Excel foi alterado simultaneamente. Tente novamente.") from ultimo_conflito
	raise RuntimeError("Não foi possível concluir a gravação pelo Power Automate.")


def _append_registro(planilha, colunas, registro):
	cabecalho = _cabecalho_planilha(planilha)
	# Migração transparente: acrescenta as novas colunas ao Excel antigo.
	for nome in colunas:
		if nome not in cabecalho:
			planilha.cell(1, planilha.max_column + 1).value = nome
			cabecalho.append(nome)
	linha = planilha.max_row + 1
	for coluna, nome in enumerate(cabecalho, start=1):
		if nome in colunas:
			planilha.cell(linha, coluna).value = registro.get(nome, "")
	return linha


def _atualizar_registro(planilha, colunas, registro_id, registro):
	cabecalho = _cabecalho_planilha(planilha)
	for nome in colunas:
		if nome not in cabecalho:
			planilha.cell(1, planilha.max_column + 1).value = nome
			cabecalho.append(nome)
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
			id INTEGER PRIMARY KEY AUTOINCREMENT, numero_processo TEXT, numero_etc TEXT,
			origem TEXT NOT NULL, assunto TEXT NOT NULL, area TEXT, responsavel TEXT,
			data_recebimento TEXT, prazo_area TEXT, prazo_fatal TEXT, situacao TEXT,
			prioridade TEXT, observacoes TEXT, doe_data TEXT, doe_edicao TEXT,
			doe_secao TEXT, doe_palavra_chave TEXT, doe_publicacao TEXT, doe_url TEXT,
			criado_por INTEGER, criado_em TEXT NOT NULL, atualizado_em TEXT
		);
		CREATE TABLE IF NOT EXISTS historico (
			id INTEGER PRIMARY KEY AUTOINCREMENT, demanda_id INTEGER, usuario_id INTEGER,
			acao TEXT, descricao TEXT, data_hora TEXT NOT NULL
		);
	""")
	colunas_demandas = {linha[1] for linha in conn.execute("PRAGMA table_info(demandas)")}
	for coluna, definicao in (
		("numero_etc", "TEXT"), ("doe_data", "TEXT"), ("doe_edicao", "TEXT"),
		("doe_secao", "TEXT"), ("doe_palavra_chave", "TEXT"),
		("doe_publicacao", "TEXT"), ("doe_url", "TEXT"),
		("criado_por", "INTEGER"), ("atualizado_em", "TEXT")
	):
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
	return tuple(request.form.get(c, "").strip() for c in (
		"numero", "numero_etc", "origem", "assunto", "area", "responsavel",
		"data_recebimento", "prazo_area", "prazo_fatal", "situacao", "prioridade",
		"observacoes", "doe_data", "doe_edicao", "doe_secao", "doe_palavra_chave",
		"doe_publicacao", "doe_url"
	))


def validar_demanda(valores):
	campos = ("numero do processo", "número ETC", "origem", "assunto", "área", "responsável",
		"data de recebimento", "prazo da área", "prazo fatal", "situação", "prioridade",
		"observações", "data DOE", "edição DOE", "seção DOE", "palavra-chave DOE",
		"publicação DOE", "URL DOE")
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


DOE_KEYWORDS = [
    "AGUAS", "AGENCIA", "DEPARTAMENTO DE AGUAS E ENERGIA ELETRICA", "DAEE",
    "SP AGUAS", "AGENCIA DE AGUAS DO ESTADO DE SÃO PAULO", "RICARDO DARUIZ BORSARI",
    "ALCEU SEGAMARCHI", "FRANCISCO EDUARDO LODUCCA", "CAMILA ROCHA CUNHA VIANA",
    "PAOLA SANCHEZ VALLEJO DE MORAES FORJAZ", "Ana Paula Zubiaurre Brites",
    "Anderson Barboza Esteves", "Nelson de Campos Lima", "Adriano Rafael Arre­pia de Queiroz",
]
DOE_BASE_URL = "https://doe.tce.sp.gov.br/v/pdf/{ano:04d}/{mes:02d}/doe-tce-{data}.pdf"

def _normalizar_pesquisa_doe(texto):
    texto = unicode_normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode("ascii")
    return " ".join(texto.upper().split())

def _url_pdf_doe(data_iso):
    dt = datetime.strptime(data_iso, "%Y-%m-%d").date()
    return DOE_BASE_URL.format(ano=dt.year, mes=dt.month, data=dt.strftime("%Y-%m-%d"))

def _download_doe_pdf(url):
    """Baixa o PDF do DOE de forma resiliente.

    Alguns endpoints do doe.tce.sp.gov.br encerram a resposta HTTP antes do
    corpo completo quando acessados por ambientes serverless. Por isso usamos
    requests, desabilitamos compressão e validamos o marcador final do PDF,
    repetindo a operação quando o arquivo chega incompleto.
    """
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; SP-AGUAS/1.0)",
        "Accept": "application/pdf,application/octet-stream;q=0.9,*/*;q=0.8",
        "Accept-Encoding": "identity",
        "Connection": "close",
    }
    ultimo_erro = None
    for tentativa in range(1, 4):
        try:
            if requests is not None:
                with requests.get(url, headers=headers, timeout=(10, 60), stream=True, allow_redirects=True) as r:
                    if r.status_code == 404:
                        raise FileNotFoundError("PDF não encontrado")
                    r.raise_for_status()
                    partes = []
                    total = 0
                    for bloco in r.iter_content(chunk_size=64 * 1024):
                        if bloco:
                            partes.append(bloco)
                            total += len(bloco)
                    conteudo = b"".join(partes)
            else:
                req = Request(url, headers=headers)
                with urlopen(req, timeout=60) as resposta:
                    conteudo = resposta.read()

            if not conteudo.startswith(b"%PDF"):
                raise RuntimeError("O servidor não retornou um PDF válido.")
            if b"%%EOF" not in conteudo[-4096:]:
                raise RuntimeError(f"O PDF foi recebido incompleto ({len(conteudo)} bytes).")
            return conteudo
        except FileNotFoundError as exc:
            raise RuntimeError("A edição PDF não foi encontrada para essa data.") from exc
        except Exception as exc:
            ultimo_erro = exc
            if tentativa < 3:
                time.sleep(0.8 * tentativa)
    raise RuntimeError(f"Não foi possível baixar o PDF completo do DOE-TCESP após 3 tentativas: {ultimo_erro}")

def _buscar_doe_pdf(data_publicacao_iso):
    """Localiza o PDF pela data de PUBLICAÇÃO, não apenas pela data do arquivo.

    O TCESP informa no próprio PDF uma data de disponibilização e uma data de
    publicação. O nome do arquivo usa a data de disponibilização. Assim, para
    uma data de publicação informada pelo usuário, procuramos os dias anteriores
    (incluindo a própria data) até encontrar a edição cuja publicação coincida.
    """
    dt = datetime.strptime(data_publicacao_iso, "%Y-%m-%d").date()
    erros = []
    for atraso in range(0, 8):
        candidata = dt - timedelta(days=atraso)
        url = _url_pdf_doe(candidata.strftime("%Y-%m-%d"))
        try:
            conteudo = _download_doe_pdf(url)
            try:
                leitor = PdfReader(io.BytesIO(conteudo))
                cabecalho = " ".join((leitor.pages[0].extract_text() or "").split()) if leitor.pages else ""
            except Exception:
                cabecalho = ""
            data_pub = re.search(r"(?:Data de publicação|Publicação)\s*[:—-]?\s*(\d{2}/\d{2}/\d{4})", cabecalho, re.I)
            if data_pub:
                publicada = datetime.strptime(data_pub.group(1), "%d/%m/%Y").date()
                if publicada != dt:
                    erros.append(f"{candidata}: publicação {publicada.strftime('%d/%m/%Y')}")
                    continue
            return conteudo, url, candidata.strftime("%Y-%m-%d")
        except Exception as exc:
            erros.append(f"{candidata}: {exc}")
    raise RuntimeError("Não foi encontrada uma edição do DOE-TCESP correspondente à data de publicação informada. Verifique a data.\n" + "\n".join(erros[-3:]))

def _processos_no_texto(texto):
    encontrados = []
    for padrao in (r"\b(?:TC|TCESP|eTC|ETC)[-\s]?\d{1,8}[/.-]\d{1,4}[/.-]\d{2,4}\b", r"\b\d{5,8}/\d{2,4}\b"):
        encontrados.extend(re.findall(padrao, texto, flags=re.I))
    vistos=[]
    for item in encontrados:
        item=re.sub(r"\s+","",item)
        if item not in vistos: vistos.append(item)
    return vistos[:10]

def _extrair_ocorrencias_doe(conteudo, data_iso, url, palavra_chave=""):
    if PdfReader is None:
        raise RuntimeError("Dependência pypdf não instalada. Execute pip install -r requirements.txt.")
    try: reader=PdfReader(io.BytesIO(conteudo))
    except Exception as exc: raise RuntimeError(f"Não foi possível ler o PDF do DOE-TCESP: {exc}") from exc
    termos=[palavra_chave] if palavra_chave else DOE_KEYWORDS
    resultados=[]
    for pagina_num,page in enumerate(reader.pages,start=1):
        try: texto=page.extract_text() or ""
        except Exception: texto=""
        normalizado=_normalizar_pesquisa_doe(texto)
        if not normalizado: continue
        for termo in termos:
            alvo=_normalizar_pesquisa_doe(termo); pos=normalizado.find(alvo)
            if pos<0: continue
            inicio=max(0,pos-280); fim=min(len(texto),pos+len(termo)+420)
            trecho=" ".join(texto[inicio:fim].split())
            resultados.append({"data":data_iso,"edicao":"DOE-TCESP","secao":f"Página {pagina_num}","palavra_chave":termo,"trecho":trecho,"pagina":pagina_num,"url":url,"processos":_processos_no_texto(trecho)})
    return resultados[:100]

@app.route("/api/doe-tcesp/pesquisar")
def pesquisar_doe_tcesp():
    if (resposta := acesso_login()): return resposta
    data_iso=request.args.get("data","").strip(); palavra=request.args.get("palavra_chave","").strip()
    try: datetime.strptime(data_iso,"%Y-%m-%d")
    except ValueError: return jsonify({"erro":"Informe uma data válida no formato AAAA-MM-DD.","resultados":[]}),400
    if palavra and _normalizar_pesquisa_doe(palavra) not in {_normalizar_pesquisa_doe(k) for k in DOE_KEYWORDS}:
        return jsonify({"erro":"Palavra-chave não cadastrada.","resultados":[]}),400
    try:
        pdf,url,data_arquivo=_buscar_doe_pdf(data_iso); resultados=_extrair_ocorrencias_doe(pdf,data_iso,url,palavra)
        return jsonify({"data":data_iso,"data_arquivo":data_arquivo,"url":url,"resultados":resultados})
    except Exception as exc:
        return jsonify({"erro":str(exc),"resultados":[]}),502

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
    doe_keywords = json.dumps(DOE_KEYWORDS, ensure_ascii=False)
    return f"""<div class="page-head"><div><h1>{titulo}</h1><p>{subtitulo}</p></div><a class="btn btn-outline" href="/demandas">← Voltar</a></div>
<div class="card"><form method="post">
<div class="form-section"><h3>Identificação</h3><div class="form-grid">
<div><label>Número do processo</label><input name="numero" value="{valor("numero_processo")}" placeholder="Ex.: TC-000000/000/00"></div>
<div><label>Número ETC correspondente</label><input name="numero_etc" value="{valor("numero_etc")}" placeholder="Ex.: ETC-000000"></div>
<div><label>Origem</label><select name="origem"><option {escolhido("origem","TCE-SP")}>TCE-SP</option><option {escolhido("origem","AUDESP")}>AUDESP</option><option {escolhido("origem","Outro")}>Outro</option></select></div>
<div class="full"><label>Assunto</label><textarea name="assunto" required placeholder="Descreva de forma objetiva o assunto da demanda">{valor("assunto")}</textarea></div>
</div></div>
<div class="form-section"><h3>Publicação diária — DOE-TCESP</h3>
<p class="muted">Pesquise a edição oficial do Diário Oficial do TCESP por data e pelas palavras-chave cadastradas.</p>
<div class="form-grid">
<div><label>Data da publicação</label><input type="date" id="doe_pesquisa_data" value="{valor("doe_data")}"></div>
<div><label>Palavra-chave</label><select id="doe_pesquisa_keyword"><option value="">Todas as palavras-chave</option></select></div>
<div class="full"><button type="button" class="btn btn-outline" onclick="pesquisarDOE()">🔎 Pesquisar publicação do DOE-TCESP</button></div>
</div>
<div id="doe_status" class="muted" style="margin-top:10px"></div><div id="doe_resultados" style="margin-top:12px"></div>
<div class="form-grid" style="margin-top:12px">
<div><label>Data DOE</label><input name="doe_data" id="doe_data" value="{valor("doe_data")}" readonly></div>
<div><label>Edição DOE</label><input name="doe_edicao" id="doe_edicao" value="{valor("doe_edicao")}" placeholder="Ex.: edição diária"></div>
<div><label>Seção DOE</label><input name="doe_secao" id="doe_secao" value="{valor("doe_secao")}"></div>
<div><label>Palavra-chave encontrada</label><input name="doe_palavra_chave" id="doe_palavra_chave" value="{valor("doe_palavra_chave")}" readonly></div>
<div class="full"><label>Publicação / trecho localizado</label><textarea name="doe_publicacao" id="doe_publicacao" placeholder="O trecho da publicação selecionada aparecerá aqui.">{valor("doe_publicacao")}</textarea></div>
<div class="full"><label>Link oficial da publicação</label><input name="doe_url" id="doe_url" value="{valor("doe_url")}" readonly></div>
</div></div>
<div class="form-section"><h3>Responsabilidade</h3><div class="form-grid"><div><label>Área</label><input name="area" value="{valor("area")}" placeholder="Área responsável"></div><div><label>Responsável</label><input name="responsavel" value="{valor("responsavel")}" placeholder="Nome do responsável"></div></div></div>
<div class="form-section"><h3>Prazos e classificação</h3><div class="form-grid three">
<div><label>Data de recebimento</label><input type="date" name="data_recebimento" value="{valor("data_recebimento")}"></div>
<div><label>Prazo da área</label><input type="date" name="prazo_area" value="{valor("prazo_area")}"></div>
<div><label>Prazo fatal</label><input type="date" name="prazo_fatal" value="{valor("prazo_fatal")}" required></div>
<div><label>Situação</label><select name="situacao">{situacao_options}</select></div><div><label>Prioridade</label><select name="prioridade">{prioridade_options}</select></div>
</div></div>
<div class="form-section"><h3>Observações</h3><textarea name="observacoes" placeholder="Informações complementares, providências ou observações internas">{valor("observacoes")}</textarea></div>
<div class="form-actions"><button type="submit">✓ {"Salvar alterações" if demanda else "Cadastrar demanda"}</button><a class="btn btn-cinza" href="/demandas">Cancelar</a>{historico_link}</div>
</form></div>
<script>
const DOE_KEYWORDS = {doe_keywords};
(function() {{
  const select = document.getElementById("doe_pesquisa_keyword");
  DOE_KEYWORDS.forEach(k => {{ const o=document.createElement("option"); o.value=k; o.textContent=k; if(k===document.getElementById("doe_palavra_chave").value)o.selected=true; select.appendChild(o); }});
}})();
async function pesquisarDOE() {{
  const data=document.getElementById("doe_pesquisa_data").value, palavra=document.getElementById("doe_pesquisa_keyword").value;
  const status=document.getElementById("doe_status"), box=document.getElementById("doe_resultados");
  if(!data) {{ status.textContent="Informe a data da edição que deseja pesquisar."; return; }}
  status.textContent="Consultando a publicação oficial do DOE-TCESP..."; box.innerHTML="";
  try {{
    const r=await fetch("/api/doe-tcesp/pesquisar?data="+encodeURIComponent(data)+"&palavra_chave="+encodeURIComponent(palavra));
    const j=await r.json(); if(!r.ok) throw new Error(j.erro||"Não foi possível consultar o DOE.");
    if(!j.resultados.length) box.innerHTML='<div class="empty"><strong>Nenhuma ocorrência encontrada.</strong>Não foi localizada publicação com os filtros informados nessa edição.</div>';
    else {{
      box.innerHTML=j.resultados.map((item,i)=>`<div class="alerta normal" style="margin-top:8px"><strong>${{escapeHtml(item.palavra_chave)}}</strong><div class="muted">Página ${{item.pagina}} · ${{escapeHtml(item.data)}} · ${{escapeHtml(item.url)}}</div><div style="margin-top:6px">${{escapeHtml(item.trecho)}}</div>${{item.processos&&item.processos.length?'<div class="muted" style="margin-top:6px">Processo(s): '+escapeHtml(item.processos.join(", "))+'</div>':''}}<button type="button" class="btn" style="margin-top:9px" onclick="usarDOE(${{i}})">Usar esta publicação</button></div>`).join("");
      window._doeResultados=j.resultados; status.textContent=`${{j.resultados.length}} ocorrência(s) encontrada(s).`;
    }}
  }} catch(e) {{ status.textContent="Erro: "+e.message; }}
}}
function usarDOE(i) {{
  const item=window._doeResultados[i];
  document.getElementById("doe_data").value=item.data; document.getElementById("doe_edicao").value=item.edicao||"";
  document.getElementById("doe_secao").value=item.secao||""; document.getElementById("doe_palavra_chave").value=item.palavra_chave||"";
  document.getElementById("doe_publicacao").value=item.trecho||""; document.getElementById("doe_url").value=item.url||"";
  if(!document.querySelector('input[name="numero"]').value && item.processos&&item.processos.length) document.querySelector('input[name="numero"]').value=item.processos[0];
}}
function escapeHtml(v) {{ return String(v??"").replace(/[&<>"']/g,c=>({{"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}}[c])); }}
</script>"""



@app.route("/demandas")
def demandas():
    if (resposta := acesso_login()): return resposta
    busca, origem = request.args.get("busca", "").strip(), request.args.get("origem", "")
    conn = conectar(); sql = "SELECT * FROM demandas WHERE 1=1"; params = []
    if busca:
        sql += " AND (numero_processo LIKE ? OR numero_etc LIKE ? OR assunto LIKE ? OR area LIKE ? OR responsavel LIKE ? OR doe_palavra_chave LIKE ? OR doe_publicacao LIKE ?)"
        params += [f"%{busca}%"] * 7
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
	ws.append(["ID", "Processo", "ETC", "Origem", "Assunto", "Área", "Responsável", "Prazo área", "Prazo fatal", "Dias entre", "Dias restantes", "Situação", "Prioridade", "Observações", "DOE data", "DOE edição", "DOE seção", "DOE palavra-chave", "DOE publicação", "DOE URL"])
	for d in dados:
		p = calcular_prazos(d["prazo_area"], d["prazo_fatal"]); ws.append([d["id"], d["numero_processo"], d["numero_etc"], d["origem"], d["assunto"], d["area"], d["responsavel"], data_br(d["prazo_area"]), data_br(d["prazo_fatal"]), p["dias_entre"], p["dias_fatal"], d["situacao"], d["prioridade"], d["observacoes"], d["doe_data"], d["doe_edicao"], d["doe_secao"], d["doe_palavra_chave"], d["doe_publicacao"], d["doe_url"]])
	arquivo = io.BytesIO(); wb.save(arquivo); arquivo.seek(0); return send_file(arquivo, as_attachment=True, download_name="SP_AGUAS_Demandas.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/status")
def status():
	if (resposta := acesso_login()): return resposta
	config = _diagnostico_configuracao_power_automate()
	resultado = {"status": "OK", "power_automate_configurado": config["configurada"], "planilha_local": PLANILHA, "vercel": bool(os.environ.get("VERCEL")), "arquivo": SHAREPOINT_FILE_PATH}
	if not config["configurada"]:
		resultado.update(status="ERRO", erro=f"Integração Power Automate não configurada. Faltantes: {', '.join(config['faltantes']) or 'POWER_AUTOMATE_ENABLED está desabilitado'}", configuracao=config)
		return jsonify(resultado), 500
	try:
		snapshot, versao = _pa_snapshot_remoto()
		resultado.update(power_automate=True, excel_acessivel=True, versao=str(versao), tabelas={k: len(v) for k, v in snapshot.items()}, mensagem="CONEXÃO POWER AUTOMATE + EXCEL OK")
		return jsonify(resultado), 200
	except Exception as erro:
		resultado.update(status="ERRO", power_automate=False, excel_acessivel=False, erro=str(erro))
		return jsonify(resultado), 500


@app.route("/teste-integracao")
def teste_integracao():
	"""Diagnóstico somente leitura da integração Vercel → Power Automate → Excel."""
	if not usuario_logado():
		return jsonify({"status": "ERRO", "mensagem": "Usuário não autenticado."}), 401
	if not administrador():
		return jsonify({"status": "ERRO", "mensagem": "Somente administradores podem executar o teste."}), 403
	config = _diagnostico_configuracao_power_automate()
	resultado = {
		"teste": "Integração SP ÁGUAS + Power Automate + Excel Online (Business)",
		"status_final": "INICIANDO",
		"pasta": SHAREPOINT_FOLDER_PATH,
		"arquivo": SHAREPOINT_FILE_NAME,
		"caminho_completo": SHAREPOINT_FILE_PATH,
		"endpoint_power_automate": POWER_AUTOMATE_URL if config["configurada"] else None,
		"observacao": "Este teste é somente leitura e não altera o Excel.",
		"etapas": [],
	}
	def etapa(numero, nome, status_nome, mensagem, **dados):
		item = {"numero": numero, "etapa": nome, "status": status_nome, "mensagem": mensagem}; item.update(dados); resultado["etapas"].append(item)
	etapa(1, "Configuração Power Automate", "OK" if config["configurada"] else "ERRO",
		"URL e segredo do fluxo estão configurados." if config["configurada"] else "A configuração está incompleta.",
		habilitada=config["habilitada"], faltantes=config["faltantes"], variaveis={k: v["valor_seguro"] for k, v in config["itens"].items()})
	if not config["configurada"]:
		resultado["status_final"] = "FALHA_CONFIGURACAO"
		resultado["mensagem_final"] = "Preencha POWER_AUTOMATE_URL, POWER_AUTOMATE_SECRET e POWER_AUTOMATE_ENABLED no Vercel e faça novo Deploy."
		return jsonify(resultado), 500
	try:
		snapshot, versao = _pa_snapshot_remoto()
		etapa(2, "Power Automate", "OK", "O fluxo respondeu ao pedido de leitura.", versao=str(versao))
		etapa(3, "Excel no SharePoint/OneDrive", "OK", "O fluxo devolveu o snapshot do arquivo.", tabelas={k: len(v) for k, v in snapshot.items()})
		resultado["status_final"] = "SUCESSO"
		resultado["mensagem_final"] = "Vercel → Power Automate → Excel está funcionando."
		return jsonify(resultado), 200
	except Exception as erro:
		etapa(2, "Power Automate", "ERRO", str(erro))
		resultado["status_final"] = "FALHA_POWER_AUTOMATE"
		resultado["mensagem_final"] = "Verifique o fluxo, a conexão do Excel Online (Business), o segredo e as permissões do arquivo."
		return jsonify(resultado), 500

@app.route("/health")
def health():
	return jsonify({"status": "ok", "service": "sp-aguas"}), 200


@app.route("/diagnostico-sync")
def diagnostico_sync():
	"""Diagnóstico somente leitura da cadeia Vercel → Power Automate → Excel."""
	if (resposta := acesso_login()):
		return resposta
	if not administrador():
		return jsonify({"status": "FALHA", "erro": "Somente administradores podem executar o diagnóstico."}), 403
	inicio = datetime.now()
	config = _diagnostico_configuracao_power_automate()
	resultado = {"status": "INICIANDO", "somente_leitura": True, "arquivo_configurado": SHAREPOINT_FILE_PATH, "etapas": []}
	def etapa(numero, nome, status_nome, mensagem, **dados):
		item = {"numero": numero, "etapa": nome, "status": status_nome, "mensagem": mensagem}; item.update(dados); resultado["etapas"].append(item)
	if not config["configurada"]:
		etapa(1, "Configuração Power Automate", "ERRO", "Integração Power Automate incompleta.", faltantes=config["faltantes"])
		resultado.update(status="FALHA_CONFIGURACAO", mensagem_final="Preencha POWER_AUTOMATE_URL, POWER_AUTOMATE_SECRET e POWER_AUTOMATE_ENABLED no Vercel.")
		return jsonify(resultado), 500
	etapa(1, "Configuração Power Automate", "OK", "Fluxo configurado.", modo=config["modo"])
	try:
		snapshot, versao = _pa_snapshot_remoto()
		etapa(2, "Comunicação com o fluxo", "OK", "O Power Automate respondeu ao pedido GET_SNAPSHOT.", versao=str(versao))
	except Exception as erro:
		etapa(2, "Comunicação com o fluxo", "ERRO", str(erro))
		resultado.update(status="FALHA_POWER_AUTOMATE", mensagem_final="Verifique a URL, o segredo e o fluxo no Power Automate.")
		return jsonify(resultado), 500
	try:
		conteudo = _snapshot_para_xlsx(snapshot)
		workbook = load_workbook(io.BytesIO(conteudo), read_only=True, data_only=True)
		abas = list(workbook.sheetnames)
		contagens = {tabela: max(workbook[tabela].max_row - 1, 0) if tabela in abas else None for tabela in TABELAS_EXCEL}
		workbook.close()
		faltantes = [aba for aba in TABELAS_EXCEL if aba not in abas]
		if faltantes:
			raise RuntimeError("Faltam as abas obrigatórias: " + ", ".join(faltantes))
		etapa(3, "Estrutura do Excel", "OK", "Snapshot válido e abas obrigatórias encontradas.", abas=abas, contagens_remotas=contagens)
	except Exception as erro:
		etapa(3, "Estrutura do Excel", "ERRO", str(erro))
		resultado.update(status="FALHA_ESTRUTURA_EXCEL", mensagem_final="Verifique a estrutura das abas usuarios, demandas e historico.")
		return jsonify(resultado), 500
	try:
		conn = conectar()
		contagens_locais = {tabela: conn.execute(f"SELECT COUNT(*) FROM {tabela}").fetchone()[0] for tabela in TABELAS_EXCEL}
		conn.close()
		etapa(4, "Dados carregados", "OK", "Dados da aplicação disponíveis em memória.", contagens_locais=contagens_locais, contagens_remotas=contagens)
	except Exception as erro:
		etapa(4, "Dados carregados", "ERRO", str(erro))
		resultado.update(status="FALHA_DADOS", mensagem_final="Não foi possível validar os dados carregados pela aplicação.")
		return jsonify(resultado), 500
	resultado.update(status="OK", duracao_segundos=round((datetime.now() - inicio).total_seconds(), 2), conclusao="Diagnóstico concluído em modo somente leitura. Nenhuma alteração foi executada.")
	return jsonify(resultado), 200


@app.route("/logout")
def logout():
	session.clear(); return redirect(url_for("login"))


criar_banco()

if __name__ == "__main__":
	app.run(host="0.0.0.0", port=5000, debug=False)

