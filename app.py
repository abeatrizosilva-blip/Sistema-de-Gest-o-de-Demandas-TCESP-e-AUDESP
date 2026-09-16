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
from unicodedata import normalize as unicode_normalize
from urllib.parse import urlparse

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
PLANILHA = CAMINHO_ONEDRIVE_WINDOWS
# IMPORTANTE: a planilha acima é o único armazenamento permanente do sistema.
# Não usar DATABASE/EXCEL_DATABASE nem fallback para outro arquivo.

ARQUIVO_LOCK = RLock()
CONEXAO_COMPARTILHADA = None
def _env(*nomes, default=""):
	"""Lê a primeira variável disponível e remove espaços acidentais."""
	for nome in nomes:
		valor = os.environ.get(nome)
		if valor is not None and str(valor).strip() != "":
			return str(valor).strip()
	return default


# Persistência local: o sistema utiliza o arquivo Excel configurado em EXCEL_DATABASE.
# No Vercel, defina EXCEL_DATABASE para um armazenamento persistente se desejar usar a aplicação
# em produção; o sistema não depende de integração externa nem de Microsoft Graph.
TABELAS_EXCEL = {
	"usuarios": ("id", "nome", "usuario", "email", "senha_hash", "perfil", "ativo", "aprovado", "criado_em"),
	"demandas": ("id", "numero_processo", "numero_etc", "origem", "assunto", "area", "responsavel", "data_recebimento", "prazo_area", "prazo_fatal", "situacao", "prioridade", "observacoes", "doe_data", "doe_edicao", "doe_secao", "doe_palavra_chave", "doe_publicacao", "doe_url", "criado_por", "criado_em", "atualizado_em"),
	"historico": ("id", "demanda_id", "usuario_id", "acao", "descricao", "data_hora"),
}


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
				valores = tuple(linha[indices[coluna]] if coluna in indices and indices[coluna] < len(linha) else "" for coluna in colunas)
				conn.execute(f"INSERT INTO {tabela} ({', '.join(colunas)}) VALUES ({', '.join('?' for _ in colunas)})", valores)
	finally:
		workbook.close()
	return True


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
	workbook.close()


class ConexaoExcel:
	def __init__(self):
		# SQLite é usado SOMENTE como índice temporário em memória para manter
		# a aplicação compatível com as consultas existentes. Nenhum arquivo .db
		# é criado e nenhum dado é persistido no SQLite. A única persistência é o XLSX.
		self._conn = sqlite3.connect(":memory:", check_same_thread=False)
		self._conn.row_factory = sqlite3.Row
		_criar_esquema(self._conn)
		if os.path.exists(PLANILHA):
			_carregar_planilha(self._conn)
			self._conn.commit()

	def execute(self, consulta, parametros=()):
		return self._conn.execute(consulta, parametros)

	def executescript(self, consulta):
		return self._conn.executescript(consulta)

	def commit(self):
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
	conn = conectar()
	with ARQUIVO_LOCK:
		_salvar_planilha(conn)
	return {"status": "OK", "arquivo": PLANILHA}


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


def executar_mutacao_atomica(mutator, confirmador=None, tentativas=1):
	"""Aplica a alteração no Excel local e recarrega o banco SQLite."""
	global CONEXAO_COMPARTILHADA
	with ARQUIVO_LOCK:
		if not os.path.isfile(PLANILHA):
			raise FileNotFoundError(
				f"Planilha configurada não encontrada: {PLANILHA}. "
			"Verifique se o OneDrive está sincronizado e se o arquivo existe nesse caminho."
			)
		workbook = load_workbook(PLANILHA)
		try:
			resultado = mutator(workbook)
			workbook.save(PLANILHA)
		finally:
			workbook.close()
		# Recria apenas o índice temporário em memória a partir da planilha.
		if CONEXAO_COMPARTILHADA is not None:
			CONEXAO_COMPARTILHADA._conn.close()
		CONEXAO_COMPARTILHADA = ConexaoExcel()
		return {"status": "OK", **(resultado if isinstance(resultado, dict) else {"resultado": resultado})}


def salvar_demanda_atomicamente(valores, usuario_id, tentativas=3):
	"""
	Salva uma nova demanda diretamente na planilha Excel local.

	O nome da função é mantido por compatibilidade com a rota de cadastro,
	mas não há API externa, Power Automate, Graph ou banco em arquivo.
	A demanda e o primeiro registro do histórico são gravados no mesmo
	workbook e, em seguida, o índice SQLite em memória é recarregado.
	"""
	if len(valores) != 18:
		raise ValueError(f"Quantidade inesperada de campos da demanda: {len(valores)}")

	def mutator(workbook):
		# Garante que as abas existam mesmo se a planilha original for antiga.
		for tabela, colunas in TABELAS_EXCEL.items():
			if tabela not in workbook.sheetnames:
				ws = workbook.create_sheet(tabela)
				ws.append(list(colunas))

		planilha = workbook["demandas"]
		cabecalho = _cabecalho_planilha(planilha)
		demanda_id = _proximo_id(planilha, cabecalho)
		agora_valor = agora()
		campos_form = (
			"numero_processo", "numero_etc", "origem", "assunto", "area",
			"responsavel", "data_recebimento", "prazo_area", "prazo_fatal",
			"situacao", "prioridade", "observacoes", "doe_data", "doe_edicao",
			"doe_secao", "doe_palavra_chave", "doe_publicacao", "doe_url"
		)
		registro = dict(zip(campos_form, valores))
		registro.update({
			"id": demanda_id,
			"criado_por": usuario_id,
			"criado_em": agora_valor,
			"atualizado_em": agora_valor,
		})
		_append_registro(planilha, TABELAS_EXCEL["demandas"], registro)

		historico = workbook["historico"]
		historico_id = _proximo_id(historico, _cabecalho_planilha(historico))
		registro_historico = {
			"id": historico_id,
			"demanda_id": demanda_id,
			"usuario_id": usuario_id,
			"acao": "CRIACAO",
			"descricao": "Demanda cadastrada.",
			"data_hora": agora_valor,
		}
		_append_registro(historico, TABELAS_EXCEL["historico"], registro_historico)
		return {"demanda_id": demanda_id, "historico_id": historico_id}

	ultimo_erro = None
	for _ in range(max(1, tentativas)):
		try:
			return executar_mutacao_atomica(mutator)
		except PermissionError as exc:
				ultimo_erro = exc
				# O arquivo pode estar temporariamente bloqueado pelo Excel/OneDrive.
				time.sleep(0.8)
		except OSError as exc:
				ultimo_erro = exc
				time.sleep(0.8)
	if ultimo_erro:
		raise RuntimeError(f"Não foi possível salvar a demanda na planilha: {ultimo_erro}") from ultimo_erro
	raise RuntimeError("Não foi possível salvar a demanda na planilha.")


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
	"""Registra histórico no banco local."""
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
	try:
		conn = conectar()
		total = conn.execute("SELECT COUNT(*) FROM usuarios").fetchone()[0]
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

def _normalizar_pesquisa_doe(texto):
    texto = unicode_normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode("ascii")
    return " ".join(texto.upper().split())

def _url_pdf_doe(data_iso):
    dt = datetime.strptime(data_iso, "%Y-%m-%d").date()
    return f"https://doe.tce.sp.gov.br/v/pdf/{dt.year:04d}/{dt.month:02d}/doe-tce-{dt:%Y-%m-%d}.pdf"

def _download_pdf_url(url):
    """Baixa um PDF informado pelo usuário, sem depender da data do DOE."""
    url = (url or "").strip()
    partes_url = urlparse(url)
    if partes_url.scheme not in {"https", "http"} or not partes_url.netloc:
        raise ValueError("Informe um link HTTP/HTTPS válido para o PDF.")
    host = (partes_url.hostname or "").lower().rstrip(".")
    if host not in {"doe.tce.sp.gov.br", "tce.sp.gov.br", "www.tce.sp.gov.br"} and not host.endswith(".tce.sp.gov.br"):
        raise ValueError("Por segurança, o link deve pertencer ao domínio oficial do TCESP.")
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; SP-AGUAS/1.0)",
        "Accept": "application/pdf,application/octet-stream,*/*;q=0.8",
        "Accept-Encoding": "identity",
        "Connection": "close",
    }
    ultimo = None
    for tentativa in range(1, 4):
        try:
            with requests.get(url, headers=headers, timeout=(10, 90), stream=True, allow_redirects=True) as r:
                r.raise_for_status()
                partes = []
                for bloco in r.iter_content(chunk_size=64 * 1024):
                    if bloco:
                        partes.append(bloco)
                conteudo = b"".join(partes)
            if not conteudo.startswith(b"%PDF"):
                raise RuntimeError("O link não retornou um arquivo PDF. Confira se o link é direto para o PDF.")
            if b"%%EOF" not in conteudo[-8192:]:
                raise RuntimeError(f"O PDF foi recebido incompleto ({len(conteudo)} bytes).")
            return conteudo
        except Exception as exc:
            ultimo = exc
            if tentativa < 3:
                time.sleep(0.7 * tentativa)
    raise RuntimeError(f"Não foi possível baixar o PDF informado: {ultimo}")

def _extrair_metadados_pdf(reader, texto_inicial=""):
    cabecalho = " ".join((texto_inicial or "").split())
    data_publicacao = ""
    data_disponibilizacao = ""
    edicao = "DOE-TCESP"
    for padrao in (
        r"(?:Data de publicação|Data da publicação|Publicação)\s*[:—-]?\s*(\d{2}/\d{2}/\d{4})",
        r"(?:publica[çc][ãa]o)\s*(?:em|:)?\s*(\d{2}/\d{2}/\d{4})",
    ):
        m = re.search(padrao, cabecalho, re.I)
        if m:
            data_publicacao = m.group(1)
            break
    m = re.search(r"(?:disponibiliza[çc][ãa]o|disponibilizado)\s*(?:em|:)?\s*(\d{2}/\d{2}/\d{4})", cabecalho, re.I)
    if m:
        data_disponibilizacao = m.group(1)
    m = re.search(r"(?:EDI[ÇC][ÃA]O|Edi[çc][ãa]o)\s*(?:n[ºo°]?\s*)?(\d+)", cabecalho, re.I)
    if m:
        edicao = f"DOE-TCESP nº {m.group(1)}"
    return {"data_publicacao": data_publicacao, "data_disponibilizacao": data_disponibilizacao, "edicao": edicao}

def _processos_no_texto(texto):
    encontrados = []
    padroes = (
        r"\b(?:e?TC|TCESP)[-\s]?\d{1,8}[/.-]\d{1,4}[/.-]\d{2,4}\b",
        r"\b\d{5,8}/\d{2,4}\b",
    )
    for padrao in padroes:
        encontrados.extend(re.findall(padrao, texto or "", flags=re.I))
    vistos = []
    for item in encontrados:
        item = re.sub(r"\s+", "", item)
        if item not in vistos:
            vistos.append(item)
    return vistos[:20]

def _extrair_ocorrencias_doe(conteudo, data_iso="", url="", palavra_chave=""):
    if PdfReader is None:
        raise RuntimeError("Dependência pypdf não instalada. Execute pip install -r requirements.txt.")
    try:
        reader = PdfReader(io.BytesIO(conteudo))
    except Exception as exc:
        raise RuntimeError(f"Não foi possível ler o PDF: {exc}") from exc
    textos = []
    for pagina_num, page in enumerate(reader.pages, start=1):
        try:
            texto = page.extract_text() or ""
        except Exception:
            texto = ""
        textos.append(texto)
    texto_inicial = " ".join(textos[:2])
    meta = _extrair_metadados_pdf(reader, texto_inicial)
    termos = [palavra_chave] if palavra_chave else DOE_KEYWORDS
    resultados = []
    for pagina_num, texto in enumerate(textos, start=1):
        normalizado = _normalizar_pesquisa_doe(texto)
        if not normalizado:
            continue
        for termo in termos:
            alvo = _normalizar_pesquisa_doe(termo)
            if not alvo:
                continue
            pos = normalizado.find(alvo)
            if pos < 0:
                continue
            inicio = max(0, pos - 420)
            fim = min(len(texto), pos + len(termo) + 650)
            trecho = " ".join(texto[inicio:fim].split())
            resultados.append({
                "data": data_iso or meta["data_publicacao"],
                "data_publicacao": meta["data_publicacao"],
                "data_disponibilizacao": meta["data_disponibilizacao"],
                "edicao": meta["edicao"],
                "secao": f"Página {pagina_num}",
                "palavra_chave": termo,
                "trecho": trecho,
                "pagina": pagina_num,
                "url": url,
                "processos": _processos_no_texto(trecho),
            })
    return resultados[:200], meta, len(textos), any(bool(t.strip()) for t in textos)

@app.route("/api/doe-tcesp/extrair", methods=["POST"])
def extrair_doe_tcesp():
    if (resposta := acesso_login()):
        return resposta
    palavra = (request.form.get("palavra_chave") or "").strip()
    if palavra and _normalizar_pesquisa_doe(palavra) not in {_normalizar_pesquisa_doe(k) for k in DOE_KEYWORDS}:
        return jsonify({"erro": "Palavra-chave não cadastrada.", "resultados": []}), 400
    try:
        link = (request.form.get("url") or "").strip()
        if not link:
            return jsonify({"erro": "Para leitura pelo servidor, informe o link direto do PDF. PDFs selecionados do computador são lidos diretamente no navegador.", "resultados": []}), 400
        conteudo = _download_pdf_url(link)
        resultados, meta, paginas, possui_texto = _extrair_ocorrencias_doe(conteudo, url=link, palavra_chave=palavra)
        if not possui_texto:
            return jsonify({"erro": "O PDF foi recebido, mas não possui texto pesquisável. Se for uma publicação digitalizada, será necessário OCR.", "resultados": [], "meta": meta}), 422
        return jsonify({"resultados": resultados, "meta": meta, "paginas": paginas, "url": link})
    except Exception as exc:
        return jsonify({"erro": str(exc), "resultados": []}), 502


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
			raise RuntimeError("O cadastro não foi confirmado no banco local.")
		except Exception as erro:
			return pagina("Nova demanda", f'<div class="card"><div class="erro">Demanda NÃO confirmada no banco local. {html.escape(str(erro))}</div><a class="btn" href="/nova-demanda">Tentar novamente</a> <a class="btn btn-cinza" href="/demandas">Voltar</a></div>')
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
<p class="muted">Em vez de o sistema procurar o DOE automaticamente, informe o link direto da publicação ou envie o PDF do dia. O sistema lerá o documento e localizará as palavras-chave e números de processo.</p>
<div class="form-grid">
<div class="full"><label>Link da publicação/PDF</label><input type="url" id="doe_link" placeholder="Cole aqui o link direto do PDF do DOE-TCESP"></div>
<div><label>Ou selecione o PDF</label><input type="file" id="doe_pdf" accept="application/pdf"></div>
<div><label>Palavra-chave</label><select id="doe_pesquisa_keyword"><option value="">Todas as palavras-chave</option></select></div>
<div class="full"><button type="button" class="btn btn-outline" onclick="extrairDOE()">📄 Ler publicação e extrair dados</button></div>
</div>
<div id="doe_status" class="muted" style="margin-top:10px"></div><div id="doe_resultados" style="margin-top:12px"></div>
<div class="form-grid" style="margin-top:12px">
<div><label>Data DOE</label><input name="doe_data" id="doe_data" value="{valor("doe_data")}" placeholder="AAAA-MM-DD"></div>
<div><label>Edição DOE</label><input name="doe_edicao" id="doe_edicao" value="{valor("doe_edicao")}" placeholder="Ex.: edição diária"></div>
<div><label>Seção DOE</label><input name="doe_secao" id="doe_secao" value="{valor("doe_secao")}"></div>
<div><label>Palavra-chave encontrada</label><input name="doe_palavra_chave" id="doe_palavra_chave" value="{valor("doe_palavra_chave")}" readonly></div>
<div class="full"><label>Publicação / trecho localizado</label><textarea name="doe_publicacao" id="doe_publicacao" placeholder="O trecho da publicação selecionada aparecerá aqui.">{valor("doe_publicacao")}</textarea></div>
<div class="full"><label>Link oficial da publicação</label><input name="doe_url" id="doe_url" value="{valor("doe_url")}" placeholder="Link do PDF ou da publicação"></div>
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

function normalizarDOE(v) {{
  return String(v || "").normalize("NFD").replace(/[\\u0300-\u036f]/g, "").toUpperCase();
}}
function processosNoTexto(texto) {{
  const encontrados = [];
  const padroes = [/(?:e?TC|TCESP)[-\\s]?\\d{{1,8}}[/.-]\\d{{1,4}}[/.-]\\d{{2,4}}/gi, /\\b\\d{{5,8}}[/.-]\\d{{2,4}}\b/g];
  padroes.forEach(re => {{ for (const m of String(texto || "").matchAll(re)) {{ const v=m[0].replace(/\\s+/g, ""); if(!encontrados.includes(v)) encontrados.push(v); }} }});
  return encontrados.slice(0,20);
}}
function extrairDataPublicacao(texto) {{
  const m = String(texto || "").match(/(?:Data de publica[cç][aã]o|Data da publica[cç][aã]o|Publica[cç][aã]o)\\s*[:—-]?\\s*(\\d{{2}}\\/\\d{{2}}\\/\\d{{4}})/i);
  return m ? m[1] : "";
}}
function formatarDataISO(data) {{
  return data && /^\\d{{2}}\\/\\d{{2}}\\/\\d{{4}}$/.test(data) ? data.split("/").reverse().join("-") : "";
}}
function carregarPdfJs() {{
  return new Promise((resolve, reject) => {{
    if(window.pdfjsLib) return resolve(window.pdfjsLib);
    const script=document.createElement("script");
    script.src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
    script.onload=() => {{ window.pdfjsLib.GlobalWorkerOptions.workerSrc="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js"; resolve(window.pdfjsLib); }};
    script.onerror=() => reject(new Error("Não foi possível carregar o leitor de PDF no navegador."));
    document.head.appendChild(script);
  }});
}}
function mostrarResultadosDOE(resultados, paginas, url, meta) {{
  const box=document.getElementById("doe_resultados"), status=document.getElementById("doe_status");
  if(meta) {{
    document.getElementById("doe_edicao").value=meta.edicao || document.getElementById("doe_edicao").value;
    if(meta.data_publicacao) document.getElementById("doe_data").value=formatarDataISO(meta.data_publicacao);
  }}
  if(!resultados.length) {{
    box.innerHTML='<div class="empty"><strong>Nenhuma palavra-chave encontrada.</strong><br>O documento foi lido, mas nenhum termo selecionado foi localizado.</div>';
    status.textContent="Documento lido sem ocorrências."; return;
  }}
  window._doeResultados=resultados;
  box.innerHTML=resultados.map((item,i)=>`<div class="alerta normal" style="margin-top:8px"><strong>${{escapeHtml(item.palavra_chave)}}</strong><div class="muted">${{escapeHtml(item.secao)}}${{item.data_publicacao?' · Publicação: '+escapeHtml(item.data_publicacao):''}}</div><div style="margin-top:6px">${{escapeHtml(item.trecho)}}</div>${{item.processos&&item.processos.length?'<div class="muted" style="margin-top:6px"><strong>Processo(s):</strong> '+escapeHtml(item.processos.join(", "))+'</div>':''}}<button type="button" class="btn" style="margin-top:9px" onclick="usarDOE(${{i}})">Usar esta ocorrência</button></div>`).join("");
  status.textContent=`${{resultados.length}} ocorrência(s) encontrada(s) em ${{paginas}} página(s).`;
  if(url) document.getElementById("doe_url").value=url;
}}
async function extrairPdfNoNavegador(arquivo, palavra) {{
  const pdfjsLib=await carregarPdfJs();
  const buffer=await arquivo.arrayBuffer();
  const pdf=await pdfjsLib.getDocument({{data:buffer}}).promise;
  const resultados=[];
  const termos=palavra ? [palavra] : DOE_KEYWORDS;
  let dataPublicacao="", edicao="DOE-TCESP";
  for(let pagina=1; pagina<=pdf.numPages; pagina++) {{
    const page=await pdf.getPage(pagina);
    const content=await page.getTextContent();
    const texto=content.items.map(x=>x.str || "").join(" ");
    if(!dataPublicacao) dataPublicacao=extrairDataPublicacao(texto);
    const normalizado=normalizarDOE(texto);
    for(const termo of termos) {{
      const alvo=normalizarDOE(termo), pos=normalizado.indexOf(alvo);
      if(pos<0) continue;
      const inicio=Math.max(0,pos-420), fim=Math.min(texto.length,pos+termo.length+650);
      resultados.push({{data:formatarDataISO(dataPublicacao),data_publicacao:dataPublicacao,edicao,secao:`Página ${{pagina}}`,palavra_chave:termo,trecho:texto.slice(inicio,fim).replace(/\\s+/g," ").trim(),pagina,url:"",processos:processosNoTexto(texto.slice(inicio,fim))}});
    }}
  }}
  return {{resultados:resultados.slice(0,200),paginas:pdf.numPages,meta:{{data_publicacao:dataPublicacao,edicao}},url:""}};
}}
async function extrairDOE() {{
  const link=document.getElementById("doe_link").value.trim();
  const arquivo=document.getElementById("doe_pdf").files[0];
  const palavra=document.getElementById("doe_pesquisa_keyword").value;
  const status=document.getElementById("doe_status");
  if(!link && !arquivo) {{ status.textContent="Informe o link do PDF ou selecione um PDF."; return; }}
  status.textContent="Lendo o documento e procurando as palavras-chave..."; document.getElementById("doe_resultados").innerHTML="";
  try {{
    if(arquivo) {{
      if(arquivo.type !== "application/pdf" && !arquivo.name.toLowerCase().endsWith(".pdf")) throw new Error("Selecione um arquivo PDF válido.");
      const j=await extrairPdfNoNavegador(arquivo,palavra);
      mostrarResultadosDOE(j.resultados,j.paginas,"",j.meta);
      return;
    }}
    const body=new URLSearchParams({{url:link,palavra_chave:palavra}});
    const r=await fetch("/api/doe-tcesp/extrair",{{method:"POST",headers:{{"Content-Type":"application/x-www-form-urlencoded;charset=UTF-8","Accept":"application/json"}},body}});
    const texto=await r.text();
    let j; try {{ j=JSON.parse(texto); }} catch(parseError) {{ throw new Error(texto ? `O servidor retornou uma resposta que não é JSON: ${{texto.slice(0,180)}}` : `O servidor não retornou dados (HTTP ${{r.status}}).`); }}
    if(!r.ok) throw new Error(j.erro || j.message || `Erro HTTP ${{r.status}}`);
    mostrarResultadosDOE(j.resultados || [],j.paginas || 0,j.url || link,j.meta);
  }} catch(e) {{ status.textContent="Erro: "+(e.message || "Não foi possível ler a publicação."); }}
}}
function usarDOE(i) {{
  const item=window._doeResultados[i]; if(!item) return;
  const data=item.data_publicacao ? formatarDataISO(item.data_publicacao) : (item.data||"");
  document.getElementById("doe_data").value=data;
  document.getElementById("doe_edicao").value=item.edicao||"";
  document.getElementById("doe_secao").value=item.secao||"";
  document.getElementById("doe_palavra_chave").value=item.palavra_chave||"";
  document.getElementById("doe_publicacao").value=item.trecho||"";
  if(item.url) document.getElementById("doe_url").value=item.url;
  const proc=item.processos&&item.processos.length?item.processos[0]:"";
  if(proc && !document.querySelector('input[name="numero"]').value) document.querySelector('input[name="numero"]').value=proc;
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
	try:
		conn = conectar()
		contagens = {tabela: conn.execute(f"SELECT COUNT(*) FROM {tabela}").fetchone()[0] for tabela in TABELAS_EXCEL}
		return jsonify({"status": "OK", "modo": "Excel local (única persistência)", "arquivo": PLANILHA, "arquivo_existente": os.path.exists(PLANILHA), "tabelas": contagens, "vercel": bool(os.environ.get("VERCEL"))})
	except Exception as erro:
		return jsonify({"status": "ERRO", "erro": str(erro), "arquivo": PLANILHA}), 500


@app.route("/teste-integracao")
def teste_integracao():
	if not usuario_logado():
		return jsonify({"status": "ERRO", "mensagem": "Usuário não autenticado."}), 401
	if not administrador():
		return jsonify({"status": "ERRO", "mensagem": "Somente administradores podem executar o teste."}), 403
	try:
		conn = conectar()
		contagens = {tabela: conn.execute(f"SELECT COUNT(*) FROM {tabela}").fetchone()[0] for tabela in TABELAS_EXCEL}
		return jsonify({"status": "SUCESSO", "modo": "Excel local (única persistência)", "arquivo": PLANILHA, "tabelas": contagens, "mensagem": "A aplicação está funcionando sem integração externa."})
	except Exception as erro:
		return jsonify({"status": "ERRO", "erro": str(erro)}), 500


@app.route("/health")
def health():
	return jsonify({"status": "ok", "service": "sp-aguas"}), 200


@app.route("/diagnostico-sync")
def diagnostico_sync():
	if (resposta := acesso_login()): return resposta
	if not administrador():
		return jsonify({"status": "FALHA", "erro": "Somente administradores podem executar o diagnóstico."}), 403
	try:
		conn = conectar()
		contagens = {tabela: conn.execute(f"SELECT COUNT(*) FROM {tabela}").fetchone()[0] for tabela in TABELAS_EXCEL}
		return jsonify({"status": "OK", "somente_leitura": True, "modo": "Excel local (única persistência)", "arquivo": PLANILHA, "arquivo_existente": os.path.exists(PLANILHA), "contagens": contagens, "conclusao": "Diagnóstico concluído sem integração externa."})
	except Exception as erro:
		return jsonify({"status": "FALHA", "erro": str(erro)}), 500


@app.route("/logout")
def logout():
	session.clear(); return redirect(url_for("login"))


criar_banco()

# Diagnóstico inicial: evita que o sistema rode silenciosamente gravando em outro arquivo.
if not os.path.isfile(PLANILHA):
	print("\n[ERRO] PLANILHA DO SISTEMA NÃO ENCONTRADA")
	print(f"[ERRO] Caminho configurado: {PLANILHA}")
	print("[ERRO] Verifique o OneDrive e confirme se o arquivo está disponível localmente.\n")
else:
	print(f"[OK] Planilha utilizada pelo sistema: {PLANILHA}")

if __name__ == "__main__":
	app.run(host="0.0.0.0", port=5000, debug=False)

