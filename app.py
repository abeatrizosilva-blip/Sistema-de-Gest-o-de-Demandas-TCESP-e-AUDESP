from datetime import date, datetime
import io
import html
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
from flask import Flask, redirect, render_template_string, request, send_file, session, url_for


app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "TROQUE-ESTA-CHAVE-POR-UMA-CHAVE-SECRETA")
CAMINHO_ONEDRIVE_WINDOWS = r"C:\Users\ana.silva\OneDrive - PRODESP\SP_AGUAS\Sistema de Gestão de Demandas - SP Aguas.xlsx"
PLANILHA = os.environ.get("EXCEL_DATABASE", os.environ.get("DATABASE", CAMINHO_ONEDRIVE_WINDOWS if os.name == "nt" and os.path.exists(CAMINHO_ONEDRIVE_WINDOWS) else "/tmp/sp_aguas.xlsx" if os.environ.get("VERCEL") else "sp_aguas.xlsx"))
SQLITE_LEGADO = os.environ.get("SQLITE_DATABASE", "sp_aguas.db")
ARQUIVO_LOCK = RLock()
CONEXAO_COMPARTILHADA = None
ONEDRIVE_ENABLED = os.environ.get("ONEDRIVE_ENABLED", "").lower() in {"1", "true", "sim", "yes"}
ONEDRIVE_TENANT_ID = os.environ.get("ONEDRIVE_TENANT_ID", "")
ONEDRIVE_CLIENT_ID = os.environ.get("ONEDRIVE_CLIENT_ID", "")
ONEDRIVE_CLIENT_SECRET = os.environ.get("ONEDRIVE_CLIENT_SECRET", "")
ONEDRIVE_USER = os.environ.get("ONEDRIVE_USER", "")
ONEDRIVE_PATH = os.environ.get("ONEDRIVE_PATH", "SP_AGUAS/sp_aguas.xlsx")

TABELAS_EXCEL = {
	"usuarios": ("id", "nome", "usuario", "email", "senha_hash", "perfil", "ativo", "aprovado", "criado_em"),
	"demandas": ("id", "numero_processo", "origem", "assunto", "area", "responsavel", "data_recebimento", "prazo_area", "prazo_fatal", "situacao", "prioridade", "observacoes", "criado_por", "criado_em", "atualizado_em"),
	"historico": ("id", "demanda_id", "usuario_id", "acao", "descricao", "data_hora"),
}


def _onedrive_configurado():
	return ONEDRIVE_ENABLED


def _onedrive_token():
	credenciais = (ONEDRIVE_TENANT_ID, ONEDRIVE_CLIENT_ID, ONEDRIVE_CLIENT_SECRET, ONEDRIVE_USER)
	if not all(credenciais):
		raise RuntimeError("OneDrive ativado, mas faltam ONEDRIVE_TENANT_ID, ONEDRIVE_CLIENT_ID, ONEDRIVE_CLIENT_SECRET ou ONEDRIVE_USER.")
	dados = urlencode({
		"client_id": ONEDRIVE_CLIENT_ID,
		"client_secret": ONEDRIVE_CLIENT_SECRET,
		"scope": "https://graph.microsoft.com/.default",
		"grant_type": "client_credentials",
	}).encode()
	url = f"https://login.microsoftonline.com/{quote(ONEDRIVE_TENANT_ID, safe='')}/oauth2/v2.0/token"
	try:
		with urlopen(Request(url, data=dados, headers={"Content-Type": "application/x-www-form-urlencoded"}), timeout=30) as resposta:
			return json.loads(resposta.read().decode())["access_token"]
	except (HTTPError, URLError, KeyError, json.JSONDecodeError) as erro:
		raise RuntimeError(f"Nao foi possivel autenticar no OneDrive: {erro}") from erro


def _onedrive_url():
	usuario = quote(ONEDRIVE_USER, safe="")
	caminho = quote(ONEDRIVE_PATH.strip("/"), safe="/")
	return f"https://graph.microsoft.com/v1.0/users/{usuario}/drive/root:/{caminho}:/content"


def _baixar_planilha_one_drive():
	if not _onedrive_configurado():
		return False
	try:
		with urlopen(Request(_onedrive_url(), headers={"Authorization": f"Bearer {_onedrive_token()}"}), timeout=60) as resposta:
			conteudo = resposta.read()
	except HTTPError as erro:
		if erro.code == 404:
			return False
		raise RuntimeError(f"Nao foi possivel baixar a planilha do OneDrive (HTTP {erro.code}).") from erro
	except URLError as erro:
		raise RuntimeError(f"Nao foi possivel acessar o OneDrive: {erro}") from erro
	os.makedirs(os.path.dirname(os.path.abspath(PLANILHA)), exist_ok=True)
	diretorio = os.path.dirname(os.path.abspath(PLANILHA))
	with tempfile.NamedTemporaryFile(suffix=".xlsx", dir=diretorio, delete=False) as temporario:
		caminho_temporario = temporario.name
	try:
		with open(caminho_temporario, "wb") as arquivo:
			arquivo.write(conteudo)
		os.replace(caminho_temporario, PLANILHA)
	finally:
		if os.path.exists(caminho_temporario):
			os.unlink(caminho_temporario)
	return True


def _enviar_planilha_one_drive():
	if not _onedrive_configurado():
		return
	try:
		with open(PLANILHA, "rb") as arquivo:
			conteudo = arquivo.read()
		request = Request(_onedrive_url(), data=conteudo, method="PUT", headers={
			"Authorization": f"Bearer {_onedrive_token()}",
			"Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
		})
		with urlopen(request, timeout=60):
			return
	except (HTTPError, URLError, OSError) as erro:
		raise RuntimeError(f"Nao foi possivel salvar a planilha no OneDrive: {erro}") from erro


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
		with ARQUIVO_LOCK:
			self._conn.commit()
			_salvar_planilha(self._conn)

	def close(self):
		pass


def conectar():
	global CONEXAO_COMPARTILHADA
	with ARQUIVO_LOCK:
		if CONEXAO_COMPARTILHADA is None:
			CONEXAO_COMPARTILHADA = ConexaoExcel()
		return CONEXAO_COMPARTILHADA


def agora():
	return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


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
	conn = conectar()
	conn.execute("INSERT INTO historico (demanda_id, usuario_id, acao, descricao, data_hora) VALUES (?, ?, ?, ?, ?)",
				 (demanda_id, usuario_id, acao, descricao, agora()))
	conn.commit()
	conn.close()


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
	conn = conectar()
	total = conn.execute("SELECT COUNT(*) total FROM usuarios").fetchone()["total"]
	conn.close()
	if total:
		return redirect(url_for("login"))
	if request.method == "POST":
		nome, usuario, senha = request.form["nome"].strip(), request.form["usuario"].strip(), request.form["senha"]
		if len(senha) < 8:
			return pagina("Configuração", '<div class="card"><div class="erro">A senha precisa ter pelo menos 8 caracteres.</div></div>')
		conn = conectar()
		conn.execute("INSERT INTO usuarios (nome, usuario, senha_hash, perfil, ativo, aprovado, criado_em) VALUES (?, ?, ?, ?, 1, 1, ?)",
					 (nome, usuario, bcrypt.hashpw(senha.encode(), bcrypt.gensalt()).decode(), "Administrador", agora()))
		conn.commit(); conn.close()
		return redirect(url_for("login"))
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
			try:
				conn = conectar(); conn.execute("INSERT INTO usuarios (nome, usuario, email, senha_hash, criado_em) VALUES (?, ?, ?, ?, ?)", (nome, usuario, email, bcrypt.hashpw(senha.encode(), bcrypt.gensalt()).decode(), agora())); conn.commit(); conn.close()
				return pagina("Cadastro realizado", '<div class="card"><h1>Cadastro realizado</h1><p>Aguarde a aprovação do administrador.</p><a class="btn" href="/">Voltar ao login</a></div>')
			except sqlite3.IntegrityError: erro = "Usuário ou e-mail já cadastrado."
		return pagina("Cadastro", f'<div class="card"><div class="erro">{erro}</div><a href="/cadastro" class="btn">Voltar</a></div>')
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
		conn = conectar(); cur = conn.execute("INSERT INTO demandas (numero_processo, origem, assunto, area, responsavel, data_recebimento, prazo_area, prazo_fatal, situacao, prioridade, observacoes, criado_por, criado_em, atualizado_em) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", (*valores, session["usuario_id"], agora(), agora())); demanda_id = cur.lastrowid; conn.commit(); conn.close(); registrar_historico(demanda_id, session["usuario_id"], "CRIACAO", "Demanda cadastrada."); return redirect(url_for("demandas"))
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
		conn.execute("UPDATE demandas SET numero_processo=?, origem=?, assunto=?, area=?, responsavel=?, data_recebimento=?, prazo_area=?, prazo_fatal=?, situacao=?, prioridade=?, observacoes=?, atualizado_em=? WHERE id=?", (*valores, agora(), id)); conn.commit(); conn.close(); registrar_historico(id, session["usuario_id"], "EDICAO", "Demanda alterada."); return redirect(url_for("demandas"))
	conn.close(); return pagina("Editar demanda", formulario(demanda))


@app.route("/excluir/<int:id>", methods=["POST"])
def excluir(id):
	if not administrador(): return "Acesso negado.", 403
	conn = conectar(); conn.execute("DELETE FROM demandas WHERE id = ?", (id,)); conn.execute("DELETE FROM historico WHERE demanda_id = ?", (id,)); conn.commit(); conn.close(); return redirect(url_for("demandas"))


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
		acao, usuario_id = request.form["acao"], request.form["usuario_id"]; conn = conectar()
		if acao == "aprovar": conn.execute("UPDATE usuarios SET aprovado=1, ativo=1 WHERE id=?", (usuario_id,))
		elif acao == "bloquear": conn.execute("UPDATE usuarios SET ativo=0 WHERE id=?", (usuario_id,))
		elif acao == "ativar": conn.execute("UPDATE usuarios SET ativo=1, aprovado=1 WHERE id=?", (usuario_id,))
		conn.commit(); conn.close()
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
	conn = conectar(); conn.execute("SELECT 1").fetchone(); conn.close(); ip = socket.gethostbyname(socket.gethostname())
	return pagina("Status", f'<div class="card"><h1>Status do sistema</h1><div class="alerta normal">Aplicação: <strong>OPERACIONAL</strong></div><p>Acesso local: http://{ip}:5000</p></div>')


@app.route("/logout")
def logout():
	session.clear(); return redirect(url_for("login"))


criar_banco()

if __name__ == "__main__":
	app.run(host="0.0.0.0", port=5000, debug=False)

