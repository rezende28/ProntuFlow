from flask import Flask, request, redirect, session
from datetime import datetime, timedelta, date
import sqlite3
import os

app = Flask(__name__)
app.secret_key = "docmed_secret_key_123"

print("PASTA ATUAL:", os.getcwd())

def init_db():
    conn = sqlite3.connect("docmed.db")
    c = conn.cursor()

    c.execute("""
    CREATE TABLE IF NOT EXISTS solicitacoes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,

        nome TEXT,
        prontuario TEXT,
        atendimento TEXT,

        telefone TEXT,
        email TEXT,

        tipo_atendimento TEXT,
        tipo_envio TEXT,

        tipo_copia TEXT,
        tipo_solicitacao TEXT,

        data_atendimento TEXT,
        data_alta TEXT,

        folhas INTEGER,
        valor REAL,

        data_solicitacao TEXT,
        data_entrega TEXT,
        data_pronto TEXT,
        data_entregue TEXT,

        status TEXT,
        observacoes TEXT
    )
    """)

    c.execute("""
CREATE TABLE IF NOT EXISTS usuarios (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome TEXT,
    usuario TEXT UNIQUE,
    senha TEXT,
    perfil TEXT
)
""")
    
    c.execute("""
CREATE TABLE IF NOT EXISTS historico (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    solicitacao_id INTEGER,

    data TEXT,
    usuario TEXT,
    acao TEXT,

    detalhes TEXT
)
""")
    
    c.execute("""
CREATE TABLE IF NOT EXISTS solicitacoes_excluidas (
    id INTEGER,
    nome TEXT,
    prontuario TEXT,
    atendimento TEXT,
    telefone TEXT,
    email TEXT,
    tipo_atendimento TEXT,
    tipo_envio TEXT,
    tipo_copia TEXT,
    tipo_solicitacao TEXT,
    data_atendimento TEXT,
    data_alta TEXT,
    folhas INTEGER,
    valor REAL,
    data_solicitacao TEXT,
    data_entrega TEXT,
    data_pronto TEXT,
    data_entregue TEXT,
    status TEXT,
    observacoes TEXT,

    excluido_por TEXT,
    data_exclusao TEXT,
    justificativa TEXT
)
""")
    
    c.execute("""
    INSERT OR IGNORE INTO usuarios
    (nome, usuario, senha, perfil)
    VALUES
    ('Administrador', 'admin', '1234', 'Administrador')
    """)
    
    try:
        c.execute("ALTER TABLE solicitacoes ADD COLUMN solicitado_por TEXT")
    except:
        pass    

    conn.commit()
    conn.close()

DIAS_PRAZO = 15

FERIADOS = [
    "01-01",
    "21-04",
    "01-05",
    "07-09",
    "12-10",
    "02-11",
    "15-11",
    "20-11",
    "25-12"
]

def eh_dia_util(data):
    if data.weekday() >= 5:
        return False

    if data.strftime("%d-%m") in FERIADOS:
        return False

    return True


def adicionar_dias_uteis(data_inicio, dias):
    data = data_inicio
    adicionados = 0

    while adicionados < dias:
        data += timedelta(days=1)

        if eh_dia_util(data):
            adicionados += 1

    return data

def cor_prazo(data_entrega, status):
    if status in ["❌ Cancelado", "📦 Entregue"]:
        return "secondary"

    if not data_entrega:
        return "secondary"

    try:
        entrega = datetime.strptime(str(data_entrega), "%Y-%m-%d").date()
    except:
        return "secondary"

    hoje = date.today()
    diff = (entrega - hoje).days

    if diff < 0:
        return "danger"   
    elif diff <= 3:
        return "warning"  
    else:
        return "success"  

def formatar_data(data):
    if not data:
        return ""

    try:
        return datetime.strptime(str(data), "%Y-%m-%d").strftime("%d/%m/%Y")
    except:
        return data    
    
def listar_solicitacoes():
    conn = sqlite3.connect("docmed.db")
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    c.execute("SELECT * FROM solicitacoes ORDER BY id DESC")
    dados = c.fetchall()

    conn.close()
    return dados

def cor_status(status):
    if status == "⏳ Pendente":
        return "warning"   
    elif status == "✅ Autorizado":
        return "info"   
    elif status == "📦 Entregue":
        return "success"   
    elif status == "❌ Cancelado":
        return "secondary"    
    else:
        return "light"
    
def salvar_historico(solicitacao_id, acao, usuario="Sistema"):
    conn = sqlite3.connect("docmed.db")
    c = conn.cursor()

    c.execute("""
        INSERT INTO historico (solicitacao_id, data, usuario, acao)
        VALUES (?, ?, ?, ?)
    """, (
        solicitacao_id,
        datetime.now().strftime("%d/%m/%Y %H:%M"),
        usuario,
        acao
    ))

    conn.commit()
    conn.close()

def verificar_login(usuario, senha):
    conn = sqlite3.connect("docmed.db")
    c = conn.cursor()

    c.execute("""
        SELECT * FROM usuarios
        WHERE usuario = ? AND senha = ?
    """, (usuario, senha))

    user = c.fetchone()
    conn.close()
    return user

import csv
from flask import Response
import io

@app.route("/exportar_excel")
def exportar_excel():
    if "usuario" not in session:
        return redirect("/login")

    conn = sqlite3.connect("docmed.db")
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("SELECT * FROM solicitacoes ORDER BY id DESC")
    solicitacoes = c.fetchall()
    conn.close()

    # Criar um fluxo de memória para gerar o arquivo CSV
    output = io.StringIO()
    writer = csv.writer(output, delimiter=';', quoting=csv.QUOTE_MINIMAL)

    # Escrever o cabeçalho das colunas no Excel (com codificação correta para acentos)
    writer.writerow([
        "ID", "Prontuário", "Atendimento", "Nome do Paciente", "Telefone", "E-mail",
        "Solicitado Por", "Tipo de Envio", "Tipo de Cópia", "Tipo de Solicitação",
        "Data Atendimento", "Data Alta", "Folhas", "Valor (R$)", 
        "Data Solicitação", "Prazo Entrega", "Data Pronto", "Data Entregue", "Status", "Observações"
    ])

    # Preencher com os dados do banco
    for s in solicitacoes:
        writer.writerow([
            s["id"],
            s["prontuario"] or "",
            s["atendimento"] or "",
            s["nome"] or "",
            s["telefone"] or "",
            s["email"] or "",
            s["solicitado_por"] or "",
            s["tipo_envio"] or "",
            s["tipo_copia"] or "",
            s["tipo_solicitacao"] or "",
            s["data_atendimento"] or "",
            s["data_alta"] or "",
            s["folhas"] or 0,
            f"{float(s['valor'] or 0):.2f}".replace('.', ','), # Formato de moeda para o Excel BR
            s["data_solicitacao"] or "",
            s["data_entrega"] or "",
            s["data_pronto"] or "",
            s["data_entregue"] or "",
            s["status"] or "",
            s["observacoes"] or ""
        ])

    # Configurar a resposta para o navegador entender que é um download de arquivo
    excel_data = output.getvalue()
    # Adiciona o BOM do UTF-8 para o Excel reconhecer os acentos (ç, á, é, ⏳) corretamente
    bom = b'\xef\xbb\xbf'
    response = Response(bom + excel_data.encode("utf-8"), mimetype="text/csv")
    response.headers["Content-Disposition"] = f"attachment; filename=relatorio_solicitacoes_{date.today()}.csv"
    return response

@app.route("/indicadores")
def indicadores():
    if "usuario" not in session:
        return redirect("/login")

    conn = sqlite3.connect("docmed.db")
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("SELECT * FROM solicitacoes ORDER BY id DESC")
    solicitacoes = c.fetchall()
    conn.close()

    # Lógica dos indicadores (incluindo agora os cancelados)
    estatisticas_mes = {}
    meses_ano = [f"2026-{str(i).zfill(2)}" for i in range(1, 13)]
    
    for m in meses_ano:
        estatisticas_mes[m] = {
            "total": 0, 
            "fisico": 0, 
            "digital": 0, 
            "no_prazo": 0, 
            "atrasado": 0, 
            "cancelado": 0, # Nova métrica adicionada
            "valor": 0.0
        }

    hoje = date.today()
    for s in solicitacoes:
        try:
            mes_chave = datetime.strptime(s["data_solicitacao"], "%Y-%m-%d").strftime("%Y-%m")
        except:
            mes_chave = None

        if mes_chave in estatisticas_mes:
            estatisticas_mes[mes_chave]["total"] += 1
            
            # Se o status contiver "Cancelado", contamos especificamente aqui
            status_atual = s["status"] or ""
            if "cancelado" in status_atual.lower():
                estatisticas_mes[mes_chave]["cancelado"] += 1
            else:
                # Contagens de tipo de envio e valor apenas se não estiver cancelado
                if s["tipo_envio"] == "Físico":
                    estatisticas_mes[mes_chave]["fisico"] += 1
                    estatisticas_mes[mes_chave]["valor"] += float(s["valor"] or 0)
                else:
                    estatisticas_mes[mes_chave]["digital"] += 1

                # Lógica de prazos apenas para os não cancelados
                try:
                    if s["data_entrega"]:
                        prazo_limite = datetime.strptime(s["data_entrega"], "%Y-%m-%d").date()
                        if s["data_entregue"]:
                            data_real = datetime.strptime(s["data_entregue"], "%Y-%m-%d").date()
                            if data_real <= prazo_limite:
                                estatisticas_mes[mes_chave]["no_prazo"] += 1
                            else:
                                estatisticas_mes[mes_chave]["atrasado"] += 1
                        elif prazo_limite < hoje:
                            estatisticas_mes[mes_chave]["atrasado"] += 1
                        else:
                            estatisticas_mes[mes_chave]["no_prazo"] += 1
                except:
                    pass

    # Montando as colunas horizontais da tabela
    linha_meses = ""
    linha_total = ""
    linha_fisico = ""
    linha_digital = ""
    linha_prazo = ""
    linha_atraso = ""
    linha_cancelado = "" # Nova linha no HTML
    linha_porcentagem = ""
    linha_valores = ""

    meses_ativos = [m for m in meses_ano if estatisticas_mes[m]["total"] > 0]
    if not meses_ativos:
        meses_ativos = [f"2026-{str(i).zfill(2)}" for i in range(1, 7)]

    for m in meses_ativos:
        nome_m = mes_da_solicitacao(f"{m}-01").split(" ")[0]
        dados = estatisticas_mes[m]
        
        # O percentual de atraso ignora os cancelados para não distorcer a produtividade real do prazo
        total_validos = dados["no_prazo"] + dados["atrasado"]
        pct_atraso = f"{(dados['atrasado'] / total_validos * 100):.0f}%" if total_validos > 0 else "0%"

        linha_meses += f"<th class='text-center bg-primary text-white'>{nome_m}</th>"
        linha_total += f"<td class='text-center fw-bold'>{dados['total']}</td>"
        linha_fisico += f"<td class='text-center text-primary'>📂 {dados['fisico']}</td>"
        linha_digital += f"<td class='text-center text-info'>💻 {dados['digital']}</td>"
        linha_prazo += f"<td class='text-center text-success'>{dados['no_prazo']}</td>"
        linha_atraso += f"<td class='text-center text-danger'>{dados['atrasado']}</td>"
        linha_cancelado += f"<td class='text-center text-muted'>❌ {dados['cancelado']}</td>" # Dados da nova linha
        linha_porcentagem += f"<td class='text-center fw-bold text-danger'>{pct_atraso}</td>"
        linha_valores += f"<td class='text-center fw-bold text-success'>R$ {dados['valor']:.2f}</td>"

    # Retornando o HTML exclusivo dessa tela
    return f"""
<!DOCTYPE html>
<html>
<head>
<title>DocMed - Indicadores</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
<style>
body {{ background: #f0f5fa; color: #000; font-family: 'Segoe UI', Arial, sans-serif; }}
.card {{ background: #fff; border-radius: 12px; border: 1px solid #cbd5e1; }}
</style>
</head>
<body>
<div class="container mt-4">
    <div class="d-flex justify-content-between align-items-center mb-3">
        <div>
            <h1 style="color: #1e3a8a;">📊 Painel Gerencial e Indicadores</h1>
            <h6 class="text-muted">Estatísticas Consolidadas do Setor</h6>
        </div>
        <div>
            <a href="/" class="btn btn-secondary">⬅️ Voltar para o Início</a>
        </div>
    </div>
    <hr>
    
    <div class="card p-4 shadow-sm mb-4">
        <div class="table-responsive">
            <table class="table table-bordered align-middle m-0">
                <thead>
                    <tr>
                        <th style="min-width: 180px;" class="bg-dark text-white">Mês</th>
                        {linha_meses}
                    </tr>
                </thead>
                <tbody>
                    <tr><td class="fw-bold bg-light">Nº de Solicitações Total</td>{linha_total}</tr>
                    <tr><td class="bg-light">Envios Físicos</td>{linha_fisico}</tr>
                    <tr><td class="bg-light">Envios Digitais</td>{linha_digital}</tr>
                    <tr><td class="bg-light text-success fw-bold">Entregues no Prazo</td>{linha_prazo}</tr>
                    <tr><td class="bg-light text-danger fw-bold">Atrasadas</td>{linha_atraso}</tr>
                    <tr><td class="bg-light text-secondary fw-bold">Canceladas</td>{linha_cancelado}</tr>
                    <tr><td class="bg-light text-danger fw-bold">% de Atrasos</td>{linha_porcentagem}</tr>
                    <tr><td class="bg-light text-success fw-bold">Valor Total (Cópias)</td>{linha_valores}</tr>
                </tbody>
            </table>
        </div>
    </div>
</div>
</body>
</html>
"""

@app.route("/")
def inicio():
    if "usuario" not in session:
        return redirect("/login")

    conn = sqlite3.connect("docmed.db")
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("SELECT * FROM solicitacoes ORDER BY id DESC")
    solicitacoes = c.fetchall()
    conn.close()

    total = len(solicitacoes)
    fisicos = 0
    digitais = 0
    valor_total = 0
    pendentes = 0
    atrasados = 0
    vencendo = 0
    cancelados = 0

    por_mes = {}
    hoje = date.today()

    for s in solicitacoes:
        if s["status"] == "⏳ Pendente":
            pendentes += 1
        elif s["status"] == "❌ Cancelado":
            cancelados += 1

        try:
            if s["data_entrega"] and s["status"] not in ["❌ Cancelado", "📦 Entregue"]:
                entrega = datetime.strptime(s["data_entrega"], "%Y-%m-%d").date()
                diff = (entrega - hoje).days
                if diff < 0:
                    atrasados += 1
                elif diff <= 3:  
                    vencendo += 1
        except:
            pass

        if s["tipo_envio"] == "Físico":
            fisicos += 1
            valor_total += float(s["valor"] or 0)
        else:
            digitais += 1

        try:
            mes = datetime.strptime(s["data_solicitacao"], "%Y-%m-%d").strftime("%Y-%m")
            por_mes[mes] = por_mes.get(mes, 0) + 1
        except:
            pass

    html = f"""
<!DOCTYPE html>
<html>
<head>
<title>DocMed V2</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
body {{ background: #f0f5fa; color: #000000; font-family: 'Segoe UI', Arial, sans-serif; }}
h1, h2, h3, h4, h5, h6 {{ color: #0f172a; font-weight: 600; }}
.card {{ background: #ffffff; color: #000000; border-radius: 12px; border: 1px solid #cbd5e1; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }}
.table td, .table th {{ vertical-align: middle; font-size: 14px; color: #000000; }}
.table-dark {{ background-color: #1e3a8a !important; color: #ffffff !important; }}
.table-dark th {{
    background-color: #1e3a8a !important;
    color: white !important;
    font-weight: bold;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}}
.btn-primary {{ background-color: #1e3a8a; border-color: #1e3a8a; }}
.btn-primary:hover {{ background-color: #1d4ed8; border-color: #1d4ed8; }}
.btn-success {{ background-color: #0d9488; border-color: #0d9488; color: white; }}
.btn-success:hover {{ background-color: #0f766e; border-color: #0f766e; color: white; }}
.btn-sm {{ padding: 4px 8px; font-size: 13px; }}

.bg-warning {{ background-color: #fef08a !important; color: #854d0e !important; }}
.bg-info {{ background-color: #ccfbf1 !important; color: #0f766e !important; }}
.bg-success {{ background-color: #dcfce7 !important; color: #166534 !important; }}
.bg-secondary {{ background-color: #e2e8f0 !important; color: #475569 !important; }}
</style>
</head>
<body>
<div class="container mt-4">

<div class="d-flex justify-content-between align-items-center">
    <div>
        <h1 style="color: #1e3a8a;">🏥 DocMed</h1>
        <h6 class="text-muted">Sistema de Controle de Prontuários</h6>
    </div>
    <div class="text-end">
        <span>👤 <b>{session.get("usuario")}</b></span> <br>
        <a href="/logout" class="btn btn-danger btn-sm mt-2">Sair</a>
        <a href="/indicadores" class="btn btn-info btn-sm mt-2 text-white" style="background-color: #1e3a8a; border: none;">📊 Ver Indicadores</a>
        {"<br><a href='/novo_usuario' class='btn btn-success btn-sm mt-2'>👤 Criar Novo Usuário</a> <a href='/lixeira' class='btn btn-secondary btn-sm mt-2'>🗑️ Lixeira</a>" if session.get("perfil") == "Administrador" else ""}
    </div>
</div>
<hr style="border-color: #cbd5e1;">

<div class="row mb-3 text-center">
    <div class="col"><div class="card p-3" style="border-left: 5px solid #eab308; background-color: #fef9c3;"><b>Pendentes:</b> {pendentes}</div></div>
    <div class="col"><div class="card p-3" style="border-left: 5px solid #f97316; background-color: #ffedd5;"><b>Vencendo (Até 3 dias):</b> {vencendo}</div></div>
    <div class="col"><div class="card p-3" style="border-left: 5px solid #ef4444; background-color: #fee2e2;"><b>Atrasados:</b> {atrasados}</div></div>
    <div class="col"><div class="card p-3" style="border-left: 5px solid #1e3a8a; background-color: #e0f2fe;"><b>Total Geral:</b> {total}</div></div>
</div>
<hr style="border-color: #cbd5e1;">

<div class="row">
    <div class="col-md-4"><canvas id="graficoEnvio"></canvas></div>
    <div class="col-md-4"><canvas id="graficoMes"></canvas></div>
    <div class="col-md-4"><canvas id="graficoStatus"></canvas></div>
</div>
<hr style="border-color: #cbd5e1;">

<div class="text-center mb-4 d-flex justify-content-center gap-3">
    <a href="/nova_solicitacao" class="btn btn-primary btn-lg px-5">➕ Nova Solicitação</a>
    <a href="/indicadores" class="btn btn-info btn-lg px-5 text-white">📊 Dashboard de Indicadores</a>
    <a href="/exportar_excel" class="btn btn-success btn-lg px-5">📊 Exportar para o Excel</a>
</div>
<hr style="border-color: #cbd5e1;">
"""

    html += f"""
<script>
new Chart(document.getElementById('graficoStatus'), {{
    type: 'doughnut',
    data: {{
        labels: ["Pendentes", "Vencendo (3d)", "Atrasados", "Cancelados"],
        datasets: [{{
            data: [{pendentes}, {vencendo}, {atrasados}, {cancelados}],
            backgroundColor: ['#fde047', '#fed7aa', '#fca5a5', '#cbd5e1']
        }}]
    }},
    options: {{ plugins: {{ legend: {{ labels: {{ color: '#000000' }} }} }} }}
}});
new Chart(document.getElementById('graficoEnvio'), {{
    type: 'pie',
    data: {{
        labels: ["Físico", "Digital"],
        datasets: [{{
            data: [{fisicos}, {digitais}],
            backgroundColor: ['#3b82f6', '#14b8a6']
        }}]
    }},
    options: {{ plugins: {{ legend: {{ labels: {{ color: '#000000' }} }} }} }}
}});
new Chart(document.getElementById('graficoMes'), {{
    type: 'bar',
    data: {{
        labels: {[mes_da_solicitacao(f"{m}-01") for m in por_mes.keys()]},
        datasets: [{{
            label: 'Solicitações',
            data: {list(por_mes.values())},
            backgroundColor: '#60a5fa'
        }}]
    }},
    options: {{ scales: {{ y: {{ ticks: {{ color: '#000000' }} }}, x: {{ ticks: {{ color: '#000000' }} }} }}, plugins: {{ legend: {{ labels: {{ color: '#000000' }} }} }} }}
}});
</script>
"""

    from collections import defaultdict
    
    def safe_id(text):
        return text.replace(" ", "_").replace("-", "_")

    html += """
    <hr style="border-color: #cbd5e1;">
    <div class="row align-items-center mb-3">
        <div class="col-md-6">
            <h3 style="color: #1e3a8a; margin-bottom: 0;">📅 Solicitações por Mês</h3>
        </div>
        <div class="col-md-6">
            <div class="input-group">
                <span class="input-group-text bg-white border-end-0">🔍</span>
                <input type="text" id="buscadorInput" class="form-control border-start-0 shadow-sm" 
                       placeholder="Buscar por nome do paciente ou nº de atendimento..." onkeyup="filtrarTabela()">
            </div>
        </div>
    </div>

    <ul class="nav nav-tabs mb-3" role="tablist">
    """

    grupos = defaultdict(list)
    for s in solicitacoes:
        try:
            mes = datetime.strptime(s["data_solicitacao"], "%Y-%m-%d").strftime("%Y-%m")
        except:
            mes = "Sem data"
        grupos[mes].append(s)

    # CORRIGIDO: reverse=False para os botões das abas ficarem na ordem cronológica correta
    for i, (mes_grupo, itens) in enumerate(sorted(grupos.items(), reverse=False)):
        tab_id = safe_id(mes_grupo)
        ativo = "active" if i == 0 else ""
        nome_mes = mes_da_solicitacao(f"{mes_grupo}-01") if mes_grupo != "Sem data" else mes_grupo

        html += f"""
        <li class="nav-item">
            <button class="nav-link {ativo}"
                data-bs-toggle="tab"
                data-bs-target="#tab-{tab_id}"
                type="button">
            {nome_mes}
        </button>
    </li>
    """

    html += """
    </ul>

    <div class="tab-content">
    """
    
    # CORRIGIDO: reverse=False para as tabelas acompanharem a mesma ordem cronológica correta
    for i, (mes_grupo, itens) in enumerate(sorted(grupos.items(), reverse=False)):
        tab_id = safe_id(mes_grupo)
        ativo = "show active" if i == 0 else ""

        html += f"""
        <div class="tab-pane fade {ativo}" id="tab-{tab_id}" role="tabpanel">
            <div class="card p-4">
                <table class="table table-striped table-hover mt-3 tabela-solicitacoes">
                    <thead class="table-dark">
                        <tr>
                            <th>Prontuário</th>
                            <th>Atendimento</th>
                            <th>Nome</th>
                            <th>Telefone</th>
                            <th>Solicitado Por</th>
                            <th>Data da Alta</th>
                            <th>Tipo de Solicitação</th>
                            <th>Status</th>
                            <th>Valor</th>
                            <th>Prazo</th>
                            <th>Ações</th>
                        </tr>
                    </thead>
                    <tbody>
        """

        for s in itens:
            botao_excluir = f'<a href="/excluir/{s["id"]}" class="btn btn-sm btn-danger">🗑️</a>' if session.get("perfil") == "Administrador" else ""

            # Pegamos o status original guardado no banco (Pendente, Autorizado ou Cancelado)
            status_atual = s["status"] or "⏳ Pendente"
            status_limpo = status_atual.lower()
            
            texto_data = "Solicitado"
            data_crua = s["data_solicitacao"]

            # LÓGICA INTELIGENTE: Se não estiver cancelado, checamos as datas preenchidas
            if "cancelado" not in status_limpo:
                if s["data_entregue"] and s["data_entregue"].strip() != "":
                    status_atual = "📦 Entregue"
                    texto_data = "Retirado"
                    data_crua = s["data_entregue"]
                elif s["data_pronto"] and s["data_pronto"].strip() != "":
                    status_atual = "🚀 Pronto para Retirada"
                    texto_data = "Pronto"
                    data_crua = s["data_pronto"]
                elif "autorizado" in status_limpo:
                    texto_data = "Autorizado"
                    data_crua = s["data_solicitacao"]
            else:
                texto_data = "Cancelado"
                data_crua = s["data_solicitacao"]

            data_final_formatada = formatar_data(data_crua) if data_crua else "Sem data"
            data_status = f"<br><small class='text-muted' style='font-size:11px;'>{texto_data}: {data_final_formatada}</small>"

            html += f"""
            <tr class="linha-registro">
                <td>{s["prontuario"]}</td>
                <td class="coluna-atendimento">{s["atendimento"]}</td>
                <td class="coluna-nome">{s["nome"]}</td>
                <td>{s["telefone"] or ""}</td>
                <td>{s["solicitado_por"] or ""}</td>
                <td>{formatar_data(s["data_alta"])}</td>
                <td>{s["tipo_solicitacao"]}</td>
                <td>
                    <span class="badge bg-{cor_status(s["status"])}">{s["status"]}</span>
                    {data_status}
                </td>
                <td>R$ {float(s["valor"] or 0):.2f}</td>
                <td>{formatar_data(s["data_entrega"])}</td>
                <td>
                    <a href="/detalhes/{s["id"]}" class="btn btn-sm btn-info">🔍</a>
                    <a href="/editar/{s["id"]}" class="btn btn-sm btn-warning">✏️</a>
                    <a href="/historico/{s["id"]}" class="btn btn-sm btn-secondary">📜</a>
                    {botao_excluir}
                </td>
            </tr>
            """

        html += """
                    </tbody>
                </table>
            </div>
        </div>
        """

    # Mantendo o buscador inteligente que varre todas as abas ao mesmo tempo e foca no mês correto
    html += """
    </div>
<hr>

<script>
function filtrarTabela() {
    var input = document.getElementById("buscadorInput");
    var filtro = input.value.toLowerCase();
    var linhas = document.getElementsByClassName("linha-registro");
    var primeiraAbaEncontrada = null;

    for (var i = 0; i < linhas.length; i++) {
        var colunaNome = linhas[i].getElementsByClassName("coluna-nome")[0];
        var colunaAtendimento = linhas[i].getElementsByClassName("coluna-atendimento")[0];
        
        if (colunaNome && colunaAtendimento) {
            var txtNome = colunaNome.textContent || colunaNome.innerText;
            var txtAtendimento = colunaAtendimento.textContent || colunaAtendimento.innerText;
            
            if (txtNome.toLowerCase().indexOf(filtro) > -1 || txtAtendimento.toLowerCase().indexOf(filtro) > -1) {
                linhas[i].style.display = "";
                
                if (!primeiraAbaEncontrada && filtro !== "") {
                    var painelAba = linhas[i].closest('.tab-pane');
                    if (painelAba) {
                        primeiraAbaEncontrada = painelAba.id;
                    }
                }
            } else {
                linhas[i].style.display = "none";
            }
        }
    }

    if (primeiraAbaEncontrada) {
        var botaoAba = document.querySelector('[data-bs-target="#' + primeiraAbaEncontrada + '"]');
        if (botaoAba && !botaoAba.classList.contains('active')) {
            var tab = new bootstrap.Tab(botaoAba);
            tab.show();
        }
    }
}
</script>

<footer class="text-center text-muted mb-3">
    🏥 DocMed V2 • Desenvolvido por Stephany Rezende © 2026
</footer>

</div>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>

</body>
</html>
"""
    return html

def mes_da_solicitacao(data_str):
    meses = [
        "Janeiro", "Fevereiro", "Março", "Abril",
        "Maio", "Junho", "Julho", "Agosto",
        "Setembro", "Outubro", "Novembro", "Dezembro"
    ]
    data = datetime.strptime(data_str, "%Y-%m-%d")
    return f"{meses[data.month - 1]} {data.year}"

@app.route("/nova_solicitacao")
def nova_solicitacao():
    if "usuario" not in session:
        return redirect("/login")

    return """
    <html>
    <head>
        <title>Nova Solicitação</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
        <style>
            body { background: #f0f5fa; color: #000000; font-family: 'Segoe UI', Arial, sans-serif; }
            .card { background: #ffffff; color: #000000; border-radius: 12px; border: 1px solid #cbd5e1; }
            .btn-primary { background-color: #1e3a8a; border-color: #1e3a8a; }
            .btn-primary:hover { background-color: #1d4ed8; }
        </style>
    </head>
    <body>
    <div class="container mt-4">
        <div class="card shadow p-4">
            <h2 class="mb-4 text-center" style="color: #1e3a8a;">➕ Nova Solicitação</h2>
<form action="/salvar" method="post">
<div class="row">
    <div class="col-md-2 mb-3">
        <label class="form-label">Prontuário</label>
        <input type="text" name="prontuario" class="form-control" required>
    </div>
    <div class="col-md-2 mb-3">
        <label class="form-label">Atendimento</label>
        <input type="text" name="atendimento" class="form-control" required>
    </div>
    <div class="col-md-4 mb-3">
        <label class="form-label">Nome do Paciente</label>
        <input type="text" name="nome" class="form-control" required>
    </div>
    <div class="col-md-4 mb-3">
        <label class="form-label d-block">Tipo de Atendimento</label>
        <input type="radio" name="tipo_atendimento" value="Emergência" required> 🚑Emergência
        <input type="radio" name="tipo_atendimento" value="Internação" class="ms-3"> 🛌Internação
    </div>
</div>
<div class="row">
    <div class="col-md-3 mb-3">
        <label class="form-label"> 📧 Email</label>
        <input type="email" name="email" class="form-control">
    </div>
    <div class="col-md-3 mb-3">
        <label class="form-label"> 📞 Telefone</label>
        <input type="text" name="telefone" class="form-control">
    </div>
    <div class="col-md-3 mb-3">
        <label class="form-label d-block">Tipo de Solicitação</label>
        <input type="radio" name="tipo_solicitacao" value="SUS" required> SUS
        <input type="radio" name="tipo_solicitacao" value="Particular" class="ms-2"> Particular
        <input type="radio" name="tipo_solicitacao" value="Jurídico" class="ms-2"> Jurídico
    </div>
    <div class="col-md-3 mb-3">
        <label class="form-label d-block">Tipo de Cópia</label>
        <input type="radio" name="tipo_copia" value="Integral" required> Integral
        <input type="radio" name="tipo_copia" value="Parcial" class="ms-2"> Parcial
    </div>
</div>
<div class="row align-items-end">
    <div class="col-md-3 mb-3">
        <label class="form-label d-block">Tipo de Envio</label>
        <input type="radio" id="fisicoRadio" name="tipo_envio" value="Físico" onchange="toggleFolhas()" required> 📁Físico
        <input type="radio" id="digitalRadio" name="tipo_envio" value="Digital" onchange="toggleFolhas()" class="ms-2"> 💻Digital
    </div>
    <div class="col-md-2 mb-3" id="campoFolhas">
        <label class="form-label">Nº de Folhas</label>
        <input type="number" name="folhas" class="form-control">
    </div>
    <div class="col-md-2 mb-3">
        <label class="form-label">Data do Atendimento</label>
        <input type="date" name="data_atendimento" class="form-control">
    </div>
    <div class="col-md-2 mb-3">
        <label class="form-label">Data da Alta</label>
        <input type="date" name="data_alta" class="form-control">
    </div>
    <div class="col-md-3 mb-3">
        <label class="form-label">Data da Solicitação</label>
        <input type="date" name="data_solicitacao" class="form-control">
    </div>
</div>
<div class="row">
    <div class="col-md-8 mb-3">
        <label class="form-label">Observações</label>
        <textarea name="observacoes" class="form-control" rows="4"></textarea>
    </div>
    <div class="col-md-4 mb-3">
        <label class="form-label d-block">Solicitado por</label>
        <div class="mt-2"><input type="radio" name="solicitado_por" value="Paciente" required> Paciente</div>
        <div><input type="radio" name="solicitado_por" value="Familiar/Responsável"> Familiar/Responsável</div>
        <div><input type="radio" name="solicitado_por" value="Procurador"> Procurador</div>
        <div><input type="radio" name="solicitado_por" value="Paciente falecido"> Paciente falecido</div>
        <div><input type="radio" name="solicitado_por" value="Interno"> Interno</div>
    </div>
</div>
<hr>
<div class="text-center">
    <button class="btn btn-primary btn-lg">💾 Salvar Solicitação</button>
    <a href="/" class="btn btn-secondary btn-lg">⬅ Voltar</a>
</div>
</form>
        </div>
    </div>
    <script>
    function toggleFolhas() {
        const digital = document.getElementById("digitalRadio");
        const campo = document.getElementById("campoFolhas");
        if (digital && digital.checked) {
            campo.style.display = "none";
        } else {
            campo.style.display = "block";
        }
    }
    window.onload = toggleFolhas;
    </script>
    </body>
    </html>
    """

@app.route("/salvar", methods=["POST"])
def salvar():
    tipo = request.form.get("tipo_envio")
    folhas = request.form.get("folhas")

    if tipo == "Digital":
        folhas = 0
        valor = 0
    else:
        if not folhas:
            folhas = 0
        valor = float(folhas) * 0.25

    data_solicitacao = request.form.get("data_solicitacao")

    if data_solicitacao:
        data_base = datetime.strptime(data_solicitacao, "%Y-%m-%d")
    else:
        data_base = datetime.now()

    data_entrega = adicionar_dias_uteis(data_base, DIAS_PRAZO)

    solicitacao = {
        "tipo_atendimento": request.form.get("tipo_atendimento"),
        "tipo_copia": request.form.get("tipo_copia"),
        "tipo_solicitacao": request.form.get("tipo_solicitacao"),
        "tipo_envio": request.form.get("tipo_envio"),
        "solicitado_por": request.form.get("solicitado_por"),
        "prontuario": request.form.get("prontuario"),
        "atendimento": request.form.get("atendimento"),
        "nome": request.form.get("nome"),
        "telefone": request.form.get("telefone"),
        "email": request.form.get("email"),
        "folhas": folhas,
        "valor": valor,
        "data_atendimento": request.form.get("data_atendimento"),
        "data_alta": request.form.get("data_alta"),
        "data_solicitacao": data_solicitacao,
        "data_base": data_base.strftime("%Y-%m-%d"),
        "data_entrega": data_entrega.strftime("%Y-%m-%d"),
        "data_pronto": "",
        "data_entregue": "",
        "status": "⏳ Pendente",
        "observacoes": request.form.get("observacoes", "")
    }

    conn = sqlite3.connect("docmed.db")
    c = conn.cursor()

    c.execute("""
INSERT INTO solicitacoes (
    nome, prontuario, atendimento, telefone, email,
    tipo_atendimento, tipo_envio, tipo_copia, tipo_solicitacao,
    data_atendimento, data_alta, folhas, valor,
    data_solicitacao, data_entrega, data_pronto, data_entregue,
    status, observacoes, solicitado_por
)
VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
""", (
    solicitacao["nome"],
    solicitacao["prontuario"],
    solicitacao["atendimento"],
    solicitacao["telefone"],
    solicitacao["email"],
    solicitacao["tipo_atendimento"],
    solicitacao["tipo_envio"],
    solicitacao["tipo_copia"],
    solicitacao["tipo_solicitacao"],
    solicitacao["data_atendimento"],
    solicitacao["data_alta"],
    solicitacao["folhas"],
    solicitacao["valor"],
    solicitacao["data_solicitacao"],
    solicitacao["data_entrega"],
    solicitacao["data_pronto"],
    solicitacao["data_entregue"],
    solicitacao["status"],
    solicitacao["observacoes"],
    solicitacao["solicitado_por"]
))

    conn.commit()

    salvar_historico(
        solicitacao_id=c.lastrowid,
        acao="📌 Solicitação cadastrada",
        usuario=session.get("usuario", "Sistema")
    )

    conn.close()
    
    return f"""
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <body style="background: #f0f5fa; color: #000000; font-family: sans-serif; padding: 40px;">
        <div class="card p-4 mx-auto shadow" style="max-width: 500px; background: white; border-radius: 12px;">
            <h2 style="color: #0d9488;">✅ Solicitação Salva!</h2>
            <p>📅 Prazo de entrega: <b>{data_entrega.strftime("%d/%m/%Y")}</b></p>
            <a href="/" class="btn btn-primary mt-2" style="background-color: #1e3a8a;">Voltar para o DocMed</a>
        </div>
    </body>
    """

@app.route("/detalhes/<int:id>")
def detalhes(id):
    conn = sqlite3.connect("docmed.db")
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("SELECT * FROM solicitacoes WHERE id = ?", (id,))
    s = c.fetchone()
    conn.close()

    if not s:
        return "<h1>❌ Solicitação não encontrada</h1>"

    return f"""
    <html>
    <head>
        <title>Detalhes</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
        <style>
            body {{ background: #f0f5fa; color: #000000; font-family: 'Segoe UI', Arial, sans-serif; }}
.card {{ background: #ffffff; color: #000000; border-radius: 12px; }}
        </style>
    </head>
    <body>
    <div class="container mt-4">
        <div class="card shadow p-4">
       <h3 style="color: #1e3a8a;" class="mb-4">🔍 Detalhes da Solicitação</h3>
       <p><b>Nome:</b> {s["nome"]}</p>
<p><b>Email:</b> {s["email"]}</p>
<p><b>Telefone:</b> {s["telefone"]}</p>
<p><b>Prontuário:</b> {s["prontuario"]}</p>
<p><b>Atendimento:</b> {s["atendimento"]}</p>
<p><b>Tipo de Atendimento:</b> {s["tipo_atendimento"]}</p>
<p><b>Data do Atendimento:</b> {formatar_data(s["data_atendimento"])}</p>
<p><b>Data da Alta:</b> {formatar_data(s["data_alta"])}</p>
<p><b>Data da Solicitação:</b> {formatar_data(s["data_solicitacao"])}</p>
<p><b>Tipo de Solicitação:</b> {s["tipo_solicitacao"]}</p>
<p><b>Tipo de Envio:</b> {s["tipo_envio"]}</p>
<p><b>Tipo de Cópia:</b> {s["tipo_copia"]}</p>
<p><b>Nº de Folhas:</b> {s["folhas"]}</p>
<p><b>Valor:</b> R$ {float(s["valor"] or 0):.2f}</p>
<p><b>Prazo para Entrega:</b> <b>{formatar_data(s["data_entrega"])}</b></p>
<p><b>Observações:</b> {s["observacoes"] or ""}</p>
<p><b>Status:</b> {s["status"]}</p>
<a href="/" class="btn btn-secondary mt-3" style="max-width: 150px;">⬅ Voltar</a>    
        </div>
    </div>
    </body>
    </html>
    """

@app.route("/excluir/<int:id>", methods=["GET", "POST"])
def excluir(id):
    if session.get("perfil") != "Administrador":
        return """
        <h2>❌ Apenas administradores podem excluir solicitações</h2>
        <a href="/">Voltar</a>
        """

    if request.method == "POST":
        justificativa = request.form.get("justificativa", "").strip()

        if not justificativa:
            return """
            <h2>❌ Informe uma justificativa</h2>
            <a href="/">Voltar</a>
            """

        conn = sqlite3.connect("docmed.db")
        c = conn.cursor()
        c.execute("SELECT * FROM solicitacoes WHERE id = ?", (id,))
        s = c.fetchone()

        if not s:
            conn.close()
            return """
            <h2>❌ Solicitação não encontrada</h2>
            <a href="/">Voltar</a>
            """

        c.execute("""
        INSERT INTO solicitacoes_excluidas (
            id, nome, prontuario, atendimento, telefone, email, tipo_atendimento, tipo_envio, tipo_copia, tipo_solicitacao,
            data_atendimento, data_alta, folhas, valor, data_solicitacao, data_entrega, data_pronto, data_entregue, status, observacoes,
            excluido_por, data_exclusao, justificativa
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            s[0], s[1], s[2], s[3], s[4], s[5], s[6], s[7], s[8], s[9], s[10], s[11], s[12], s[13], s[14], s[15], s[16], s[17], s[18], s[19],
            session.get("usuario"), datetime.now().strftime("%d/%m/%Y %H:%M"), justificativa
        ))

        salvar_historico(
            solicitacao_id=id,
            acao=f"🗑️ Solicitação excluída. Motivo: {justificativa}",
            usuario=session.get("usuario", "Sistema")
        )

        c.execute("DELETE FROM solicitacoes WHERE id = ?", (id,))
        conn.commit()
        conn.close()
        return redirect("/")

    return """
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <body style="background: #f0f5fa; color: #000000;">
    <div class="container mt-4">
        <div class="card shadow p-4" style="background: white;">
            <h2 style="color: #1e3a8a;">🔒 Exclusão Restrita</h2>
            <p>Informe a justificativa para exclusão da solicitação.</p>
            <form method="POST">
                <label class="form-label">Justificativa</label>
                <textarea name="justificativa" class="form-control" rows="4" required></textarea>
                <br>
                <button class="btn btn-danger">🗑️ Confirmar Exclusão</button>
                <a href="/" class="btn btn-secondary">Cancelar</a>
            </form>
        </div>
    </div>
    </body>
    """

@app.route("/editar/<int:id>", methods=["GET", "POST"])
def editar(id):
    conn = sqlite3.connect("docmed.db")
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("SELECT * FROM solicitacoes WHERE id = ?", (id,))
    s = c.fetchone()

    if not s:
        conn.close()
        return "Solicitação não encontrada"
    
    if request.method == "POST":
        c.execute("""
        UPDATE solicitacoes
        SET status = ?, observacoes = ?, data_pronto = ?, data_entregue = ?
        WHERE id = ?
        """, (
            request.form.get("status"), request.form.get("observacoes", ""),
            request.form.get("data_pronto", ""), request.form.get("data_entregue", ""), id
        ))
        conn.commit()

        status = request.form.get("status")
        if status == "📦 Entregue":
            salvar_historico(solicitacao_id=id, acao="📦 Prontuário entregue ao paciente", usuario=session.get("usuario", "Sistema"))
        elif status == "✅ Autorizado":
            salvar_historico(solicitacao_id=id, acao="✅ Solicitação autorizada", usuario=session.get("usuario", "Sistema"))
        else:
            salvar_historico(solicitacao_id=id, acao="✏️ Solicitação editada", usuario=session.get("usuario", "Sistema"))

        conn.close()
        return redirect("/")

    conn.close()
       
    return f"""
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
<body style="background: #f0f5fa; color: #000000;">
<div class="container mt-4">
    <div class="card shadow p-4" style="background: white;">
        <h2 style="color: #1e3a8a;">✏️ Editar Solicitação</h2>
<form method="POST">
    <label class="form-label">Status</label>
<select name="status" class="form-select">
    <option value="⏳ Pendente">⏳ Pendente</option>
    <option value="✅ Autorizado">✅ Autorizado</option>
    <option value="❌ Cancelado">❌ Cancelado</option>
</select>
    <br>
    <div id="motivoBox" style="display:none;">
        <label>Motivo do Cancelamento</label>
        <input type="text" name="motivo_cancelamento" class="form-control" value="">
    </div>
    <br>
    <label>Data Pronto</label>
<input type="date" name="data_pronto" class="form-control" value="{s['data_pronto'] if s['data_pronto'] else ''}">
<br>
<label>Data Entregue</label>
<input type="date" name="data_entregue" class="form-control" value="{s['data_entregue'] if s['data_entregue'] else ''}">
    <br>
    <label>Observações</label>
    <textarea name="observacoes" class="form-control">{s['observacoes'] if s['observacoes'] else ''}</textarea>
    <br>
    <button class="btn btn-success" style="background-color: #0d9488; border-color: #0d9488;">💾 Salvar</button>
    <a href="/" class="btn btn-secondary">Voltar</a>
</form>
</div>
</div>
</body>
<script>
function toggleMotivo() {{
    var status = document.getElementById("status").value;
    var box = document.getElementById("motivoBox");
    box.style.display = (status == "❌ Cancelado") ? "block" : "none";
}}
toggleMotivo();
</script>
"""

@app.route("/marcar_pronto/<int:id>")
def marcar_pronto(id):
    hoje = date.today().strftime("%Y-%m-%d")
    conn = sqlite3.connect("docmed.db")
    c = conn.cursor()
    c.execute("UPDATE solicitacoes SET data_pronto = ?, status = ? WHERE id = ?", (hoje, "📄 Pronto para retirada", id))
    conn.commit()
    conn.close()
    return redirect("/")

@app.route("/marcar_entregue/<int:id>")
def marcar_entregue(id):
    hoje = date.today().strftime("%Y-%m-%d")
    conn = sqlite3.connect("docmed.db")
    c = conn.cursor()
    c.execute("UPDATE solicitacoes SET data_entregue = ?, status = ? WHERE id = ?", (hoje, "📦 Entregue", id))
    conn.commit()
    conn.close()
    return redirect("/")

@app.route("/historico/<int:id>")
def historico(id):
    conn = sqlite3.connect("docmed.db")
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("SELECT * FROM historico WHERE solicitacao_id = ? ORDER BY id DESC", (id,))
    logs = c.fetchall()
    conn.close()

    html = """
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <body style="background: #f0f5fa; color: #000000;">
    <div class="container mt-4">
        <div class="card shadow p-4" style="background: white;">
        <h2 style="color: #1e3a8a;">📜 Linha do Tempo da Solicitação</h2>
        <hr>
        <div class="timeline">
    """

    for h in logs:
        html += f"""
        <div style="border-left: 3px solid #1e3a8a; margin-left: 10px; padding-left: 15px; margin-bottom: 15px;">
            <div style="font-size:14px; color:gray;">🕒 {h['data']}</div>
            <div style="font-weight:bold; color: #000000;">👤 {h['usuario']}</div>
            <div style="color: #000000;">📝 {h['acao']}</div>
        </div>
        """

    html += """
        </div>
        <a href="/" class="btn btn-secondary mt-3">Voltar</a>
        </div>
    </div>
    </body>
    """
    return html

@app.route("/lixeira")
def lixeira():
    if "usuario" not in session:
        return redirect("/login")

    if session.get("perfil") != "Administrador":
        return """
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
        <div class="container mt-4 text-center">
            <h2>⛔ Acesso Negado</h2>
            <p>Apenas administradores podem acessar a lixeira.</p>
            <a href="/" class="btn btn-primary">Voltar para o Início</a>
        </div>
        """

    conn = sqlite3.connect("docmed.db")
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("SELECT * FROM solicitacoes_excluidas ORDER BY data_exclusao DESC")
    excluidas = c.fetchall()
    conn.close()

    html = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>Lixeira - DocMed</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
        <style>
            body { background: #f0f5fa; color: #000000; font-family: Arial, sans-serif; }
            .card { background: #ffffff; border-radius: 14px; border: 1px solid #cbd5e1; box-shadow: 0 2px 10px rgba(0,0,0,0.05); }
        </style>
    </head>
    <body>
    <div class="container mt-4">
        <div class="d-flex justify-content-between align-items-center mb-3">
            <div>
                <h1 style="color: #1e3a8a;">🗑️ Lixeira de Solicitações</h1>
                <h6 class="text-muted">Histórico de registros removidos e auditados</h6>
            </div>
            <a href="/" class="btn btn-secondary">⬅ Voltar ao Início</a>
        </div>
        <hr>
        <div class="card p-4 shadow-sm">
            <table class="table table-striped table-hover">
                <thead class="table-dark" style="background-color: #1e3a8a;">
                    <tr>
                        <th>ID Orig.</th>
                        <th>Prontuário</th>
                        <th>Nome do Paciente</th>
                        <th>Excluído Por</th>
                        <th>Data Exclusão</th>
                        <th>Justificativa</th>
                    </tr>
                </thead>
                <tbody>
    """

    if not excluidas:
        html += """
                    <tr>
                        <td colspan="6" class="text-center text-muted py-4">Nenhum registro na lixeira.</td>
                    </tr>
        """
    else:
        for e in excluidas:
            html += f"""
                    <tr>
                        <td>{e["id"]}</td>
                        <td>{e["prontuario"]}</td>
                        <td>{e["nome"]}</td>
                        <td><span class="badge bg-danger">{e["excluido_por"]}</span></td>
                        <td>{e["data_exclusao"]}</td>
                        <td class="text-danger"><i>{e["justificativa"]}</i></td>
                    </tr>
            """

    html += """
                </tbody>
            </table>
        </div>
    </div>
    </body>
    </html>
    """
    return html

@app.route("/novo_usuario", methods=["GET", "POST"])
def novo_usuario():
    if "usuario" not in session:
        return redirect("/login")

    if session.get("perfil") != "Administrador":
        return """
        <h2>⛔ Acesso Negado</h2>
        <p>Apenas administradores podem criar usuários.</p>
        <a href="/">Voltar</a>
        """

    if request.method == "POST":
        nome = request.form.get("nome")
        usuario = request.form.get("usuario")
        senha = request.form.get("senha")
        perfil = request.form.get("perfil")

        conn = sqlite3.connect("docmed.db")
        c = conn.cursor()

        try:
            c.execute("INSERT INTO usuarios (nome, usuario, senha, perfil) VALUES (?, ?, ?, ?)", (nome, usuario, senha, perfil))
            conn.commit()
        except:
            conn.close()
            return "<h2>❌ Usuário já existe</h2><a href='/novo_usuario'>Voltar</a>"

        conn.close()
        return redirect("/")

    return """
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <body style="background: #f0f5fa; color: #000000;">
    <div class="container mt-4">
        <div class="card shadow p-4" style="background: white;">
            <h2 style="color: #1e3a8a;">👤 Novo Usuário</h2>
            <form method="POST">
                <label>Nome</label>
                <input type="text" name="nome" class="form-control" required>
                <br>
                <label>Usuário</label>
                <input type="text" name="usuario" class="form-control" required>
                <br>
                <label>Senha</label>
                <input type="password" name="senha" class="form-control" required>
                <br>
                <label>Perfil</label>
                <select name="perfil" class="form-control">
                    <option>Administrador</option>
                    <option>Atendente</option>
                </select>
                <br>
                <button class="btn btn-success" style="background-color: #0d9488; border-color: #0d9488;">💾 Salvar Usuário</button>
                <a href="/" class="btn btn-secondary">Voltar</a>
            </form>
        </div>
    </div>
    </body>
    """

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        usuario = request.form.get("usuario")
        senha = request.form.get("senha")
        user = verificar_login(usuario, senha)

        if user:
            session["usuario"] = user[1]   
            session["perfil"] = user[4]    
            return redirect("/")
        else:
            return "<h3>❌ Login inválido</h3><a href='/login'>Voltar</a>"

    return """
<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="UTF-8">
<title>Login DocMed</title>
<style>
body { margin: 0; height: 100vh; display: flex; justify-content: center; align-items: center; background: linear-gradient(135deg, #1e3a8a, #0f172a); }
.overlay { position: absolute; width: 100%; height: 100%; background: rgba(0,0,0,0.3); }
@keyframes fadeUp { from { opacity: 0; transform: translateY(25px); } to { opacity: 1; transform: translateY(0); } }
.login-box { position: relative; z-index: 2; width: 320px; padding: 25px; background: rgba(255,255,255,0.9); border-radius: 12px; text-align: center; box-shadow: 0 10px 25px rgba(0,0,0,0.2); animation: fadeUp 0.6s ease-out; color: #000000; }
h2 { color: #1e3a8a; margin-bottom: 20px; font-family: sans-serif; }
input { width: 100%; padding: 12px; margin: 8px 0; border: 1px solid #cbd5e1; border-radius: 6px; box-sizing: border-box; color: #000000; }
button { width: 100%; padding: 12px; margin-top: 10px; border: none; border-radius: 6px; background: #1e3a8a; color: white; cursor: pointer; box-sizing: border-box; font-weight: bold; }
button:hover { background: #1d4ed8; }
a { display: block; margin-top: 10px; font-size: 13px; color: #475569; text-decoration: underline; }
</style>
</head>
<body>
<div class="overlay"></div>
<div class="login-box">
    <h2>Bem-vindo(a) ao DocMed</h2>
    <form method="POST">
        <input name="usuario" placeholder="Usuário" required>
        <input name="senha" type="password" placeholder="Senha" required>
        <button type="submit" id="loginBtn" onclick="mostrarLoading()">
            <span id="btnText">🔐 Entrar</span>
        </button>
    </form>
    <a href="#">Esqueci minha senha</a>
</div>
<script>
function mostrarLoading() {
    var btn = document.getElementById("loginBtn");
    var btnText = document.getElementById("btnText");
    btnText.innerText = "⏳ Entrando...";
    btn.style.opacity = "0.7";
    btn.style.cursor = "not-allowed";
}
</script>
</body>
</html>
"""

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")

if __name__ == "__main__":
    init_db()
    app.run(debug=True)