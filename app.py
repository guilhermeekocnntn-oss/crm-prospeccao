import sys
import os

# Ajusta encodificação do stdout para UTF-8 no Windows
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

from flask import Flask, render_template, request, jsonify, redirect, url_for
from google_sheets import obter_todos_leads_eko, atualizar_lead_tempo_real, ESTADOS_MAP
from database import importar_planilhas_pasta, obter_todos_leads_db, adicionar_lead_db, buscar_cidades_ibge

app = Flask(__name__)
app.secret_key = "crm_prospeccao_secret_key"

# Adicione esta linha no início do app.py:
CACHE_FILE = "/tmp/cache_leads.json"

# Variável global para guardar os dados na memória RAM
LEADS_MEMORIA = []

def carregar_dados_iniciais(force_refresh=False):
    global LEADS_MEMORIA
    print("[INFO] Carregando dados de leads para a memória RAM...")
    
    # 1. Carrega o banco SQLite local instantaneamente (milissegundos)
    leads_db = []
    try:
        leads_db = obter_todos_leads_db()
    except Exception as e:
        print(f"[AVISO] Erro ao carregar dados do SQLite: {e}")

    # 2. No Vercel, lê o cache instantâneo se existir. Só faz varredura completa se force_refresh=True ou local.
    leads_drive = []
    if os.path.exists(CACHE_FILE) or force_refresh or not os.environ.get("VERCEL"):
        try:
            leads_drive = obter_todos_leads_eko(force_refresh=force_refresh)
        except Exception as e:
            print(f"[AVISO] Erro ao carregar dados do Google Drive: {e}")
    
    LEADS_MEMORIA = leads_drive + leads_db
    print(f"[OK] Total unificado na memória: {len(LEADS_MEMORIA)} leads.")
    return LEADS_MEMORIA

import json

def obter_cidades_por_uf(leads):
    cidades_map = {}
    for l in leads:
        uf = (l.get('uf') or 'OUTROS').strip().upper()
        cidade = (l.get('cidade') or l.get('aba') or '').strip()
        if uf and cidade:
            if uf not in cidades_map:
                cidades_map[uf] = set()
            cidades_map[uf].add(cidade)
    return {uf: sorted(list(cidades)) for uf, cidades in cidades_map.items()}

def obter_tipos_e_status(leads):
    tipos_set = set()
    status_set = set()
    
    for l in leads:
        t = str(l.get('tipo') or '').strip()
        if t and t.lower() not in ('none', '', '-'):
            tipos_set.add(t)
            
        s = str(l.get('status') or '').strip()
        if s and s.lower() not in ('none', '', '-'):
            status_set.add(s)

    # Status padrão comuns
    status_padrao = [
        "A Ligar (Novo)",
        "Não Atendeu (T1)",
        "Qualificado (Em Negociação)",
        "Cliente Fechado",
        "Sem Interesse",
        "Retorno Agendado"
    ]
    for sp in status_padrao:
        status_set.add(sp)

    # Tipos padrão se a lista for pequena
    tipos_padrao = ["Distribuidora", "Salão", "Rede de Salões", "Varejo", "Atacado", "Outros"]
    for tp in tipos_padrao:
        tipos_set.add(tp)

    return sorted(list(tipos_set)), sorted(list(status_set))

@app.route('/api/cidades/<uf>')
def api_cidades(uf):
    """Retorna a lista de cidades para a UF — sempre mescla abas das planilhas + todas as cidades IBGE"""
    uf_upper = uf.upper()
    cidades_planilha = obter_cidades_por_uf(LEADS_MEMORIA).get(uf_upper, [])

    cidades_ibge = []
    try:
        dados_ibge = buscar_cidades_ibge(uf_upper)
        cidades_ibge = [d['cidade'] for d in dados_ibge]
    except Exception as e:
        print(f"Erro ao buscar IBGE para {uf_upper}: {e}")

    # Mescla: abas da planilha aparecem primeiro (marcadas com ★), depois o restante do IBGE
    cidades_planilha_set = set(cidades_planilha)
    cidades_ibge_extras = [c for c in cidades_ibge if c not in cidades_planilha_set]

    # Abas existentes ficam no topo com marcador visual
    cidades_com_aba = [f"★ {c}" for c in sorted(cidades_planilha)]
    cidades_finais = cidades_com_aba + sorted(cidades_ibge_extras)

    if not cidades_finais:
        cidades_finais = ["Geral"]

    return jsonify({"uf": uf_upper, "cidades": cidades_finais, "fonte": "misto"})

@app.route('/')
def index():
    global LEADS_MEMORIA
    
    # Se a memória estiver vazia, recarrega
    if not LEADS_MEMORIA:
        carregar_dados_iniciais()

    # Métricas Gerais
    total_leads = len(LEADS_MEMORIA)
    total_agendados = sum(
        1 for l in LEADS_MEMORIA 
        if 'agendado' in str(l.get('status', '')).lower() or 'retorno' in str(l.get('status', '')).lower()
    )
    total_fechados = sum(
        1 for l in LEADS_MEMORIA 
        if 'fechado' in str(l.get('status', '')).lower()
    )

    # Parâmetros dos filtros recebidos do formulário
    busca_raw = request.args.get('q', request.args.get('busca', '')).strip()
    busca_termo = busca_raw.lower()
    estado_filtro = request.args.get('uf', request.args.get('estado', '')).strip()
    cidade_filtro = request.args.get('cidade', '').strip()
    tipo_filtro = request.args.get('tipo', '').strip()
    status_filtro = request.args.get('status', '').strip()
    potencial_filtro = request.args.get('potencial', '').strip()

    # Mapeamentos e Listas dinâmicas extraídas das planilhas do projeto
    cidades_por_uf = obter_cidades_por_uf(LEADS_MEMORIA)
    tipos_disponiveis, status_disponiveis = obter_tipos_e_status(LEADS_MEMORIA)

    # FILTRAGEM DIRETO DA MEMÓRIA RAM
    leads_filtrados = LEADS_MEMORIA

    if estado_filtro:
        leads_filtrados = [
            l for l in leads_filtrados 
            if l.get('uf') == estado_filtro or l.get('uf_nome') == estado_filtro
        ]

    if cidade_filtro:
        leads_filtrados = [
            l for l in leads_filtrados 
            if l.get('cidade') == cidade_filtro or l.get('aba') == cidade_filtro
        ]

    if tipo_filtro:
        leads_filtrados = [
            l for l in leads_filtrados 
            if l.get('tipo') == tipo_filtro
        ]

    if status_filtro:
        leads_filtrados = [
            l for l in leads_filtrados 
            if l.get('status') == status_filtro
        ]

    if potencial_filtro:
        leads_filtrados = [
            l for l in leads_filtrados 
            if l.get('potencial') == potencial_filtro
        ]

    if busca_termo:
        leads_filtrados = [
            l for l in leads_filtrados 
            if busca_termo in str(l.get('empresa', '')).lower() 
            or busca_termo in str(l.get('cidade', '')).lower()
            or busca_termo in str(l.get('bairro', '')).lower()
            or busca_termo in str(l.get('telefone', '')).lower()
            or busca_termo in str(l.get('decisor', '')).lower()
        ]

    return render_template(
        'dashboard.html', 
        leads=leads_filtrados, 
        total_leads=total_leads,
        total_agendados=total_agendados,
        total_fechados=total_fechados,
        estados=ESTADOS_MAP,
        cidades_por_uf=cidades_por_uf,
        cidades_json=json.dumps(cidades_por_uf, ensure_ascii=False),
        tipos=tipos_disponiveis,
        status_lista=status_disponiveis,
        f_busca=busca_raw,
        f_uf=estado_filtro,
        f_cidade=cidade_filtro,
        f_tipo=tipo_filtro,
        f_status=status_filtro,
        f_potencial=potencial_filtro
    )

@app.route('/atualizar_lead', methods=['POST'])
def atualizar():
    dados = request.json or {}
    sheet_id = dados.get('sheet_id')
    nome_aba = dados.get('aba')
    linha = dados.get('linha_id')
    coluna = dados.get('coluna')
    novo_valor = dados.get('valor')

    # Atualiza imediatamente na memória RAM do site
    global LEADS_MEMORIA
    for lead in LEADS_MEMORIA:
        if str(lead.get('sheet_id')) == str(sheet_id) and str(lead.get('aba')) == str(nome_aba) and str(lead.get('linha_id')) == str(linha):
            lead[coluna] = novo_valor

    # Atualiza no Google Drive se for do Drive
    if sheet_id and not str(sheet_id).startswith('db_'):
        sucesso = atualizar_lead_tempo_real(sheet_id, nome_aba, linha, coluna, novo_valor)
    else:
        sucesso = True # Registro local atualizado em memória
        
    return jsonify({"sucesso": sucesso})

@app.route('/adicionar_lead', methods=['POST'])
def adicionar_lead():
    """Rota para cadastrar um novo lead individual no CRM"""
    dados = request.form.to_dict() if request.form else (request.json or {})
    novo_lead = adicionar_lead_db(dados)
    
    global LEADS_MEMORIA
    LEADS_MEMORIA.insert(0, novo_lead) # Adiciona no topo da lista na memória
    
    if request.is_json:
        return jsonify({"sucesso": True, "lead": novo_lead})
    return redirect(url_for('index'))

@app.route('/sincronizar_drive')
def sincronizar_drive():
    """Botão/Rota opcional para forçar a atualização completa vinda do Google Drive"""
    carregar_dados_iniciais(force_refresh=True)
    return jsonify({"status": "sucesso", "total": len(LEADS_MEMORIA)})

@app.route('/importar', methods=['POST'])
def importar():
    """Rota para importar planilhas locais para o banco de dados"""
    caminho_pasta = request.form.get('caminho_pasta', '').strip()
    if caminho_pasta:
        qtd = importar_planilhas_pasta(caminho_pasta)
        if qtd > 0:
            carregar_dados_iniciais(force_refresh=False)
    return redirect(url_for('index'))

if __name__ == '__main__':
    carregar_dados_iniciais()
    app.run(debug=True, port=5000)