#!/usr/bin/env python3
"""
etl.py — ETL unificado para dashboards de conformidade We Handle / SABESP
Uso:
  python etl.py --projeto eteca --csv relatorio_docs.csv [--csv-int status_int.csv]
  python etl.py --projeto eteca --csv relatorio_docs.xlsx [--csv-int status_int.csv]

Gera em data/<projeto>/:
  YYYY-MM-DD.json  — dados completos da extração
  historico.json   — série compacta para gráficos de evolução
  manifest.json    — lista de datas disponíveis
"""

import argparse, collections, csv, json, os, re, sys
import datetime

# ─── CLI ─────────────────────────────────────────────────────────────────────

def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument('--projeto', required=True, help='ID do projeto (ex: eteca)')
    p.add_argument('--csv',     required=True, dest='csv_doc',
                   help='CSV ou XLSX de documentos (relatorio_de_documentos_...)')
    p.add_argument('--csv-int', dest='csv_int', default=None,
                   help='CSV de integração (status_integracao_...) — opcional')
    return p.parse_args()

# ─── LEITURA DE ARQUIVOS ──────────────────────────────────────────────────────

def read_tabular(path):
    """Lê CSV (utf-8-sig) ou XLSX e retorna list[dict]."""
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.xlsx', '.xls'):
        try:
            import openpyxl
        except ImportError:
            sys.exit('openpyxl não instalado. Execute: pip install openpyxl')
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.values)
        wb.close()
        if not rows:
            return []
        headers = [str(h).strip() if h is not None else '' for h in rows[0]]
        result = []
        for row in rows[1:]:
            d = {}
            for h, v in zip(headers, row):
                d[h] = str(v).strip() if v is not None else None
            result.append(d)
        return result
    else:
        with open(path, encoding='utf-8-sig', newline='') as f:
            return list(csv.DictReader(f))


def load_config(projeto):
    config_path = os.path.join('projects', projeto, 'config.json')
    if not os.path.exists(config_path):
        sys.exit(f'config.json não encontrado: {config_path}')
    with open(config_path, encoding='utf-8') as f:
        return json.load(f)


def load_historico(projeto):
    path = os.path.join('data', projeto, 'historico.json')
    if os.path.exists(path):
        with open(path, encoding='utf-8') as f:
            return json.load(f)
    return []


def save_json(path, obj):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    content = json.dumps(obj, ensure_ascii=False, indent=2)
    # Abre em modo binário com truncate explícito para evitar null bytes residuais
    with open(path, 'wb') as f:
        f.truncate(0)
        f.write(content.encode('utf-8'))

# ─── CONSTANTES DE NEGÓCIO ───────────────────────────────────────────────────

NC_STATUSES = {
    'Pendente', 'Pendente - Isenção Aguardando Aprovação', 'Em validação',
    'Inválido', 'Vencido', 'Reprovado pelo cliente', 'Pendente - Isenção Reprovada'
}
FOCUS = {'Pendente', 'Inválido', 'Pendente - Isenção Reprovada', 'Vencido', 'Reprovado pelo cliente'}
CRIT  = {'Inválido', 'Vencido', 'Reprovado pelo cliente', 'Pendente - Isenção Reprovada'}
CI_OK = {'Acesso Liberado', 'Pendências porém Integrado'}
INT_PRIORITY = {'Acesso Liberado': 0, 'Pendências porém Integrado': 1, 'Falta Integração': 2, 'Pendências e Falta Integração': 3}

CONF_TARGET = 95
ADER_TARGET = 99
CP_TARGET   = 95
CI_TARGET   = 95

STATUS_ROWS_DEFS = [
    ('Isenção aprovada',       'Isento',                                     '#16a34a', False, False),
    ('Aprovado',               'OK',                                          '#16a34a', False, False),
    ('Aprovado com ressalva',  'Liberado com ressalvas no cliente',           '#16a34a', False, False),
    ('Aprovado pelo cliente',  'Aprovado pelo cliente',                       '#16a34a', False, False),
    ('Isenção em análise',     'Pendente - Isenção Aguardando Aprovação',    '#2563a8', True,  False),
    ('Pendente',               'Pendente',                                    '#d97706', True,  True ),
    ('Em validação',           'Em validação',                                '#2563a8', True,  False),
    ('Perto do vencimento',    'Perto do Vencimento',                         '#d97706', False, False),
    ('Reprovado',              'Inválido',                                    '#dc2626', True,  False),
    ('Vencido',                'Vencido',                                     '#dc2626', True,  False),
    ('Reprovado pelo cliente', 'Reprovado pelo cliente',                      '#dc2626', True,  False),
    ('Isenção reprovada',      'Pendente - Isenção Reprovada',               '#dc2626', True,  False),
]

NC_STATUS_DEFS = [
    ('Pendente',                                'Pendente',              '#d97706'),
    ('Inválido',                                'Reprovado',             '#dc2626'),
    ('Vencido',                                 'Vencido',               '#dc2626'),
    ('Reprovado pelo cliente',                  'Reprovado pelo cliente','#dc2626'),
    ('Pendente - Isenção Reprovada',            'Isenção reprovada',     '#dc2626'),
    ('Pendente - Isenção Aguardando Aprovação', 'Isenção em análise',    '#2563a8'),
    ('Em validação',                            'Em validação',          '#2563a8'),
]

DOC_SHORT = {
    'contrato de prestação de serviço': 'Contrato de Prestação de Serviço',
    'certidão negativa de débito municipal (iptu e iss)': 'CND Municipal',
    'certidão negativa de débito federal - créditos tributários': 'CND Federal',
    'certidão negativa de débito fgts - crf': 'CND FGTS — CRF',
    'cnd fgts crf': 'CND FGTS — CRF',
    'certidão negativa de débito trabalhista - cndt': 'CNDT',
    'cnd negativa de débitos trabalhistas (cndt)': 'CNDT',
    'contrato social consolidado - inicial e atualizações ou última alteração contratual': 'Contrato Social Consolidado',
    'contrato social consolidado - inicial e atualizações ou registro de microempreendedor individual mei - requerimento de empresario - estatuto social consolidado': 'Contrato Social / Reg. MEI',
    'gps/darf - dctfweb': 'GPS/DARF — DCTFweb',
    'comprovante gps/darf (inss)': 'Comp. GPS/DARF (INSS)',
    'comprovante de pagamento gps,darf (inss)': 'Comp. de Pagamento GPS/DARF',
    'comprovante fgts digital (gfd)': 'Comp. FGTS Digital (GFD)',
    'gfd - guia do fgts digital (antiga grf)': 'GFD — Guia FGTS Digital',
    'relação fgts digital': 'Rel. FGTS Digital',
    'folha de pagamento (específico por serviço)': 'Folha de Pagamento',
    'folha de pagamento': 'Folha de Pagamento',
    'folha de ponto': 'Folha de Ponto',
    'aso - atestado de saúde ocupacional': 'ASO',
    'aso': 'ASO',
    'aso demissional': 'ASO Demissional',
    'nr 18 - treinamento admissional para trabalho em obra': 'NR 18 — Admissional',
    'nr 18 - treinamento admissional': 'NR 18 — Admissional',
    'nr 18 - treinamento admissional para trabalho em construção civil': 'NR 18 — Admissional',
    'nr 6 - treinamento sobre o uso adequado, guarda e conservação dos epis': 'NR 6 — EPI',
    'nr 6 - treinamento sobre epi': 'NR 6 — EPI',
    'nr 6 - treinamento sobre o uso adequado, guarda e conservação de epi': 'NR 6 — EPI',
    'nr 35 - trabalho em altura': 'NR 35 — Altura',
    'nr 10 - segurança em instalações elétricas': 'NR 10 — Eletricidade',
    'nr 20 - inflamáveis e combustíveis': 'NR 20 — Inflamáveis',
    'nr33 trabalho em espaços confinados 16 hrs trabalhador e 40 hrs supervisor': 'NR 33 — Espaços Confinados',
    'nr35 - treinamento trabalho em altura acesso por cordas': 'NR35 Altura — Cordas',
    'nr 18 - treinamento para operadores de plataforma elevatória tipo tesoura': 'NR18 — Op. Plataforma',
    'ordem de serviço de segurança': 'OS Segurança',
    'ficha de epi': 'Ficha de EPI',
    'campanha sst (específico por serviço).': 'Campanha SST',
    'dds - dialogo diário de segurança consolidado mensal (específico por serviço)': 'DDS — Diálogo Diário de Seg.',
    'rg - registro geral': 'RG',
    'rg': 'RG',
    'pgr - programa de gerenciamento de riscos': 'PGR',
    'pgr - programa de gerenciamento de riscos ou declaração de informação digital (para empresas dispensadas da elaboração de pgr) (específico por serviço)': 'PGR',
    'pcmso - programa de controle médico de saúde ocupacional': 'PCMSO',
    'pcmso (específico para o serviço)': 'PCMSO',
    'ltcat - laudo técnico das condições ambientais de trabalho': 'LTCAT',
    'ltcat (específico para o serviço)': 'LTCAT',
    'pca - programa de conservação auditiva': 'PCA',
    'ppr - programa de proteção respiratória (especifico por serviço)': 'PPR',
    'fispq - ficha de informações de segurança do produto químico': 'FISPQ — Ficha de Segurança',
    'insalubridade (por serviço)': 'Insalubridade (por serviço)',
    'periculosidade (por serviço)': 'Periculosidade (por serviço)',
    'constituição da cipa': 'Constituição da CIPA',
    'autorização de subcontratação': 'Autorização de Subcontratação',
    'termo de constituição do consórcio': 'Termo de Constituição do Consórcio',
    'convenção coletiva': 'Convenção Coletiva',
    'trct - termo de rescisão do contrato de trabalho': 'TRCT — Rescisão',
    'trct - cópia da rescisão assinada  (termo de rescisão de contrato de trabalho e termo de quitação de rescisão do contrato de trabalho) ': 'TRCT — Rescisão',
    'trct - cópia da rescisão assinada (termo de rescisão de contrato de trabalho e termo de quitação de rescisão do contrato de trabalho)': 'TRCT — Rescisão',
    'comprovante fgts multa (grrf)': 'Comp. FGTS Multa (GRRF)',
    'plano de trabalho em altura': 'Plano Trabalho em Altura',
    'plano de trabalho em altura (específico por serviço)': 'Plano Trabalho em Altura',
    'plano de resgate para trabalho em altura': 'Plano Resgate Altura',
    'plano de resgate em espaço confinado': 'Plano Resgate Esp. Confinado',
    'plano de resgate para trabalho em espaço confinado (específico por serviço)': 'Plano Resgate Esp. Confinado',
    'plano de resgate para trabalho em espaço confinado (específico por serviço) ': 'Plano Resgate Esp. Confinado',
    'plano de ventilação para espaços confinados': 'Plano Ventilação Confinados',
    'projeto de escoramento de vala': 'Proj. Escoramento de Vala',
    'projeto de escavação de vala': 'Proj. Escavação de Vala',
    'projeto de trabalho em altura - quando aplicável (localidades sabesp)': 'Proj. Trabalho em Altura',
    'projeto de linha de vida para trabalhos em altura, assinado por responsável técnico': 'Proj. Linha de Vida',
    'plano de rigging': 'Plano de Rigging',
    'plano de rigging (içamento de carga)': 'Plano de Rigging (Içamento)',
    'plano de segurança de escavações': 'Plano Seg. Escavações',
    'plano de demolição nr-18': 'Plano de Demolição NR-18',
    'laudo de instalações elétricas': 'Laudo Instal. Elétricas',
    'inventário de máquinas e equipamentos': 'Inventário de Máquinas',
    'relação de endereços de almoxarifados e depósitos da contratada': 'Rel. Endereços Almoxarifados',
    'relação de veículos utilizados no contrato da sabesp (emplacamento e foto)': 'Rel. Veículos Contrato',
    'relação dos alojamentos ou canteiros de obras e sua localização': 'Rel. Alojamentos',
    'relação de inspeção de alojamentos': 'Rel. Inspeção Alojamentos',
    'relação dos profissionais responsáveis pelas questões de sst da contratada e subcontratadas': 'Rel. Profissionais SST',
    'pae - plano de atendimento a emergências': 'PAE — Emergência',
    'treinamento de solda': 'Trein. Soldagem',
    'apólice de seguro garantia (construção/fornecimento de bens)': 'Apólice Seguro Garantia',
    'apólice de seguro garantia (construção/fornecimento/serviços) com cobertura adicional para ações trabalhistas e previdenciárias (específico por serviço)': 'Apólice Seguro Garantia',
    'apólice de seguro rcc': 'Apólice Seguro RCC',
    'apólice de seguro rcg': 'Apólice Seguro RCG',
    'cnd municipal - iptu e iss': 'CND Municipal',
    'cnd federal - creditos tributarios e à divida ativa da união': 'CND Federal',
    'apr - análise preliminar de risco': 'APR — Análise Preliminar de Risco',
    'laudo de conformidade das instalações elétricas do canteiro de obras': 'Laudo Conf. Instal. Elétricas',
    'evidências do programa de saúde mental (específico por serviço)': 'Evidências Prog. Saúde Mental',
    'programa de saúde mental (específico por serviço)': 'Programa de Saúde Mental',
    'nr13 - treinamento de segurança na operação de unidades de processos - vasos de pressão (40hrs)': 'NR 13 — Vasos de Pressão',
    'gps/darf - documento de arrecadação fiscal / dctfweb (relatórios gerais: recibo de entrega, resumo de débitos e resumo de créditos) e per/dcomp (quando houver), das-mei ou dae (documento de arrecadação do e-social)': 'GPS/DARF — DCTFweb (Doc. Arrecadação)',
}

DOC_COLORS = {
    'contrato': '#6366f1',
    'certidão': '#0891b2', 'cnd': '#0891b2', 'cndt': '#0891b2',
    'gps': '#7c3aed', 'darf': '#7c3aed', 'gfd': '#7c3aed', 'fgts': '#7c3aed',
    'folha': '#059669',
    'aso': '#dc2626',
    'nr ': '#d97706', 'nr1': '#d97706', 'nr3': '#d97706',
    'pgr': '#be123c', 'pcmso': '#be123c', 'ltcat': '#be123c', 'pca': '#be123c', 'ppr': '#be123c',
    'plano': '#2563a8', 'projeto': '#2563a8',
    'apólice': '#0f766e',
    'trct': '#9f1239',
    'rg': '#64748b',
    'relação': '#475569',
}

def shorten_doc(doc):
    if doc is None: return '—'
    key = doc.strip().lower()
    if key in DOC_SHORT: return DOC_SHORT[key]
    for k, v in DOC_SHORT.items():
        if key.startswith(k): return v
    return doc.strip()

def doc_color(doc):
    if doc is None: return '#64748b'
    key = doc.strip().lower()
    for prefix, color in DOC_COLORS.items():
        if key.startswith(prefix): return color
    return '#64748b'

def badge(st):
    if st == 'Vencido': return 'venc'
    if st in ('Inválido', 'Pendente - Isenção Reprovada', 'Reprovado pelo cliente'): return 'rep'
    if st in ('Em validação', 'Pendente - Isenção Aguardando Aprovação'): return 'val'
    if st in ('Pendente', 'Perto do Vencimento'): return 'pend'
    return 'ok'

def shorten_st(st):
    m = {
        'Pendente': 'Pendente',
        'Perto do Vencimento': 'Perto do vencimento',
        'Inválido': 'Reprovado',
        'Vencido': 'Vencido',
        'Reprovado pelo cliente': 'Reprovado pelo cliente',
        'Pendente - Isenção Reprovada': 'Isenção reprovada',
        'Pendente - Isenção Aguardando Aprovação': 'Isenção em análise',
        'Em validação': 'Em validação',
        'OK': 'Aprovado',
        'Aprovado pelo cliente': 'Aprovado pelo cliente',
        'Liberado com ressalvas no cliente': 'Aprovado com ressalva',
        'Isento': 'Isenção aprovada',
    }
    return m.get(st, st)

def parse_date(s):
    if not s: return None
    s = str(s).strip()
    for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%m/%d/%Y'):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None

def fmt_pct1(v):
    return f"{v:.1f}%".replace('.', ',')

def kpi_class(pct, target):
    return 'kpi-green' if pct >= target else 'kpi-yellow'

# ─── PROCESSAMENTO DOCUMENTAL ────────────────────────────────────────────────

def process_docs(rows, sub_names, proprio_label):

    def clean_name(raw):
        if raw is None or (isinstance(raw, str) and not raw.strip()):
            return proprio_label
        if raw in sub_names: return sub_names[raw]
        cleaned = re.sub(r'^\d{2}\.\d{3}\.\d{3}\s+', '', str(raw))
        cleaned = re.sub(r'^\d{8,11}\s+', '', cleaned)
        return cleaned.title().strip()

    def clean_func(f):
        if not f or str(f).strip() in ('', 'None'): return '—'
        return str(f).strip().title()

    total = len(rows)
    if total == 0:
        return None

    sc = collections.Counter(r.get('Status do Documento', '') for r in rows)
    nc = sum(v for k, v in sc.items() if k in NC_STATUSES)
    na = sc.get('Pendente', 0)
    conf_pct = (total - nc) / total * 100
    ader_pct = (total - na) / total * 100
    n_subs = len(set(r.get('Subcontratado') for r in rows if r.get('Subcontratado')))

    # Status rows (Seção 2)
    defs_with_counts = []
    for label, key, color, is_nc, is_na in STATUS_ROWS_DEFS:
        qty = sc.get(key, 0)
        defs_with_counts.append((label, qty, color, is_nc, is_na))
    defs_with_counts.sort(key=lambda x: -x[1])

    status_rows = []
    for label, qty, color, is_nc, is_na in defs_with_counts:
        pct = qty / total * 100 if total else 0
        status_rows.append({
            'label': label, 'qty': qty, 'color': color,
            'is_nc': is_nc, 'is_na': is_na, 'pct': round(pct, 1)
        })

    donut = {
        'labels': [x['label'] for x in status_rows],
        'data':   [x['qty']   for x in status_rows],
        'colors': [x['color'] for x in status_rows],
    }

    # NC por status (Seção 3)
    nc_status = []
    nc_by_st = []
    for st_key, st_label, st_color in NC_STATUS_DEFS:
        st_rows = [r for r in rows if r.get('Status do Documento') == st_key]
        if st_rows:
            nc_by_st.append({'label': st_label, 'color': st_color, 'rows': st_rows})
    nc_by_st.sort(key=lambda x: -len(x['rows']))
    max_nc = len(nc_by_st[0]['rows']) if nc_by_st else 1
    for i, item in enumerate(nc_by_st):
        entries = [{'empresa': clean_name(r.get('Subcontratado')),
                    'func': clean_func(r.get('Funcionário')),
                    'doc_s': shorten_doc(r.get('Documento'))} for r in item['rows']]
        nc_status.append({
            'id': f'ns{i+1}', 'label': item['label'], 'color': item['color'],
            'count': len(item['rows']),
            'pct': round(len(item['rows']) / max_nc * 100),
            'entries': entries
        })

    # Documentos com pendências (Seção 4)
    focus_rows = [r for r in rows if r.get('Status do Documento') in FOCUS]
    doc_focus = collections.defaultdict(list)
    for r in focus_rows:
        doc_focus[r.get('Documento')].append(r)
    all_docs = sorted(doc_focus.items(), key=lambda x: -len(x[1]))
    max_count = len(all_docs[0][1]) if all_docs else 1
    docs = []
    for i, (doc, drows) in enumerate(all_docs):
        cnt = len(drows)
        color = doc_color(doc)
        entries = [{'empresa': clean_name(r.get('Subcontratado')),
                    'func': clean_func(r.get('Funcionário')),
                    'badge': badge(r.get('Status do Documento', '')),
                    'st_s': shorten_st(r.get('Status do Documento', ''))} for r in drows]
        docs.append({'id': f'd{i+1}', 'name': shorten_doc(doc), 'count': cnt,
                     'color': color, 'pct': round(cnt / max_count * 100), 'entries': entries})

    # Ranking subcontratadas (Seção 5)
    sub_total_c = collections.defaultdict(int)
    sub_focus_c = collections.defaultdict(int)
    sub_crit_c  = collections.defaultdict(int)
    sub_entries = collections.defaultdict(list)
    for r in rows:
        emp = r.get('Subcontratado')
        st  = r.get('Status do Documento', '')
        sub_total_c[emp] += 1
        if st in FOCUS:
            sub_focus_c[emp] += 1
            sub_entries[emp].append({
                'func': clean_func(r.get('Funcionário')),
                'doc_s': shorten_doc(r.get('Documento')),
                'badge': badge(st), 'st_s': shorten_st(st)
            })
        if st in CRIT:
            sub_crit_c[emp] += 1
    sub_filtered = {e: sub_focus_c[e] for e in sub_total_c if sub_focus_c[e] > 0}
    sub_sorted = sorted(sub_filtered.items(), key=lambda x: -x[1])
    subs = []
    for i, (emp, focus) in enumerate(sub_sorted):
        tot  = sub_total_c[emp]
        taxa = round(focus / tot * 100) if tot else 0
        crit = sub_crit_c[emp]
        pill_color = '#dc2626' if taxa > 50 else ('#ca8a04' if taxa >= 10 else '#16a34a')
        subs.append({
            'id': f's{i+1}', 'nome': clean_name(emp),
            'focus': focus, 'total': tot, 'taxa': taxa, 'crit': crit,
            'pill_color': pill_color, 'entries': sub_entries[emp]
        })

    # Seção 7 — Documentos próximos do vencimento
    today = datetime.date.today()
    def alert_tag(dt):
        if dt is None: return ('cinza', 2)
        days = (dt - today).days
        if days <= 10: return ('vermelho', 0)
        if days <= 20: return ('amarelo', 1)
        return ('cinza', 2)

    venc_entries = []
    for r in rows:
        st_doc = r.get('Status do Documento')
        if st_doc not in ('Perto do Vencimento', 'Vencido'):
            continue
        dt = parse_date(r.get('Data de Vencimento'))
        emp = clean_name(r.get('Subcontratado'))
        func = clean_func(r.get('Funcionário'))
        dias_val = (dt - today).days if dt else None
        doc_s    = shorten_doc(r.get('Documento'))
        if st_doc == 'Vencido':
            tag = 'vencido'
        else:
            tag, _ = alert_tag(dt)
        # vencidos têm dias negativo → ordenam primeiro naturalmente
        venc_entries.append({
            'empresa':   emp,
            'func':      func,
            'doc':       doc_s,
            'data_venc': dt.strftime('%d/%m/%Y') if dt else '—',
            'dias':      dias_val,
            'tag':       tag,
            '_ord':      (dias_val if dias_val is not None else 999, emp, func, doc_s),
        })
    venc_entries.sort(key=lambda x: x['_ord'])
    vencimento_proximo = [{k: v for k, v in e.items() if k != '_ord'} for e in venc_entries]

    # Mapa funcionário → docs NC completo (inclui Isenção Ag. e Em Validação)
    worker_docs = collections.defaultdict(list)
    for r in rows:
        st = r.get('Status do Documento', '')
        if st in NC_STATUSES:
            func = clean_func(r.get('Funcionário'))
            if func and func != '—':
                worker_docs[func].append({
                    'doc':  shorten_doc(r.get('Documento')),
                    'st_s': shorten_st(st),
                })

    # Relatório completo de todos os documentos (aba Todos Documentos)
    ST_ORDER = {
        'Pendente': 0, 'Isenção reprovada': 1, 'Reprovado': 2,
        'Reprovado pelo cliente': 3, 'Vencido': 4, 'Perto do vencimento': 5,
        'Isenção em análise': 6, 'Em validação': 7,
        'Isenção aprovada': 8, 'Aprovado com ressalva': 9,
        'Aprovado pelo cliente': 10, 'Aprovado': 11,
    }
    all_docs = []
    for r in rows:
        st  = r.get('Status do Documento', '').strip()
        emp = clean_name(r.get('Subcontratado'))
        func = clean_func(r.get('Funcionário'))
        doc  = shorten_doc(r.get('Documento', ''))
        st_s = shorten_st(st)
        all_docs.append({
            'empresa': emp,
            'func':    func,
            'doc':     doc,
            'st_s':    st_s,
            'sev':     badge(st),
        })
    all_docs.sort(key=lambda x: (
        ST_ORDER.get(x['st_s'], 99), x['empresa'], x['func'], x['doc']
    ))

    return {
        'total': total, 'nc': nc, 'na': na,
        'conf_pct': round(conf_pct, 5),
        'ader_pct': round(ader_pct, 5),
        'n_subs': n_subs,
        'conf_fmt': fmt_pct1(conf_pct),
        'ader_fmt': fmt_pct1(ader_pct),
        'conf_class': kpi_class(conf_pct, CONF_TARGET),
        'ader_class': kpi_class(ader_pct, ADER_TARGET),
        'status_rows': status_rows,
        'donut': donut,
        'nc_status': nc_status,
        'docs': docs,
        'subs': subs,
        'worker_docs': dict(worker_docs),
        'vencimento_proximo': vencimento_proximo,
        'all_docs': all_docs,
    }

# ─── PROCESSAMENTO DE INTEGRAÇÃO ─────────────────────────────────────────────

def process_integration(rows_int, sub_names, proprio_label):
    """
    Deduplica por Funcionário (nome), melhor status vence.
    Retorna indicadores CP, CI e array cpci com trabalhadores não-conformes.
    """
    def clean_name(raw):
        if raw is None or (isinstance(raw, str) and not raw.strip()):
            return proprio_label
        if raw in sub_names: return sub_names[raw]
        cleaned = re.sub(r'^\d{2}\.\d{3}\.\d{3}\s+', '', str(raw))
        cleaned = re.sub(r'^\d{8,11}\s+', '', cleaned)
        return cleaned.title().strip()

    by_name = collections.defaultdict(list)
    for r in rows_int:
        nome = (r.get('Funcionário') or '').strip()
        if nome:
            by_name[nome].append(r)

    def best_row(rlist):
        return min(rlist, key=lambda r: INT_PRIORITY.get(r.get('Status', ''), 99))

    workers = {nome: best_row(rlist) for nome, rlist in by_name.items()}
    total_w = len(workers)
    if total_w == 0:
        return None

    cp_ok = sum(1 for r in workers.values() if r.get('Status') == 'Acesso Liberado')
    ci_ok = sum(1 for r in workers.values() if r.get('Status') in CI_OK)
    ci_acesso   = cp_ok
    ci_pend_int = sum(1 for r in workers.values() if r.get('Status') == 'Pendências porém Integrado')
    ci_falta    = sum(1 for r in workers.values() if r.get('Status') == 'Falta Integração')

    cp_pct = cp_ok / total_w * 100
    ci_pct = ci_ok / total_w * 100

    # Trabalhadores com status != Acesso Liberado para Seção 6
    cpci = []
    for nome, r in sorted(workers.items()):
        st = r.get('Status', '')
        if st != 'Acesso Liberado':
            forn_raw = r.get('Fornecedor') or r.get('Subcontratado') or ''
            cpci.append({
                'fornecedor': clean_name(forn_raw.strip() if forn_raw else None),
                'funcionario': nome.title(),
                'status': st,
            })
    cpci.sort(key=lambda x: (x['fornecedor'], x['funcionario']))

    return {
        'total_workers': total_w,
        'cp_ok': cp_ok,
        'cp_nc': total_w - cp_ok,
        'cp_pct': round(cp_pct, 5),
        'cp_fmt': fmt_pct1(cp_pct),
        'cp_class': kpi_class(cp_pct, CP_TARGET),
        'ci_ok': ci_ok,
        'ci_nc': total_w - ci_ok,
        'ci_pct': round(ci_pct, 5),
        'ci_fmt': fmt_pct1(ci_pct),
        'ci_class': kpi_class(ci_pct, CI_TARGET),
        'ci_acesso': ci_acesso,
        'ci_pend_int': ci_pend_int,
        'ci_falta': ci_falta,
        'cpci': cpci,
    }

# ─── DETALHE CP/CI POR FUNCIONÁRIO ──────────────────────────────────────────

def process_cpci_detail(rows_doc, rows_int, sub_names, proprio_label):
    """
    Workers vêm de relatorio_de_documentos (CPF único).
    Status de integração vem de status_integracao (join por CPF).
    Workers sem par no CSV de integração recebem status_int = 'N.D.'.
    """

    def clean_emp(raw):
        if not raw or not str(raw).strip():
            return proprio_label
        raw = str(raw).strip()
        if raw in sub_names: return sub_names[raw]
        cleaned = re.sub(r'^\d{2}\.\d{3}\.\d{3}\s+', '', raw)
        cleaned = re.sub(r'^\d{8,11}\s+', '', cleaned)
        return cleaned.strip() or proprio_label

    def doc_sev(st):
        if st in ('Inválido', 'Pendente - Isenção Reprovada', 'Vencido', 'Reprovado pelo cliente'):
            return 'rep'
        if st in ('Em validação', 'Pendente - Isenção Aguardando Aprovação'): return 'val'
        return 'pen'

    ALOC_DISPLAY_MAP = {
        'Funcionários Alocados':        'Alocado',
        'Funcionários que já participaram': 'Já participou',
        'Demitidos':                    'Demitido',
    }
    ALOC_SORT_MAP = {
        'Funcionários Alocados':        0,
        'Funcionários que já participaram': 1,
        'Demitidos':                    2,
    }
    INT_SORT_D = {
        'Pendências e Falta Integração': 0,
        'Falta Integração':              1,
        'Pendências porém Integrado':    2,
        'Acesso Liberado':               3,
        'N.D.':                          4,
    }

    # 1. Mapa CPF → melhor status de integração (de status_integracao)
    by_cpf_int = {}
    for r in rows_int:
        cpf = (r.get('CPF') or '').strip()
        if not cpf: continue
        st  = r.get('Status', '')
        pri = INT_PRIORITY.get(st, 99)
        if cpf not in by_cpf_int:
            by_cpf_int[cpf] = {'status_int': st, 'pri': pri}
        else:
            if pri < by_cpf_int[cpf]['pri']:
                by_cpf_int[cpf] = {'status_int': st, 'pri': pri}

    # 2. Documentos pessoais (CPF preenchido) e de empresa (CPF vazio) com NC
    p_docs = collections.defaultdict(list)   # CPF  → [{doc, st_s, sev}]
    c_docs = collections.defaultdict(list)   # CNPJ → [{doc, st_s, sev}]
    for r in rows_doc:
        st = r.get('Status do Documento', '')
        if st not in NC_STATUSES: continue
        cpf       = (r.get('CPF') or '').strip()
        sub_cnpj  = (r.get('Subcontratado CNPJ') or '').strip()
        main_cnpj = (r.get('CNPJ') or '').strip()
        entry = {'doc': shorten_doc(r.get('Documento')),
                 'st_s': shorten_st(st), 'sev': doc_sev(st)}
        if cpf:
            p_docs[cpf].append(entry)
        else:
            c_docs[sub_cnpj or main_cnpj].append(entry)

    # 3. Workers únicos de relatorio_de_documentos (somente linhas com CPF)
    #    Para alocação: prioridade Alocado > Participou > Demitido
    by_cpf_doc = {}
    for r in rows_doc:
        cpf  = (r.get('CPF') or '').strip()
        nome = (r.get('Funcionário') or '').strip()
        if not cpf or not nome: continue
        sub_raw  = (r.get('Subcontratado') or '').strip()
        sub_cnpj = (r.get('Subcontratado CNPJ') or '').strip()
        cnpj     = sub_cnpj or (r.get('CNPJ') or '').strip()
        cargo    = str(r.get('Cargo') or '').strip().title()
        regime   = str(r.get('Regime de Contratação') or '').strip()
        aloc_raw = str(r.get('Status do Funcionário') or '').strip()
        aloc_pri = ALOC_SORT_MAP.get(aloc_raw, 9)
        empresa  = clean_emp(sub_raw)
        if cpf not in by_cpf_doc:
            by_cpf_doc[cpf] = {
                'nome': nome.title(), 'empresa': empresa, 'cnpj': cnpj,
                'cargo': cargo, 'regime': regime,
                'alocacao_raw': aloc_raw, 'aloc_pri': aloc_pri,
            }
        else:
            w = by_cpf_doc[cpf]
            # Alocação: vence o melhor status
            if aloc_pri < w['aloc_pri']:
                w['alocacao_raw'] = aloc_raw
                w['aloc_pri'] = aloc_pri
            # Preenche campos faltantes
            if not w['cargo'] and cargo:   w['cargo'] = cargo
            if not w['regime'] and regime: w['regime'] = regime

    if not by_cpf_doc:
        return None

    def dedup(lst):
        seen, out = set(), []
        for d in lst:
            k = (d['doc'], d['st_s'])
            if k not in seen:
                seen.add(k); out.append(d)
        return out

    # 4. Montar detail cruzando docs + integração
    detail = []
    for cpf, w in by_cpf_doc.items():
        int_info   = by_cpf_int.get(cpf)
        status_int = int_info['status_int'] if int_info else 'N.D.'
        p = dedup(p_docs.get(cpf, []))
        c = dedup(c_docs.get(w['cnpj'], []))
        sevs    = [d['sev'] for d in p + c]
        row_sev = 'rep' if 'rep' in sevs else ('val' if 'val' in sevs else ('pen' if sevs else 'ok'))
        aloc_raw = w['alocacao_raw']
        detail.append({
            'nome':       w['nome'],
            'empresa':    w['empresa'],
            'cargo':      w['cargo'],
            'regime':     w['regime'],
            'status_int': status_int,
            'alocacao':   ALOC_DISPLAY_MAP.get(aloc_raw, aloc_raw or '—'),
            'alocacao_raw': aloc_raw,
            'n_pess': len(p), 'n_emp': len(c),
            'docs_pessoais': p, 'docs_empresa': c, 'sev': row_sev,
        })

    # 5. Ordenação: pior integração → melhor; pior docs → melhor; alocação; nome
    def sort_key(w):
        return (INT_SORT_D.get(w['status_int'], 4),
                0 if w['n_pess'] > 0 else 1, -w['n_pess'],
                0 if w['n_emp']  > 0 else 1, -w['n_emp'],
                ALOC_SORT_MAP.get(w['alocacao_raw'], 9),
                w['nome'])
    detail.sort(key=sort_key)

    total  = len(detail)
    n_lib  = sum(1 for w in detail if w['status_int'] == 'Acesso Liberado')
    n_pess = sum(1 for w in detail if w['n_pess'] > 0)
    n_emp  = sum(1 for w in detail if w['n_emp']  > 0)
    n_nd   = sum(1 for w in detail if w['status_int'] == 'N.D.')
    n_alocados     = sum(1 for w in detail if w['alocacao_raw'] == 'Funcionários Alocados')
    n_participaram = sum(1 for w in detail if w['alocacao_raw'] == 'Funcionários que já participaram')
    n_pend_falta   = sum(1 for w in detail if w['status_int'] == 'Pendências e Falta Integração')
    n_falta_int    = sum(1 for w in detail if w['status_int'] == 'Falta Integração')
    n_pend_integr  = sum(1 for w in detail if w['status_int'] == 'Pendências porém Integrado')

    # Ativos = trabalhadores presentes no CSV de integração (status_int != 'N.D.')
    n_ativos            = sum(1 for w in detail if w['status_int'] != 'N.D.')
    n_ativos_doc_ok     = sum(1 for w in detail if w['status_int'] != 'N.D.' and w['n_pess'] == 0)
    n_ativos_doc_pend   = sum(1 for w in detail if w['status_int'] != 'N.D.' and w['n_pess'] > 0)
    n_ativos_acesso_lib = sum(1 for w in detail if w['status_int'] == 'Acesso Liberado')
    n_ativos_int_pend   = n_ativos - n_ativos_acesso_lib
    n_ativos_pend_emp   = sum(1 for w in detail if w['status_int'] != 'N.D.' and w['n_emp'] > 0)

    cnt = collections.Counter()
    for w in detail:
        ps = 'Com pendências' if w['n_pess'] > 0 else 'Sem pendências'
        es = 'Com pendências' if w['n_emp']  > 0 else 'Sem pendências'
        cnt[(w['status_int'], ps, es)] += 1

    combos = sorted(
        [{'si': k[0], 'sp': k[1], 'se': k[2], 'qty': v} for k, v in cnt.items()],
        key=lambda x: (INT_SORT_D.get(x['si'], 4),
                       0 if x['sp'] == 'Com pendências' else 1,
                       0 if x['se'] == 'Com pendências' else 1)
    )

    return {'total': total, 'n_liberado': n_lib,
            'n_pend_pess': n_pess, 'n_pend_emp': n_emp,
            'n_nd': n_nd,
            'n_alocados': n_alocados, 'n_participaram': n_participaram,
            'n_pend_falta_int': n_pend_falta, 'n_falta_int': n_falta_int,
            'n_pend_integrado': n_pend_integr,
            'n_ativos': n_ativos,
            'n_ativos_doc_ok': n_ativos_doc_ok,
            'n_ativos_doc_pend': n_ativos_doc_pend,
            'n_ativos_acesso_lib': n_ativos_acesso_lib,
            'n_ativos_int_pend': n_ativos_int_pend,
            'n_ativos_pend_emp': n_ativos_pend_emp,
            'combos': combos, 'workers': detail}
# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    args = parse_args()
    projeto  = args.projeto
    csv_doc  = args.csv_doc
    csv_int  = args.csv_int

    config        = load_config(projeto)
    sub_names     = config.get('sub_names', {})
    proprio_label = config.get('proprio_label', f'Consórcio {projeto.upper()} (próprio)')

    # Data de extração a partir do nome do arquivo
    date_match = re.search(r'(\d{4}-\d{2}-\d{2})T', os.path.basename(csv_doc))
    extraction_date = date_match.group(1) if date_match else datetime.date.today().isoformat()

    print(f'[etl] Projeto: {projeto}  Data: {extraction_date}')
    print(f'[etl] CSV docs: {csv_doc}')

    # Leitura dos dados
    rows_doc = read_tabular(csv_doc)
    print(f'[etl] Linhas docs: {len(rows_doc)}')

    doc_result = process_docs(rows_doc, sub_names, proprio_label)
    if doc_result is None:
        sys.exit('[etl] CSV de documentos vazio ou sem colunas reconhecidas.')

    int_result   = None
    cpci_detail  = None
    rows_int     = []
    if csv_int:
        print(f'[etl] CSV integração: {csv_int}')
        rows_int = read_tabular(csv_int)
        print(f'[etl] Linhas integração (raw): {len(rows_int)}')
        int_result = process_integration(rows_int, sub_names, proprio_label)
        if int_result:
            print(f'[etl] Trabalhadores únicos: {int_result["total_workers"]}  CP={int_result["cp_fmt"]}  CI={int_result["ci_fmt"]}')
        cpci_detail = process_cpci_detail(rows_doc, rows_int, sub_names, proprio_label)
        if cpci_detail:
            print(f'[etl] cpci_detail: {cpci_detail["total"]} funcionários')
            # CP: recalcular sobre ativos (CSV integração como denominador autoritário)
            # CI: manter do process_integration (CSV integração é fonte autoritária)
            if int_result:
                nat = int_result.get('total_workers', 0)  # ativos = presentes no CSV integração
                if nat > 0:
                    cp_ok      = cpci_detail['n_ativos_doc_ok']
                    new_cp_pct = cp_ok / nat * 100
                    int_result['cp_ok']    = cp_ok
                    int_result['cp_pct']   = round(new_cp_pct, 5)
                    int_result['cp_fmt']   = fmt_pct1(new_cp_pct)
                    int_result['cp_class'] = kpi_class(new_cp_pct, CP_TARGET)
                    # CI permanece inalterado (process_integration já computou corretamente)
                # Sincronizar cpci_detail com contagens autoritárias do CSV integração
                cpci_detail['n_ativos']            = nat
                cpci_detail['n_ativos_acesso_lib'] = int_result.get('ci_ok', 0)
                cpci_detail['n_ativos_int_pend']   = nat - int_result.get('ci_ok', 0)

    # Histórico
    historico = load_historico(projeto)
    historico = [h for h in historico if h['data'] != extraction_date]  # idempotência
    previous  = historico[-1] if historico else None

    def pct_delta(curr, prev_val):
        if prev_val is None or prev_val == 0: return None
        return round((curr - prev_val) / prev_val * 100, 3)

    variacao = {
        'data_anterior': previous['data'] if previous else None,
        'd_conf':  pct_delta(doc_result['conf_pct'], previous['conf_pct'] if previous else None),
        'd_ader':  pct_delta(doc_result['ader_pct'], previous['ader_pct'] if previous else None),
        'd_total': pct_delta(doc_result['total'],    previous['total']    if previous else None),
        'd_subs':  pct_delta(doc_result['n_subs'],   previous['n_subs']   if previous else None),
        'd_cp': pct_delta(int_result['cp_pct'], previous.get('cp_pct') if previous else None) if int_result and previous and previous.get('cp_pct') else None,
        'd_ci': pct_delta(int_result['ci_pct'], previous.get('ci_pct') if previous else None) if int_result and previous and previous.get('ci_pct') else None,
    }

    # Meta do JSON
    meta = {
        'projeto':    projeto,
        'data':       extraction_date,
        'total':      doc_result['total'],
        'nc':         doc_result['nc'],
        'na':         doc_result['na'],
        'conf_pct':   doc_result['conf_pct'],
        'ader_pct':   doc_result['ader_pct'],
        'n_subs':     doc_result['n_subs'],
        'conf_fmt':   doc_result['conf_fmt'],
        'ader_fmt':   doc_result['ader_fmt'],
        'conf_class': doc_result['conf_class'],
        'ader_class': doc_result['ader_class'],
    }
    if int_result:
        meta.update({
            'total_workers': int_result['total_workers'],
            'cp_ok':    int_result['cp_ok'],
            'cp_nc':    int_result['cp_nc'],
            'cp_pct':   int_result['cp_pct'],
            'cp_fmt':   int_result['cp_fmt'],
            'cp_class': int_result['cp_class'],
            'ci_ok':    int_result['ci_ok'],
            'ci_nc':    int_result['ci_nc'],
            'ci_pct':   int_result['ci_pct'],
            'ci_fmt':   int_result['ci_fmt'],
            'ci_class': int_result['ci_class'],
            'ci_acesso':    int_result['ci_acesso'],
            'ci_pend_int':  int_result['ci_pend_int'],
            'ci_falta':     int_result['ci_falta'],
        })
    if cpci_detail:
        meta['n_ativos'] = cpci_detail.get('n_ativos', 0)
    else:
        meta.update({
            'total_workers': None, 'cp_ok': None, 'cp_nc': None,
            'cp_pct': None, 'cp_fmt': 'N/D', 'cp_class': 'kpi-nd',
            'ci_ok': None, 'ci_nc': None,
            'ci_pct': None, 'ci_fmt': 'N/D', 'ci_class': 'kpi-nd',
            'ci_acesso': None, 'ci_pend_int': None, 'ci_falta': None,
        })

    # JSON completo da extração
    output = {
        'meta':        meta,
        'variacao':    variacao,
        'status_rows': doc_result['status_rows'],
        'donut':       doc_result['donut'],
        'nc_status':   doc_result['nc_status'],
        'docs':        doc_result['docs'],
        'subs':        doc_result['subs'],
        'worker_docs':        doc_result['worker_docs'],
        'vencimento_proximo': doc_result['vencimento_proximo'],
        'all_docs':           doc_result['all_docs'],
        'cpci':               int_result['cpci'] if int_result else [],
        'cpci_detail':        cpci_detail,
    }

    out_path = os.path.join('data', projeto, f'{extraction_date}.json')
    save_json(out_path, output)
    print(f'[etl] Gravado: {out_path}')

    # Atualiza historico.json
    hist_entry = {
        'data':      extraction_date,
        'conf_pct':  doc_result['conf_pct'],
        'ader_pct':  doc_result['ader_pct'],
        'total':     doc_result['total'],
        'n_subs':    doc_result['n_subs'],
        'cp_pct':    int_result['cp_pct'] if int_result else None,
        'ci_pct':    int_result['ci_pct'] if int_result else None,
    }
    historico.append(hist_entry)
    historico.sort(key=lambda h: h['data'])
    hist_path = os.path.join('data', projeto, 'historico.json')
    save_json(hist_path, historico)
    print(f'[etl] Hist\u00f3rico atualizado: {hist_path}  ({len(historico)} entradas)')

    # Atualiza manifest.json
    datas = sorted(h['data'] for h in historico)
    manifest = {'datas': datas, 'ultima': datas[-1]}
    save_json(os.path.join('data', projeto, 'manifest.json'), manifest)
    print(f'[etl] Manifest atualizado. Última: {manifest["ultima"]}')
    print('[etl] Concluído.')


if __name__ == '__main__':
    main()
