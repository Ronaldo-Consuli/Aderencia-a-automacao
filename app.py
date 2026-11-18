#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Painel de Aderência Automação para Streamlit.

Execute: streamlit run app.py

Regras:
- Monitor de Digitação é a base dos totais, separados por Grupo.
- Pareamento 1:1 por Grupo/Área, material, equipamento, local, tipo, característica e data/hora.
- Divergência do valor não impede a correspondência.
- Exclui Acidez apenas da Destilação, dos dois lados.
- Registros sem Autolab não são prova de digitação manual.
- A data é identificada automaticamente nos arquivos importados.

Usa somente streamlit e a biblioteca-padrão do Python.
"""
from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from html import escape as html_escape
from io import BytesIO, StringIO
import csv
import re
import unicodedata
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile

XLS_NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
LIMITE_UPLOAD = 25 * 1024 * 1024


# ========================= Motor de conciliação =========================

def _indice_coluna(referencia: str) -> int:
    resultado = re.match(r'[A-Za-z]+', referencia or '')
    if not resultado:
        raise ValueError('Referência de célula inválida no Excel.')
    indice = 0
    for letra in resultado.group().upper():
        indice = indice * 26 + ord(letra) - 64
    return indice - 1


def ler_xlsx(dados: bytes) -> list[dict]:
    """Lê a primeira aba do XLSX exportado pelo SAP, sem bibliotecas externas."""
    try:
        with ZipFile(BytesIO(dados), 'r') as pasta:
            nomes = pasta.namelist()
            planilhas = [s for s in nomes if re.fullmatch(r'xl/worksheets/sheet\d+\.xml', s)]
            planilhas.sort(key=lambda s: int(re.search(r'sheet(\d+)', s).group(1)))
            if not planilhas:
                raise ValueError('O arquivo Excel não tem uma planilha legível.')

            compartilhadas = []
            if 'xl/sharedStrings.xml' in nomes:
                raiz = ET.fromstring(pasta.read('xl/sharedStrings.xml'))
                compartilhadas = [
                    ''.join(no.text or '' for no in item.iter(f'{XLS_NS}t'))
                    for item in raiz.findall(f'{XLS_NS}si')
                ]

            raiz = ET.fromstring(pasta.read(planilhas[0]))
            corpo = raiz.find(f'{XLS_NS}sheetData')
            if corpo is None:
                raise ValueError('A planilha não contém dados.')

            linhas = []
            for numero, linha in enumerate(corpo.findall(f'{XLS_NS}row'), start=1):
                valores = {}
                for cel in linha.findall(f'{XLS_NS}c'):
                    coluna = _indice_coluna(cel.get('r', 'A1'))
                    tipo = cel.get('t')
                    val = cel.find(f'{XLS_NS}v')
                    if tipo == 'inlineStr':
                        no = cel.find(f'{XLS_NS}is')
                        resultado = ''.join(t.text or '' for t in no.iter(f'{XLS_NS}t')) if no is not None else ''
                    elif val is None or val.text is None:
                        resultado = ''
                    elif tipo == 's':
                        resultado = compartilhadas[int(val.text)]
                    elif tipo == 'b':
                        resultado = val.text == '1'
                    elif tipo in ('str', 'e'):
                        resultado = val.text
                    else:
                        try:
                            resultado = Decimal(val.text)
                        except InvalidOperation:
                            resultado = val.text
                    valores[coluna] = resultado
                if valores:
                    linhas.append((int(linha.get('r', numero)), valores))

            if not linhas:
                raise ValueError('A planilha está vazia.')

            cabecalho = {col: str(v).strip() for col, v in linhas[0][1].items()}
            saida = []
            for linha_excel, valores in linhas[1:]:
                if not any(str(v).strip() for v in valores.values()):
                    continue
                registro = {nome: valores.get(col, '') for col, nome in cabecalho.items()}
                registro['_linha'] = linha_excel
                saida.append(registro)
            return saida
    except BadZipFile as erro:
        raise ValueError('Arquivo inválido: selecione um Excel no formato .xlsx.') from erro


def normalizar(valor, tirar_codigo=False) -> str:
    texto = str(valor if valor is not None else '').strip()
    if tirar_codigo:
        texto = re.sub(r'^\(\s*[^)]+\s*\)\s*[-–—]\s*', '', texto)
    texto = unicodedata.normalize('NFKD', texto).casefold()
    texto = ''.join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r'\s+', ' ', texto).strip()


def caracteristica_autolab(valor) -> str:
    return re.sub(r'^\(\s*[^)]+\s*\)\s*[-–—]\s*', '', str(valor or '')).strip()


def numero(valor):
    if valor is None or str(valor).strip() == '':
        return None
    if isinstance(valor, Decimal):
        return valor
    if isinstance(valor, (int, float)):
        return Decimal(str(valor))
    s = str(valor).replace('\xa0', '').replace(' ', '').strip()
    if '.' in s and ',' in s:
        s = s.replace('.', '').replace(',', '.') if s.rfind(',') > s.rfind('.') else s.replace(',', '')
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def dt_resultado(valor):
    if isinstance(valor, datetime):
        return valor.replace(microsecond=0)
    if isinstance(valor, (int, float, Decimal)):
        try:
            return (datetime(1899, 12, 30) + timedelta(days=float(valor))).replace(microsecond=0)
        except (OverflowError, ValueError):
            return None
    s = str(valor or '').strip()
    for formato in (
        '%d/%m/%Y %H:%M:%S', '%d/%m/%Y %H:%M',
        '%d/%m/%y %H:%M:%S', '%d/%m/%y %H:%M',
        '%d/%m/%Y', '%d/%m/%y',
        '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d',
    ):
        try:
            return datetime.strptime(s, formato)
        except ValueError:
            pass
    return None


def texto_valor(valor) -> str:
    if valor is None or str(valor).strip() == '':
        return ''
    n = numero(valor)
    if n is None:
        return str(valor).strip()
    txt = format(n, 'f')
    return txt.rstrip('0').rstrip('.') if '.' in txt else txt


def valores_iguais(a, b) -> bool:
    aa, bb = numero(a), numero(b)
    if aa is not None and bb is not None:
        return aa == bb
    return normalizar(a) == normalizar(b)


def chave_monitor(r):
    dt = dt_resultado(r.get('Dt. Resultado'))
    if dt is None:
        return None
    return (
        normalizar(r.get('Grupo')),
        normalizar(r.get('Material')), normalizar(r.get('Equipamento')),
        normalizar(r.get('Local')), normalizar(r.get('Tipo de Análise')),
        normalizar(r.get('Característica')), dt,
    )


def chave_autolab(r):
    dt = dt_resultado(r.get('Data Resultado'))
    if dt is None:
        return None
    return (
        normalizar(r.get('Área')),
        normalizar(r.get('Material')), normalizar(r.get('Equipamento')),
        normalizar(r.get('Local')), normalizar(r.get('Tipo')),
        normalizar(r.get('Característica'), tirar_codigo=True), dt,
    )


def validar_cabecalho(registros, campos, nome):
    if not registros:
        raise ValueError(f'O arquivo {nome} não possui registros.')
    faltantes = [campo for campo in campos if campo not in registros[0]]
    if faltantes:
        raise ValueError(f'Arquivo {nome}: colunas ausentes: {", ".join(faltantes)}.')


def excluir_acidez_destilacao(registro, campo_grupo, campo_caracteristica, auto=False):
    return (normalizar(registro.get(campo_grupo)) == 'destilacao'
            and normalizar(registro.get(campo_caracteristica), tirar_codigo=auto) == 'acidez')


def detalhe_monitor(r, tipo='sem'):
    dt = dt_resultado(r.get('Dt. Resultado'))
    return {
        'tipo': tipo,
        'grupo': str(r.get('Grupo', '')).strip(),
        'material': str(r.get('Material', '')).strip(),
        'equipamento': str(r.get('Equipamento', '')).strip(),
        'local': str(r.get('Local', '')).strip(),
        'caracteristica': str(r.get('Característica', '')).strip(),
        'hora': dt.strftime('%d/%m/%Y %H:%M') if dt else str(r.get('Dt. Resultado', '')),
        'monitor': texto_valor(r.get('Vl. Resultado')),
        'autolab': '',
        'status': '',
        'linha': r['_linha'],
    }


def detalhe_autolab(r):
    dt = dt_resultado(r.get('Data Resultado'))
    return {
        'tipo': 'auto_sem',
        'grupo': str(r.get('Área', '')).strip(),
        'material': str(r.get('Material', '')).strip(),
        'equipamento': str(r.get('Equipamento', '')).strip(),
        'local': str(r.get('Local', '')).strip(),
        'caracteristica': caracteristica_autolab(r.get('Característica')),
        'hora': dt.strftime('%d/%m/%Y %H:%M') if dt else str(r.get('Data Resultado', '')),
        'monitor': '',
        'autolab': texto_valor(r.get('Resultado')),
        'status': str(r.get('Status', '')).strip(),
        'linha': r['_linha'],
    }


def datas_em_comum(autolab_bytes: bytes, monitor_bytes: bytes) -> list[date]:
    """Datas de resultado presentes em ambos os arquivos, da mais recente à mais antiga."""
    auto = ler_xlsx(autolab_bytes)
    mon = ler_xlsx(monitor_bytes)
    validar_cabecalho(auto, ['Data Resultado'], 'Autolab')
    validar_cabecalho(mon, ['Dt. Resultado'], 'Monitor')
    dias_auto = {dt.date() for row in auto
                 if (dt := dt_resultado(row.get('Data Resultado'))) is not None}
    dias_mon = {dt.date() for row in mon
                if (dt := dt_resultado(row.get('Dt. Resultado'))) is not None}
    dias = sorted(dias_auto & dias_mon, reverse=True)
    if not dias:
        raise ValueError('Os arquivos não possuem datas de resultados em comum.')
    return dias


def analisar(autolab_bytes: bytes, monitor_bytes: bytes, dia: date) -> dict:
    auto_todos = ler_xlsx(autolab_bytes)
    moni_todos = ler_xlsx(monitor_bytes)
    validar_cabecalho(auto_todos,
                     ['Área', 'Material', 'Equipamento', 'Local', 'Tipo',
                      'Característica', 'Data Resultado', 'Resultado', 'Status'], 'Autolab')
    validar_cabecalho(moni_todos,
                     ['Grupo', 'Material', 'Equipamento', 'Local', 'Tipo de Análise',
                      'Característica', 'Dt. Resultado', 'Vl. Resultado'], 'Monitor')

    monitor_dia = [r for r in moni_todos if dt_resultado(r.get('Dt. Resultado'))
                   and dt_resultado(r['Dt. Resultado']).date() == dia]
    auto_dia = [r for r in auto_todos if dt_resultado(r.get('Data Resultado'))
                and dt_resultado(r['Data Resultado']).date() == dia]
    excluidos = sum(excluir_acidez_destilacao(r, 'Grupo', 'Característica') for r in monitor_dia)
    monitor = [r for r in monitor_dia
               if not excluir_acidez_destilacao(r, 'Grupo', 'Característica')]
    auto = [r for r in auto_dia
            if not excluir_acidez_destilacao(r, 'Área', 'Característica', auto=True)]

    por_chave_mon = defaultdict(list)
    por_chave_auto = defaultdict(list)
    for idx, row in enumerate(monitor):
        k = chave_monitor(row)
        if k is not None:
            por_chave_mon[k].append(idx)
    for idx, row in enumerate(auto):
        k = chave_autolab(row)
        if k is not None:
            por_chave_auto[k].append(idx)

    # Pareamento 1:1: primeiro maximiza pares com valores iguais para classificar
    # corretamente divergências em chaves repetidas; depois pareia as sobras.
    pares = []
    usados_mon, usados_auto = set(), set()
    for chave, aut_indices in por_chave_auto.items():
        mon_indices = por_chave_mon.get(chave, [])
        if not mon_indices:
            continue
        disponiveis = set(mon_indices)
        for ia in aut_indices:
            im = next((i for i in mon_indices if i in disponiveis
                       and valores_iguais(auto[ia].get('Resultado'), monitor[i].get('Vl. Resultado'))), None)
            if im is not None:
                pares.append((ia, im))
                disponiveis.remove(im)
                usados_mon.add(im)
                usados_auto.add(ia)
        for ia in aut_indices:
            if ia in usados_auto or not disponiveis:
                continue
            im = next(i for i in mon_indices if i in disponiveis)
            pares.append((ia, im))
            disponiveis.remove(im)
            usados_mon.add(im)
            usados_auto.add(ia)

    dados_grupos = {}
    for row in monitor:
        grupo = str(row.get('Grupo', '')).strip() or 'Sem grupo'
        info = dados_grupos.setdefault(grupo, {'grupo': grupo, 'monitor': 0, 'autolab': 0})
        info['monitor'] += 1
    for ia, im in pares:
        grupo = str(monitor[im].get('Grupo', '')).strip() or 'Sem grupo'
        dados_grupos[grupo]['autolab'] += 1

    grupos = []
    for nome, g in sorted(dados_grupos.items(), key=lambda item: normalizar(item[0])):
        g = dict(g)
        g['diferenca'] = g['monitor'] - g['autolab']
        g['aderencia'] = round(100 * g['autolab'] / g['monitor'], 4) if g['monitor'] else 0
        grupos.append(g)

    detalhes = [detalhe_monitor(r) for i, r in enumerate(monitor) if i not in usados_mon]
    divergencias = []
    for ia, im in pares:
        if not valores_iguais(auto[ia].get('Resultado'), monitor[im].get('Vl. Resultado')):
            d = detalhe_monitor(monitor[im], 'divergente')
            d['autolab'] = texto_valor(auto[ia].get('Resultado'))
            d['status'] = str(auto[ia].get('Status', '')).strip()
            divergencias.append(d)
    auto_sem = [detalhe_autolab(r) for i, r in enumerate(auto) if i not in usados_auto]
    detalhes.extend(divergencias)
    detalhes.extend(auto_sem)
    detalhes.sort(key=lambda d: (normalizar(d['grupo']), d['hora'],
                                  normalizar(d['material']), normalizar(d['caracteristica'])))

    total_mon = len(monitor)
    total_auto = len(usados_mon)
    return {
        'dia': dia.isoformat(),
        'grupos': grupos,
        'monitor': total_mon,
        'autolab': total_auto,
        'diferenca': total_mon - total_auto,
        'aderencia': round(100 * total_auto / total_mon, 4) if total_mon else 0,
        'divergencias': len(divergencias),
        'autolab_sem_monitor': len(auto_sem),
        'acidez_excluida': excluidos,
        'autolab_linhas_dia': len(auto_dia),
        'detalhes': detalhes,
    }

# ========================= Interface Streamlit =========================

CSS = """
<style>
:root{--verde:#00983A;--azul:#004881;--navy:#143257;--vermelho:#C0392B;
 --muted:#6B7785;--fundo:#E9EDF1;--borda:#E0E5EA;--texto:#2B3542}
.stApp{background:var(--fundo);color:var(--texto)}
.block-container{max-width:1320px;padding-top:1.35rem;padding-bottom:2rem}
#MainMenu,footer{visibility:hidden}
[data-testid="stSidebar"]{background:#F5F7FA;border-right:1px solid var(--borda)}
.hero{background:linear-gradient(118deg,#00983A 0%,#004881 64%,#083050 100%);
 border-radius:13px;padding:25px 29px;margin-bottom:16px;color:#fff}
.hero h1{font-size:28px;line-height:1.2;color:#fff;margin:0;font-weight:780;letter-spacing:-.4px}
.hero p{font-size:14px;color:#eaf3f7;margin:7px 0 0}
.subheading{color:#143257;font-size:19px;font-weight:730;margin:22px 0 13px}
.kpi-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:13px;margin:15px 0 4px}
.kpi-card{background:#fff;border:1px solid #E0E5EA;border-radius:10px;padding:17px 15px;
 box-shadow:0 1px 2px rgba(12,34,64,.06);text-align:center}
.kpi-label{color:#6B7785;font-size:12px;font-weight:600}
.kpi-num{color:#143257;font-size:29px;line-height:1.3;font-weight:750;margin-top:6px}
.kpi-card.primary .kpi-num{color:#00983A}
.groups{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:13px;margin:0 0 16px}
.gcard{background:#fff;border:1px solid #E0E5EA;border-radius:10px;padding:16px 18px 15px;
 min-width:0;box-shadow:0 1px 2px rgba(12,34,64,.05);text-align:center}
.gname{color:#3C4D60;font-size:13px;font-weight:650;min-height:34px;line-height:1.35;
 display:flex;align-items:center;justify-content:center}
.gpercent{color:#143257;font-size:29px;line-height:1.15;font-weight:760;margin:5px 0 12px}
.gbar{background:#E7EDF1;height:7px;border-radius:20px;overflow:hidden}
.gfill{background:#00983A;height:100%;border-radius:20px}
.gfoot{display:flex;justify-content:space-between;gap:8px;color:#6B7785;font-size:11px;margin-top:12px}
.gfoot span{flex:1;min-width:0;text-align:center}
div[data-testid="stFileUploader"]{background:#fff;border-radius:9px}
[data-testid="stSidebar"] .stDownloadButton button{width:100%;background:#004881;color:white;
 border:1px solid #004881;border-radius:8px}
[data-testid="stSidebar"] .stDownloadButton button:hover{background:#143257;color:white}
[data-testid="stDataFrame"]{background:#fff;border-radius:9px}
@media(max-width:1000px){.kpi-grid{grid-template-columns:repeat(2,minmax(0,1fr))}
 .groups{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media(max-width:650px){.hero{padding:18px 17px}.hero h1{font-size:22px}
 .kpi-grid,.groups{grid-template-columns:1fr}.gname{min-height:0}}
</style>
"""


def fmt_int(value: int) -> str:
    return f"{value:,}".replace(",", ".")


def fmt_pct(value: float) -> str:
    return f"{value:.2f}%".replace(".", ",")


def _cartoes_resumo(relatorio: dict) -> str:
    stats = [
        ("Aderência geral", fmt_pct(relatorio["aderencia"]), "primary"),
        ("Total Monitor", fmt_int(relatorio["monitor"]), ""),
        ("Identificados Autolab", fmt_int(relatorio["autolab"]), ""),
        ("Sem correspondência", fmt_int(relatorio["diferenca"]), ""),
    ]
    return '<div class="kpi-grid">' + "".join(
        f'<div class="kpi-card {estilo}"><div class="kpi-label">{label}</div>'
        f'<div class="kpi-num">{numero}</div></div>' for label, numero, estilo in stats
    ) + '</div>'


def _cartoes_grupos(relatorio: dict) -> str:
    cards = []
    for g in relatorio["grupos"]:
        nome = html_escape(g["grupo"])
        pct = min(100, max(0, float(g["aderencia"])))
        cards.append(
            '<div class="gcard">'
            f'<div class="gname">{nome}</div>'
            f'<div class="gpercent">{fmt_pct(g["aderencia"])}</div>'
            f'<div class="gbar"><div class="gfill" style="width:{pct:.4f}%"></div></div>'
            f'<div class="gfoot"><span>{fmt_int(g["autolab"])} / {fmt_int(g["monitor"])} no Autolab</span>'
            f'<span>{fmt_int(g["diferenca"])} sem vínculo</span></div></div>'
        )
    return '<div class="groups">' + "".join(cards) + '</div>'


COLUNAS_SEM = [
    ('hora', 'Data/hora'), ('grupo', 'Grupo'), ('material', 'Material'),
    ('equipamento', 'Equipamento'), ('local', 'Local'),
    ('caracteristica', 'Característica'), ('monitor', 'Valor Monitor'),
]


def linhas_sem_autolab(relatorio: dict, grupo: str = 'Todos os grupos') -> list[dict]:
    """Somente linhas do Monitor sem par no Autolab, preservando o Grupo original."""
    return [
        {titulo: item.get(chave, '') for chave, titulo in COLUNAS_SEM}
        for item in relatorio['detalhes']
        if item['tipo'] == 'sem'
        and (grupo == 'Todos os grupos' or item['grupo'] == grupo)
    ]


def _csv_sem_autolab(registros: list[dict]) -> bytes:
    buff = StringIO(newline='')
    campos = [coluna for _, coluna in COLUNAS_SEM]
    gravador = csv.DictWriter(buff, fieldnames=campos, delimiter=';', lineterminator='\r\n')
    gravador.writeheader()
    for item in registros:
        seguro = {}
        for chave, valor in item.items():
            txt = str(valor)
            # Não executar fórmulas ao abrir o CSV em um editor de planilhas.
            if txt and txt.lstrip().startswith(('=', '+', '-', '@')) and numero(txt) is None:
                txt = "'" + txt
            seguro[chave] = txt
        gravador.writerow(seguro)
    return buff.getvalue().encode('utf-8-sig')


def gerar_html(relatorio: dict) -> bytes:
    """Relatório HTML autônomo (CSS e tabela embutidos; funciona offline)."""
    dia = date.fromisoformat(relatorio['dia']).strftime('%d/%m/%Y')
    todos = linhas_sem_autolab(relatorio)
    cabecalho = ''.join(f'<th>{html_escape(titulo)}</th>' for _, titulo in COLUNAS_SEM)
    linhas = []
    for registro in todos:
        grupo = html_escape(str(registro['Grupo']), quote=True)
        valores = ''.join(
            f'<td>{html_escape(str(registro[titulo]))}</td>'
            for _, titulo in COLUNAS_SEM
        )
        linhas.append(f'<tr data-grupo="{grupo}">{valores}</tr>')
    opcoes = ['<option value="">Todos os grupos</option>']
    for g in relatorio['grupos']:
        nome = html_escape(g['grupo'], quote=True)
        opcoes.append(f'<option value="{nome}">{html_escape(g["grupo"])}</option>')
    estilos = CSS.replace('<style>', '', 1).replace('</style>', '', 1)
    conteudo = f"""<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Aderência Automação — {dia}</title>
<style>
{estilos}
body{{margin:0;padding:24px;background:#E9EDF1;color:#2B3542;font-family:Arial,Helvetica,sans-serif}}
.wrap{{max-width:1320px;margin:0 auto}}
.tabela-box{{background:#fff;border:1px solid #E0E5EA;border-radius:10px;overflow:auto;max-height:650px}}
table{{border-collapse:collapse;width:100%;font-size:12px;min-width:900px}}
th{{background:#143257;color:white;position:sticky;top:0;text-align:left;padding:11px 10px}}
td{{padding:9px 10px;border-top:1px solid #E0E5EA;vertical-align:top}}
tr:nth-child(even){{background:#F7F9FB}}
.controls{{display:flex;gap:12px;justify-content:space-between;align-items:center;margin:0 0 12px}}
select{{background:#fff;border:1px solid #CCD5DE;border-radius:8px;padding:10px;max-width:320px}}
.controls p{{color:#6B7785;font-size:12px;margin:0}}
@media(max-width:700px){{body{{padding:12px}}.controls{{flex-wrap:wrap}}}}
@media print{{.controls select{{display:none}}.tabela-box{{max-height:none}}}}
</style>
</head>
<body>
<main class="wrap">
<div class="hero"><h1>Aderência Automação</h1><p>{dia}</p></div>
{_cartoes_resumo(relatorio)}
<div class="subheading">Aderência por Grupo</div>
{_cartoes_grupos(relatorio)}
<div class="subheading">Consultar resultados sem Autolab</div>
<div class="controls"><select id="grupo" aria-label="Grupo" onchange="filtrarGrupo()">
{''.join(opcoes)}</select><p id="quantidade">{fmt_int(len(todos))} registros</p></div>
<div class="tabela-box"><table><thead><tr>{cabecalho}</tr></thead>
<tbody id="resultados">{''.join(linhas)}</tbody></table></div>
</main>
<script>
function filtrarGrupo() {{
 const selecionado = document.getElementById('grupo').value;
 const linhas = document.querySelectorAll('#resultados tr');
 let total = 0;
 for (const linha of linhas) {{
  const mostrar = !selecionado || linha.dataset.grupo === selecionado;
  linha.hidden = !mostrar;
  if(mostrar) total++;
 }}
 document.getElementById('quantidade').textContent = total.toLocaleString('pt-BR') + ' registros';
}}
</script>
</body></html>"""
    return conteudo.encode('utf-8')


def main() -> None:
    import streamlit as st

    st.set_page_config(page_title='Aderência Automação', page_icon='📊',
                       layout='wide', initial_sidebar_state='expanded')
    st.markdown(CSS, unsafe_allow_html=True)

    with st.sidebar:
        st.markdown('### Planilhas')
        arquivo_monitor = st.file_uploader('Monitor de Digitação (.xlsx)', type=['xlsx'], key='mon')
        arquivo_auto = st.file_uploader('Resultados Autolab (.xlsx)', type=['xlsx'], key='aut')

    if arquivo_monitor is None or arquivo_auto is None:
        st.markdown('<div class="hero"><h1>Aderência Automação</h1></div>',
                    unsafe_allow_html=True)
        st.info('Importe as duas planilhas pela lateral para visualizar a aderência.')
        return

    if arquivo_auto.size > LIMITE_UPLOAD or arquivo_monitor.size > LIMITE_UPLOAD:
        st.error('Cada arquivo deve ter no máximo 25 MB.')
        return

    @st.cache_data(show_spinner=False, max_entries=12)
    def _datas(auto: bytes, mon: bytes) -> list[date]:
        return datas_em_comum(auto, mon)

    @st.cache_data(show_spinner=False, max_entries=12)
    def _processar(auto: bytes, mon: bytes, data_iso: str) -> dict:
        return analisar(auto, mon, date.fromisoformat(data_iso))

    auto_bytes = arquivo_auto.getvalue()
    mon_bytes = arquivo_monitor.getvalue()
    try:
        dias = _datas(auto_bytes, mon_bytes)
        with st.sidebar:
            if len(dias) == 1:
                dia = dias[0]
            else:
                dia = st.selectbox('Data dos resultados', dias,
                                   format_func=lambda d: d.strftime('%d/%m/%Y'))
        relatorio = _processar(auto_bytes, mon_bytes, dia.isoformat())
    except (ValueError, TypeError, KeyError, BadZipFile, ET.ParseError, IndexError) as exc:
        st.error(f'Não foi possível analisar os arquivos: {exc}')
        return

    st.markdown(
        f'<div class="hero"><h1>Aderência Automação</h1>'
        f'<p>{dia.strftime("%d/%m/%Y")}</p></div>',
        unsafe_allow_html=True,
    )

    if relatorio['monitor'] == 0:
        st.warning('O Monitor não possui resultados válidos na data selecionada.')
        return

    with st.sidebar:
        st.download_button(
            'Gerar HTML', data=gerar_html(relatorio),
            file_name=f'aderencia_automacao_{dia.isoformat()}.html',
            mime='text/html', type='primary', use_container_width=True,
        )

    st.markdown(_cartoes_resumo(relatorio), unsafe_allow_html=True)
    st.markdown('<div class="subheading">Aderência por Grupo</div>', unsafe_allow_html=True)
    st.markdown(_cartoes_grupos(relatorio), unsafe_allow_html=True)

    st.markdown('<div class="subheading">Consultar resultados sem Autolab</div>',
                unsafe_allow_html=True)
    grupos = ['Todos os grupos'] + [g['grupo'] for g in relatorio['grupos']]
    grupo = st.selectbox('Grupo', grupos)
    registros = linhas_sem_autolab(relatorio, grupo)
    st.caption(f'{fmt_int(len(registros))} registros sem correspondência')
    if registros:
        st.dataframe(registros, use_container_width=True, hide_index=True, height=420)
    else:
        st.info('Nenhum resultado sem Autolab neste grupo.')

    st.download_button(
        'Exportar tabela (CSV)',
        data=_csv_sem_autolab(registros),
        file_name=f'aderencia_sem_autolab_{dia.isoformat()}.csv',
        mime='text/csv', type='secondary',
    )


if __name__ == '__main__':
    main()
