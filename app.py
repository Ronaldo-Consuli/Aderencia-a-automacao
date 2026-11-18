#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Painel de Aderência Autolab × Monitor de Digitação para Streamlit.

Execute: streamlit run app.py

Regras:
- Monitor de Digitação é a base dos totais, separados por Grupo.
- Pareamento 1:1 por material, equipamento, local, tipo, característica e data/hora.
- Divergência do valor não impede a correspondência.
- Exclui Acidez apenas da Destilação, dos dois lados.
- Registros sem Autolab não são prova de digitação manual.

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
DATA_PADRAO = date(2026, 11, 16)
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


def escolher_data(entrada: str) -> date:
    for formato in ('%d/%m/%Y', '%Y-%m-%d'):
        try:
            return datetime.strptime(entrada, formato).date()
        except (ValueError, TypeError):
            pass
    raise ValueError('Data inválida. Use 16/11/2026 ou 2026-11-16.')


def chave_monitor(r):
    dt = dt_resultado(r.get('Dt. Resultado'))
    if dt is None:
        return None
    return (
        normalizar(r.get('Material')), normalizar(r.get('Equipamento')),
        normalizar(r.get('Local')), normalizar(r.get('Tipo de Análise')),
        normalizar(r.get('Característica')), dt,
    )


def chave_autolab(r):
    dt = dt_resultado(r.get('Data Resultado'))
    if dt is None:
        return None
    return (
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
.block-container{max-width:1320px;padding-top:1.4rem;padding-bottom:2rem}
#MainMenu,footer{visibility:hidden}
.hero{background:linear-gradient(118deg,#00983A 0%,#004881 64%,#083050 100%);
 border-radius:13px;padding:25px 29px;margin-bottom:16px;color:white}
.hero h1{font-size:27px;line-height:1.2;color:#fff;margin:0;font-weight:780;letter-spacing:-.4px}
.hero p{font-size:13px;color:#eaf3f7;margin:7px 0 0}
.subheading{color:#143257;font-size:19px;font-weight:730;margin:22px 0 13px}
.kpi-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:13px;margin:15px 0 4px}
.kpi-card{background:#fff;border:1px solid #E0E5EA;border-radius:10px;padding:17px 19px;
 box-shadow:0 1px 2px rgba(12,34,64,.06)}
.kpi-label{color:#6B7785;font-size:12px;font-weight:600}
.kpi-num{color:#143257;font-size:29px;line-height:1.3;font-weight:750;margin-top:6px}
.kpi-card.primary .kpi-num{color:#00983A}
.groups{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:13px;margin:0 0 16px}
.gcard{background:#fff;border:1px solid #E0E5EA;border-radius:10px;padding:16px 18px 15px;
 min-width:0;box-shadow:0 1px 2px rgba(12,34,64,.05)}
.gname{color:#3C4D60;font-size:13px;font-weight:650;min-height:34px;line-height:1.35}
.gpercent{color:#143257;font-size:29px;line-height:1.15;font-weight:760;margin:5px 0 12px}
.gbar{background:#E7EDF1;height:7px;border-radius:20px;overflow:hidden}
.gfill{background:#00983A;height:100%;border-radius:20px}
.gfoot{display:flex;justify-content:space-between;gap:9px;color:#6B7785;font-size:11px;margin-top:12px}
.note{color:#6B7785;font-size:12px;margin:11px 0 0}
/* Native widgets: compact and consistent */
div[data-testid="stFileUploader"]{background:#fff;border-radius:9px}
.stButton button[kind="primary"]{background:#004881;border-color:#004881}
[data-testid="stDataFrame"]{background:#fff;border-radius:9px}
@media(max-width:1000px){.kpi-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.groups{grid-template-columns:repeat(2,minmax(0,1fr))}}
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


COLUNAS = {
    'Sem Autolab': [
        ('hora', 'Data/hora'), ('grupo', 'Grupo'), ('material', 'Material'),
        ('equipamento', 'Equipamento'), ('local', 'Local'),
        ('caracteristica', 'Característica'), ('monitor', 'Valor Monitor'),
    ],
    'Valores divergentes': [
        ('hora', 'Data/hora'), ('grupo', 'Grupo'), ('material', 'Material'),
        ('local', 'Local'), ('caracteristica', 'Característica'),
        ('monitor', 'Valor Monitor'), ('autolab', 'Valor Autolab'), ('status', 'Status'),
    ],
    'Autolab sem Monitor': [
        ('hora', 'Data/hora'), ('grupo', 'Grupo/Área'), ('material', 'Material'),
        ('local', 'Local'), ('caracteristica', 'Característica'),
        ('autolab', 'Valor Autolab'), ('status', 'Status'),
    ],
}

TIPOS = {'Sem Autolab': 'sem', 'Valores divergentes': 'divergente',
         'Autolab sem Monitor': 'auto_sem'}


def _linhas_detalhe(relatorio: dict, tipo: str, grupo: str, busca: str) -> list[dict]:
    filtros = []
    termo = normalizar(busca)
    for reg in relatorio['detalhes']:
        if reg['tipo'] != TIPOS[tipo]:
            continue
        if grupo != 'Todos os grupos' and reg['grupo'] != grupo:
            continue
        if termo and termo not in normalizar(' '.join(str(v) for v in reg.values())):
            continue
        filtros.append({titulo: reg.get(chave, '') for chave, titulo in COLUNAS[tipo]})
    return filtros


def _csv(registros: list[dict], tipo: str) -> bytes:
    buff = StringIO(newline='')
    campos = [coluna for _, coluna in COLUNAS[tipo]]
    gravador = csv.DictWriter(buff, fieldnames=campos, delimiter=';', lineterminator='\r\n')
    gravador.writeheader()
    for item in registros:
        seguro = {}
        for k, val in item.items():
            txt = str(val)
            # Evita que conteúdo textual externo seja interpretado como fórmula no Excel.
            if txt and txt.lstrip().startswith(('=', '+', '@')):
                txt = "'" + txt
            seguro[k] = txt
        gravador.writerow(seguro)
    return buff.getvalue().encode('utf-8-sig')


def main() -> None:
    import streamlit as st

    st.set_page_config(page_title='Aderência Autolab × Monitor', page_icon='📊',
                       layout='wide', initial_sidebar_state='collapsed')
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown(
        '<div class="hero"><h1>Aderência Autolab × Monitor</h1>'
        '<p>Indicadores por Grupo · Monitor de Digitação como referência</p></div>',
        unsafe_allow_html=True,
    )

    with st.container(border=True):
        col_monitor, col_auto, col_dia = st.columns([1.35, 1.35, .75], gap='medium')
        with col_monitor:
            arquivo_monitor = st.file_uploader('Monitor de Digitação (.xlsx)', type=['xlsx'], key='mon')
        with col_auto:
            arquivo_auto = st.file_uploader('Resultados Autolab (.xlsx)', type=['xlsx'], key='aut')
        with col_dia:
            dia = st.date_input('Data dos resultados', value=DATA_PADRAO, format='DD/MM/YYYY')

    if arquivo_monitor is None or arquivo_auto is None:
        st.info('Carregue os dois arquivos Excel para visualizar a aderência.')
        return

    if arquivo_auto.size > LIMITE_UPLOAD or arquivo_monitor.size > LIMITE_UPLOAD:
        st.error('Cada arquivo deve ter no máximo 25 MB.')
        return

    @st.cache_data(show_spinner=False, max_entries=12)
    def _processar(auto: bytes, mon: bytes, data_iso: str) -> dict:
        return analisar(auto, mon, date.fromisoformat(data_iso))

    try:
        relatorio = _processar(arquivo_auto.getvalue(), arquivo_monitor.getvalue(), dia.isoformat())
    except (ValueError, TypeError, KeyError, BadZipFile, ET.ParseError) as exc:
        st.error(f'Não foi possível analisar os arquivos: {exc}')
        return

    if relatorio['monitor'] == 0:
        st.warning('O Monitor não possui resultados válidos na data selecionada.')
        return

    st.markdown(_cartoes_resumo(relatorio), unsafe_allow_html=True)
    st.markdown('<div class="subheading">Aderência por Grupo</div>', unsafe_allow_html=True)
    st.markdown(_cartoes_grupos(relatorio), unsafe_allow_html=True)

    st.markdown('<div class="subheading">Consultar resultados</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1.2, 1.65, 1.4], gap='medium')
    with c1:
        grupos = ['Todos os grupos'] + [g['grupo'] for g in relatorio['grupos']]
        grupo = st.selectbox('Grupo', grupos)
    with c2:
        tipo = st.selectbox('Consulta', list(TIPOS))
    with c3:
        busca = st.text_input('Buscar', placeholder='Material, característica, horário...')

    registros = _linhas_detalhe(relatorio, tipo, grupo, busca)
    st.caption(f'{fmt_int(len(registros))} registros · {tipo.lower()}')
    if registros:
        st.dataframe(registros, use_container_width=True, hide_index=True, height=375)
    else:
        st.info('Nenhum registro encontrado para este filtro.')

    st.download_button(
        'Exportar consulta (CSV)',
        data=_csv(registros, tipo),
        file_name=f"aderencia_{TIPOS[tipo]}_{relatorio['dia']}.csv",
        mime='text/csv',
        type='secondary',
    )
    st.markdown(
        '<div class="note">Acidez excluída apenas da Destilação. Divergências de valor contam na aderência. '
        'Ausência de correspondência não comprova digitação manual.</div>',
        unsafe_allow_html=True,
    )


if __name__ == '__main__':
    main()
