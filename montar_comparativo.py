import os
import pandas as pd
import re
import glob
import shutil
import datetime
import unicodedata
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

# Em produção no Render (cron na nuvem), PASTA_ARRANCADA_BASE aponta pra uma pasta local
# temporária onde os arquivos já foram baixados do Drive via API (ver main.py). No Mac,
# sem a env var, cai no caminho local de sempre (Google Drive Desktop montado).
PASTA_ARRANCADA_BASE = os.environ.get(
    'PASTA_ARRANCADA_BASE',
    '/Users/silvanopiacentine/Library/CloudStorage/GoogleDrive-piaseg@piasegmais.com.br/Meu Drive/Piaseg Franchising/TI/Arrancada de Vendas ',
)
PASTA_25 = PASTA_ARRANCADA_BASE + '/25'
PASTA_26 = PASTA_ARRANCADA_BASE + '/26'

# Backup permanente dos arquivos-fonte: toda vez que um mês é visto numa pasta "ao vivo",
# guardamos uma cópia aqui. Se o arquivo original sumir/for substituído depois (já
# aconteceu — ver memória do projeto), a rodada seguinte ainda encontra o mês via backup
# em vez de simplesmente perder aquele mês do Comparativo Geral.
PIPELINE_DIR = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_BACKUP_25 = os.path.join(PIPELINE_DIR, 'arquivo_fonte_backup', '25')
ARQUIVO_BACKUP_26 = os.path.join(PIPELINE_DIR, 'arquivo_fonte_backup', '26')

# Backups timestampados do Excel consolidado, feitos antes de cada sobrescrita.
EXCEL_BACKUPS_DIR = PASTA_ARRANCADA_BASE + '/backups'
EXCEL_BACKUPS_MANTER = 20
RAMOS_EXCLUIDOS = {
    'CAPITALIZACAO',
    'CONSORCIO',
    'FINANCIAMENTO',
    'RURAL',  # nome do ramo no formato antigo (Jan-Jun/26); mantido por compatibilidade
    'SEGURO AGRICOLA COM COBERTURA DO FESR',
    'SEGURO AGRICOLA SEM COBERTURA DO FESR',
    'SEGURO DE FLORESTAS COM FESR',
    'SEGURO DE FLORESTAS SEM FESR',
}
META_MINIMA = 30000
CRESCIMENTO_ALTO = 0.20  # acima disso = 5 pontos

# Unidades que não são franqueados de verdade (matriz, contas internas etc.) e devem
# ser sempre excluídas do comparativo, independente do mês.
UNIDADES_EXCLUIDAS = {
    'DOURADOS', 'CAMPO GRANDE', 'PIASEG CONSULTORIA', 'IVAN/ROQUE',
    'STUDIO AGRONEGOCIOS', 'STUDIO AGRONEGOCIOS LTDA',
}


def normalizar_unidade(nome):
    n = normalizar(nome)
    n = re.sub(r'\s*/\s*', '/', n)
    n = re.sub(r'\s+', ' ', n).strip()
    return n

MESES_ORDEM = ['JANEIRO', 'FEVEREIRO', 'MARCO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO',
               'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO']
MES_DISPLAY = {
    'JANEIRO': 'Janeiro', 'FEVEREIRO': 'Fevereiro', 'MARCO': 'Março', 'ABRIL': 'Abril',
    'MAIO': 'Maio', 'JUNHO': 'Junho', 'JULHO': 'Julho', 'AGOSTO': 'Agosto',
    'SETEMBRO': 'Setembro', 'OUTUBRO': 'Outubro', 'NOVEMBRO': 'Novembro', 'DEZEMBRO': 'Dezembro',
}

header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
header_font = Font(color="FFFFFF", bold=True, size=11)
thin = Side(style='thin', color="D9D9D9")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
cores_pontos = {5: "70AD47", 3: "C6E0B4", 1: "FFE699", 0: "F8CBAD"}


def normalizar(txt):
    txt = str(txt).strip().upper()
    return ''.join(c for c in unicodedata.normalize('NFD', txt) if unicodedata.category(c) != 'Mn')


def nome_curto(unidade):
    # A unidade sempre vem como "NOME - PIASEG CONSULTORIA" (às vezes truncado
    # ou com um segmento extra, ex: "NOME - FRANQUEADO - PIASEG CONSULTORI").
    # Basta pegar o primeiro segmento antes do " - ".
    return str(unidade).split(' - ')[0].strip()


def _listar_excels(pasta):
    exts = ('*.XLS', '*.xls', '*.XLSX', '*.xlsx')
    arquivos = []
    for ext in exts:
        arquivos.extend(glob.glob(f'{pasta}/{ext}'))
    return arquivos


def detectar_meses_disponiveis(pasta):
    """Varre a pasta e retorna {MES_NORMALIZADO: caminho_do_arquivo} para os arquivos
    cujo *nome* já identifica o mês (ex: "Setembro 26.XLS")."""
    encontrados = {}
    for caminho in _listar_excels(pasta):
        nome_norm = normalizar(os.path.basename(caminho))
        for mes in MESES_ORDEM:
            if mes in nome_norm:
                encontrados[mes] = caminho
                break
    return encontrados


LIMIAR_MINIMO_LINHAS_MES = 15  # abaixo disso é só sobra de mês vizinho, não um mês de verdade


def sincronizar_arquivos_sem_nome_de_mes(pasta):
    """A automação de download diário solta arquivos com nome gerado pelo sistema
    (ex: "RptAnaliseProducao (9).XLS"), sem nenhum mês no nome — `detectar_meses_disponiveis`
    simplesmente os ignora, e a atualização parece "travada" mesmo com arquivo novo chegando
    todo dia (foi exatamente o que aconteceu em 2026-10-01).

    Esta função varre TODOS esses arquivos sem nome de mês (não só o mais recente — o robô
    às vezes exporta num formato mais completo, às vezes num mais enxuto, sem relação direta
    com a data do arquivo), descobre pelo conteúdo (Início de Vigência) quais meses cada um
    cobre, e para cada (ano, mês) usa o candidato com **mais linhas** — se isso superar o que
    já existe como arquivo nomeado, gera/substitui `"<Mês> <aa>.xlsx"` na mesma pasta. Arquivos
    nomeados continuam podendo ser a fonte, nada aqui os apaga; só cria/atualiza quando o
    candidato sem nome é mais completo.
    """
    ano_pasta = os.path.basename(pasta.rstrip('/'))
    if not ano_pasta.isdigit():
        return
    ano_completo = 2000 + int(ano_pasta)

    ja_nomeados = set()
    for caminho in _listar_excels(pasta):
        nome_norm = normalizar(os.path.basename(caminho))
        if any(mes in nome_norm for mes in MESES_ORDEM):
            ja_nomeados.add(caminho)

    candidatos = [c for c in _listar_excels(pasta) if c not in ja_nomeados]
    if not candidatos:
        return

    # melhor candidato por mês: {MES_NORMALIZADO: (qtd_linhas, dataframe_original)}
    melhor_por_mes = {}
    for caminho in candidatos:
        try:
            df = pd.read_excel(caminho, sheet_name='RptAnaliseProducao', header=0)
        except Exception as e:
            print(f"AVISO: não consegui ler '{os.path.basename(caminho)}' (sem nome de mês) — {e}")
            continue
        df_mapeado = mapear_colunas(df)
        if 'INICIO_VIGENCIA' not in df_mapeado.columns:
            continue
        datas = pd.to_datetime(df_mapeado['INICIO_VIGENCIA'], errors='coerce')
        for (ano, mes_num), indices in df_mapeado.groupby([datas.dt.year, datas.dt.month]).groups.items():
            if pd.isna(ano) or pd.isna(mes_num) or int(ano) != ano_completo:
                continue
            if len(indices) < LIMIAR_MINIMO_LINHAS_MES:
                continue
            mes_chave = MESES_ORDEM[int(mes_num) - 1]
            atual = melhor_por_mes.get(mes_chave)
            if atual is None or len(indices) > atual[0]:
                melhor_por_mes[mes_chave] = (len(indices), df.loc[indices])

    for mes_chave, (qtd, subset) in melhor_por_mes.items():
        destino = os.path.join(pasta, f'{MES_DISPLAY[mes_chave]} {ano_pasta}.xlsx')
        qtd_existente = 0
        if os.path.exists(destino) or mes_chave in detectar_meses_disponiveis(pasta):
            caminho_existente = detectar_meses_disponiveis(pasta).get(mes_chave, destino)
            try:
                qtd_existente = len(pd.read_excel(caminho_existente, sheet_name='RptAnaliseProducao', header=0))
            except Exception:
                qtd_existente = 0
        if qtd <= qtd_existente:
            continue
        with pd.ExcelWriter(destino, engine='openpyxl') as writer:
            subset.to_excel(writer, sheet_name='RptAnaliseProducao', index=False)
        print(f"Gerado/atualizado '{os.path.basename(destino)}' a partir de arquivo(s) sem nome de mês "
              f"({qtd} linhas, antes {qtd_existente}).")


def detectar_meses_com_backup(pasta_ao_vivo, pasta_backup):
    """Une o que está na pasta 'ao vivo' com o que já foi arquivado antes.

    A pasta ao vivo tem prioridade quando o mesmo mês existe nos dois lugares
    (o usuário pode ter substituído o arquivo por uma versão corrigida). Todo
    mês encontrado ao vivo é copiado para o backup nesta mesma passada, então
    mesmo que ele seja apagado depois, a próxima rodada ainda o encontra aqui.
    """
    os.makedirs(pasta_backup, exist_ok=True)
    sincronizar_arquivos_sem_nome_de_mes(pasta_ao_vivo)
    ao_vivo = detectar_meses_disponiveis(pasta_ao_vivo)
    for mes, caminho in ao_vivo.items():
        extensao = os.path.splitext(caminho)[1] or '.XLS'
        try:
            shutil.copyfile(caminho, os.path.join(pasta_backup, f'{mes}{extensao}'))
        except OSError:
            pass  # não deixa um problema de cópia derrubar a atualização
    do_backup = detectar_meses_disponiveis(pasta_backup)
    combinado = dict(do_backup)
    combinado.update(ao_vivo)
    apenas_no_backup = sorted(set(do_backup) - set(ao_vivo), key=MESES_ORDEM.index)
    return combinado, apenas_no_backup


def mapear_colunas(df):
    """Mapeia os cabeçalhos reais do relatório para nomes canônicos. Formatos de
    exportação já vistos (o sistema de origem muda o nome das colunas com frequência):
      - Antigo (5 col.): Prêmio, Início de Vigência, Corretora, Unidade de Negócio, Célula
      - Julho/26 em diante (6 col.): Prêmio, Início de Vigência, Corretora, Seguradora, Divisão, Ramo

    Não mapeia "Grupo de Produção" para UNIDADE aqui: essa coluna é o VENDEDOR/sub-agente
    dentro do franqueado, não o franqueado em si (confirmado comparando com "Divisão" em
    arquivos que trazem as duas colunas — ex: "EMELY TAYLOR / TERRA" em Grupo de Produção
    vs "TERRA - PIASEG CONSULTORIA" em Divisão). Ver `resolver_unidade_grupo_producao()`
    para como ela é tratada quando é a única coluna disponível.
    "Divisão" = franqueado. "Ramo" (quando não é a coluna do franqueado) / "Célula" = ramo do seguro.
    """
    mapa = {}
    for col in df.columns:
        n = normalizar(col)
        if 'PREMIO' in n:
            mapa[col] = 'PREMIO'
        elif 'VIGENCIA' in n:
            mapa[col] = 'INICIO_VIGENCIA'
        elif n == 'CORRETORA':
            mapa[col] = 'CORRETORA'
        elif n == 'SEGURADORA':
            mapa[col] = 'SEGURADORA'
        elif n == 'DIVISAO' or 'UNIDADE' in n:
            mapa[col] = 'UNIDADE'
        elif n == 'RAMO' or 'CELULA' in n:
            mapa[col] = 'CELULA'
    return df.rename(columns=mapa)


class ColunaFranqueadoAusente(Exception):
    """O relatório não tem coluna de franqueado (Unidade de Negócio/Divisão) —
    sem ela é impossível atribuir as vendas a alguém, então o mês não pode ser processado."""


def construir_crosswalk_grupo_divisao(caminhos):
    """Constrói um dicionário {Grupo de Produção -> Divisão} a partir de qualquer arquivo
    que traga as duas colunas ao mesmo tempo (ex: Agosto/26 passou a trazer "Grupo de
    Produção" além de "Divisão"). Usado como fallback em meses cujo export só tem
    "Grupo de Produção" (ex: Setembro/26 em diante), para não fragmentar o mesmo
    franqueado em vários "sub-vendedores" diferentes no ranking.
    """
    crosswalk = {}
    for caminho in caminhos:
        try:
            df = pd.read_excel(caminho, sheet_name='RptAnaliseProducao', header=0)
        except Exception:
            continue
        cols = {normalizar(c): c for c in df.columns}
        if 'DIVISAO' not in cols or 'GRUPO DE PRODUCAO' not in cols:
            continue
        moda = df.groupby(cols['GRUPO DE PRODUCAO'])[cols['DIVISAO']].agg(lambda s: s.mode()[0])
        for grupo, divisao in moda.items():
            crosswalk.setdefault(str(grupo).strip(), divisao)
    return crosswalk


def resolver_unidade_grupo_producao(df_raw, crosswalk):
    """Quando o arquivo só tem 'Grupo de Produção' (sem 'Divisão'), usa o crosswalk pra
    converter pro franqueado correto; o que não está no crosswalk (vendedor novo, por
    exemplo) fica como está — melhor esforço, tratado como o próprio franqueado."""
    col_grupo = next((c for c in df_raw.columns if normalizar(c) == 'GRUPO DE PRODUCAO'), None)
    if col_grupo is None:
        return None
    return df_raw[col_grupo].apply(lambda v: crosswalk.get(str(v).strip(), v))


def carregar_de_caminho(caminho, crosswalk_grupo_divisao=None):
    df_raw = pd.read_excel(caminho, sheet_name='RptAnaliseProducao', header=0)
    df = mapear_colunas(df_raw)
    if 'UNIDADE' not in df.columns:
        unidade_resolvida = resolver_unidade_grupo_producao(df_raw, crosswalk_grupo_divisao or {})
        if unidade_resolvida is None:
            raise ColunaFranqueadoAusente(
                f"{os.path.basename(caminho)}: colunas encontradas foram {list(df_raw.columns)}, "
                f"sem nenhuma de franqueado (Unidade de Negócio / Divisão / Grupo de Produção)."
            )
        df['UNIDADE'] = unidade_resolvida
    if 'CELULA' not in df.columns:
        df['CELULA'] = ''
    df = df[~df['CELULA'].apply(normalizar).isin(RAMOS_EXCLUIDOS)].reset_index(drop=True)
    df['UNIDADE_CURTA'] = df['UNIDADE'].apply(nome_curto)
    df = df[~df['UNIDADE_CURTA'].apply(normalizar_unidade).isin(UNIDADES_EXCLUIDAS)].reset_index(drop=True)
    return df


def pontuar(r):
    if r['TOTAL_26'] < META_MINIMA:
        return 0
    if r['TOTAL_25'] == 0:
        return 5 if r['TOTAL_26'] > 0 else 0
    pct = r['VAR_%']
    if pct > CRESCIMENTO_ALTO:
        return 5
    if pct > 0:
        return 3
    if pct == 0:
        return 1
    return 0


def montar_comparativo_mensal(df25, df26):
    """Comparativo de um único mês, com pontuação calculada a partir da meta mínima do mês."""
    agg25 = df25.groupby('UNIDADE_CURTA')['PREMIO'].agg(['sum', 'count']).rename(columns={'sum': 'TOTAL_25', 'count': 'QTD_25'})
    agg26 = df26.groupby('UNIDADE_CURTA')['PREMIO'].agg(['sum', 'count']).rename(columns={'sum': 'TOTAL_26', 'count': 'QTD_26'})
    comp = agg25.join(agg26, how='outer').fillna(0)
    comp['QTD_25'] = comp['QTD_25'].astype(int)
    comp['QTD_26'] = comp['QTD_26'].astype(int)
    comp['VAR_R$'] = comp['TOTAL_26'] - comp['TOTAL_25']
    comp['VAR_%'] = comp.apply(lambda r: (r['VAR_R$'] / r['TOTAL_25']) if r['TOTAL_25'] != 0 else None, axis=1)
    comp['PONTOS'] = comp.apply(pontuar, axis=1)
    comp = ordenar_e_ranquear(comp)
    return comp


def ordenar_e_ranquear(comp):
    comp = comp.copy()
    comp['VAR_%_DESEMPATE'] = comp['VAR_%'].fillna(float('inf'))
    tem_unidade_coluna = 'UNIDADE_CURTA' in comp.columns
    comp = comp.sort_values(['PONTOS', 'VAR_%_DESEMPATE'], ascending=[False, False])
    comp = comp.reset_index() if not tem_unidade_coluna else comp.reset_index(drop=True)
    comp['POSICAO_TXT'] = (comp.index + 1).map(lambda n: f"{n}º")
    return comp


def montar_comparativo_geral(comps_mensais):
    """Soma Pontos e Totais de cada aba mensal (Comparativo Janeiro, Fevereiro, ...)."""
    partes = []
    for comp in comps_mensais:
        partes.append(comp[['UNIDADE_CURTA', 'TOTAL_25', 'QTD_25', 'TOTAL_26', 'QTD_26', 'PONTOS']])
    todos = pd.concat(partes, ignore_index=True)
    geral = todos.groupby('UNIDADE_CURTA').sum(numeric_only=True)
    geral['VAR_R$'] = geral['TOTAL_26'] - geral['TOTAL_25']
    geral['VAR_%'] = geral.apply(lambda r: (r['VAR_R$'] / r['TOTAL_25']) if r['TOTAL_25'] != 0 else None, axis=1)
    geral = ordenar_e_ranquear(geral)
    return geral


def escrever_comparativo(wb, titulo, comp):
    ws = wb.create_sheet(titulo)
    headers = ['Posição', 'Unidade de Negócio', 'Total Vendido Ano 25', 'Qtd Negócios Ano 25',
               'Total Vendido Ano 26', 'Qtd Negócios Ano 26', 'Variação (R$)', 'Variação (%)', 'Pontos']
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border

    for _, row in comp.iterrows():
        ws.append([
            row['POSICAO_TXT'], row['UNIDADE_CURTA'], row['TOTAL_25'], row['QTD_25'],
            row['TOTAL_26'], row['QTD_26'], row['VAR_R$'], row['VAR_%'], row['PONTOS'],
        ])

    n = ws.max_row
    total_row = n + 1
    ws.cell(row=total_row, column=2, value="TOTAL GERAL")
    ws.cell(row=total_row, column=3, value=f"=SUM(C2:C{n})")
    ws.cell(row=total_row, column=4, value=f"=SUM(D2:D{n})")
    ws.cell(row=total_row, column=5, value=f"=SUM(E2:E{n})")
    ws.cell(row=total_row, column=6, value=f"=SUM(F2:F{n})")
    ws.cell(row=total_row, column=7, value=f"=E{total_row}-C{total_row}")
    ws.cell(row=total_row, column=8, value=f"=IFERROR(G{total_row}/C{total_row},\"\")")
    ws.cell(row=total_row, column=9, value=f"=SUM(I2:I{n})")
    for c in range(1, 10):
        cell = ws.cell(row=total_row, column=c)
        cell.font = Font(bold=True)
        cell.fill = total_fill
        cell.border = border

    for r in range(2, total_row + 1):
        for c in (3, 5, 7):
            ws.cell(row=r, column=c).number_format = 'R$ #,##0.00'
        ws.cell(row=r, column=8).number_format = '0.0%'
        for c in (1, 4, 6, 9):
            ws.cell(row=r, column=c).alignment = Alignment(horizontal='center')
        for c in range(1, 10):
            ws.cell(row=r, column=c).border = border

    for col in ('G', 'H'):
        ws.conditional_formatting.add(f"{col}2:{col}{n}", CellIsRule(operator='lessThan', formula=['0'], fill=PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")))
        ws.conditional_formatting.add(f"{col}2:{col}{n}", CellIsRule(operator='greaterThanOrEqual', formula=['0'], fill=PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")))

    for pontos, cor in cores_pontos.items():
        ws.conditional_formatting.add(f"I2:I{n}", CellIsRule(operator='equal', formula=[str(pontos)], fill=PatternFill(start_color=cor, end_color=cor, fill_type="solid")))

    widths = [10, 32, 20, 18, 20, 18, 16, 14, 10]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:I{n}"
    return ws


def escrever_base(wb, titulo, df):
    ws = wb.create_sheet(titulo)
    ws.append(['Prêmio', 'Início de Vigência', 'Corretora', 'Unidade de Negócio', 'Célula'])
    for _, row in df.iterrows():
        ws.append([row['PREMIO'], row['INICIO_VIGENCIA'], row['CORRETORA'], row['UNIDADE'], row['CELULA']])
    for c in range(1, 6):
        ws.cell(row=1, column=c).font = Font(bold=True, color="FFFFFF")
        ws.cell(row=1, column=c).fill = header_fill
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 45
    ws.column_dimensions['D'].width = 40
    ws.column_dimensions['E'].width = 20
    return ws


def carregar_tudo():
    """Detecta os meses em comum entre as pastas 25/26, carrega e monta todos os comparativos.

    Usa a união da pasta "ao vivo" com o backup interno de arquivos-fonte (ver
    `detectar_meses_com_backup`), então um mês que suma da pasta do usuário depois de já
    ter sido processado uma vez continua aparecendo no Comparativo Geral.
    """
    meses_25, so_backup_25 = detectar_meses_com_backup(PASTA_25, ARQUIVO_BACKUP_25)
    meses_26, so_backup_26 = detectar_meses_com_backup(PASTA_26, ARQUIVO_BACKUP_26)
    meses_comuns = [m for m in MESES_ORDEM if m in meses_25 and m in meses_26]

    if not meses_comuns:
        raise RuntimeError('Nenhum mês em comum encontrado entre as pastas 25 e 26 (nem ao vivo, nem no backup).')

    so_no_backup = sorted(set(so_backup_25) | set(so_backup_26), key=MESES_ORDEM.index)
    if so_no_backup:
        nomes = ', '.join(MES_DISPLAY[m] for m in so_no_backup)
        print(f"AVISO: {nomes} não está mais na pasta Campanha/25 ou /26 — "
              f"usando a cópia de backup salva em rodadas anteriores.")

    crosswalk = construir_crosswalk_grupo_divisao(list(meses_25.values()) + list(meses_26.values()))

    dados_mes = {}      # mes -> (df25, df26)
    comps_mensais = {}  # mes -> comp
    meses_pulados = []  # [(mes, motivo)]
    for mes in meses_comuns:
        try:
            df25 = carregar_de_caminho(meses_25[mes], crosswalk)
            df26 = carregar_de_caminho(meses_26[mes], crosswalk)
        except ColunaFranqueadoAusente as e:
            meses_pulados.append((mes, str(e)))
            continue
        dados_mes[mes] = (df25, df26)
        comps_mensais[mes] = montar_comparativo_mensal(df25, df26)

    meses_comuns = [m for m in meses_comuns if m in comps_mensais]
    if meses_pulados:
        for mes, motivo in meses_pulados:
            print(f"AVISO: {MES_DISPLAY[mes]} NÃO foi processado — {motivo}")
    if not meses_comuns:
        raise RuntimeError('Nenhum mês pôde ser processado (todos sem coluna de franqueado).')

    comp_geral = montar_comparativo_geral(list(comps_mensais.values()))
    return meses_comuns, dados_mes, comps_mensais, comp_geral


def _backup_excel_existente(out_path):
    """Guarda uma cópia timestampada do Excel atual antes de sobrescrever, e avisa se a
    nova versão tem menos meses do que a anterior (sinal de possível perda de dados)."""
    if not os.path.exists(out_path):
        return
    os.makedirs(EXCEL_BACKUPS_DIR, exist_ok=True)
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    destino = os.path.join(EXCEL_BACKUPS_DIR, f'Comparativo_{ts}.xlsx')
    shutil.copyfile(out_path, destino)

    existentes = sorted(glob.glob(os.path.join(EXCEL_BACKUPS_DIR, 'Comparativo_*.xlsx')))
    for antigo in existentes[:-EXCEL_BACKUPS_MANTER]:
        os.remove(antigo)


def gerar_excel(meses_comuns, dados_mes, comps_mensais, comp_geral, out_path):
    meses_antigos = None
    if os.path.exists(out_path):
        wb_antigo = load_workbook(out_path, read_only=True)
        meses_antigos = [s for s in wb_antigo.sheetnames if s.startswith('Comparativo ') and s != 'Comparativo Geral']
        wb_antigo.close()

    _backup_excel_existente(out_path)

    wb = Workbook()
    wb.remove(wb.active)

    escrever_comparativo(wb, "Comparativo Geral", comp_geral)
    for mes in meses_comuns:
        escrever_comparativo(wb, f"Comparativo {MES_DISPLAY[mes]}", comps_mensais[mes])
    for mes in meses_comuns:
        df25, df26 = dados_mes[mes]
        sufixo = MES_DISPLAY[mes][:3]
        escrever_base(wb, f"Base {sufixo}25", df25)
        escrever_base(wb, f"Base {sufixo}26", df26)

    wb.save(out_path)

    if meses_antigos is not None and len(meses_antigos) > len(meses_comuns):
        print(f"AVISO: a versão anterior da planilha tinha {len(meses_antigos)} meses "
              f"({', '.join(meses_antigos)}) e esta tem {len(meses_comuns)}. "
              f"Um backup da versão anterior foi salvo em {EXCEL_BACKUPS_DIR}/.")

    return out_path


if __name__ == '__main__':
    meses_comuns, dados_mes, comps_mensais, comp_geral = carregar_tudo()
    out_path = gerar_excel(
        meses_comuns, dados_mes, comps_mensais, comp_geral,
        '/Users/silvanopiacentine/Desktop/Campanha/Comparativo_Campanha_Jan25_vs_Jan26.xlsx'
    )
    print("Salvo em:", out_path)
    print("Meses processados:", [MES_DISPLAY[m] for m in meses_comuns])
    print("\n=== COMPARATIVO GERAL (soma de Pontos e Total Vendido Ano 26 por mês) ===")
    print(comp_geral[['UNIDADE_CURTA', 'TOTAL_25', 'TOTAL_26', 'VAR_%', 'PONTOS', 'POSICAO_TXT']].to_string(index=False))
