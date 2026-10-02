"""
Rotina diária da campanha "Arrancada de Vendas", rodando na nuvem (Render Cron Job) —
não depende do Mac do usuário estar ligado.

1. Baixa os arquivos-fonte (pastas 25/ e 26/) do Google Drive via conta de serviço.
2. Processa tudo com a MESMA lógica de negócio do pipeline local (montar_comparativo.py).
3. Sobe a planilha consolidada de volta pro Drive.
4. Gera o dashboard e publica no Vercel (link público).
"""
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile

# IMPORTANTE: PASTA_ARRANCADA_BASE precisa estar definida ANTES de importar
# montar_comparativo, porque ele lê essa env var só uma vez, na hora do import.
TMP_ROOT = tempfile.mkdtemp(prefix='campanha_')
os.environ['PASTA_ARRANCADA_BASE'] = TMP_ROOT

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import drive_sync
from montar_comparativo import carregar_tudo, gerar_excel, MES_DISPLAY, META_MINIMA, EXCEL_BACKUPS_DIR

SCRATCH = os.path.dirname(os.path.abspath(__file__))
DEPLOY_DIR = os.path.join(TMP_ROOT, 'deploy')
DEPLOY_INDEX = os.path.join(DEPLOY_DIR, 'index.html')
REGULAMENTO_NOME_DRIVE = 'Arrancada_de_Vendas (1) 2.pdf'
REGULAMENTO_DEPLOY_NAME = 'regulamento-campanha.pdf'
EXCEL_NOME = 'Comparativo_Campanha_Jan25_vs_Jan26.xlsx'
TEMPLATE_PATH = os.path.join(SCRATCH, 'dashboard_template.html')
LOGO_LIGHT_B64_PATH = os.path.join(SCRATCH, 'logo_black_gold_b64.txt')
LOGO_DARK_B64_PATH = os.path.join(SCRATCH, 'logo_gold_b64.txt')
VERCEL_PROJECT_JSON = os.path.join(SCRATCH, 'vercel_project.json')

CONECTIVOS = {'da', 'de', 'do', 'das', 'dos', 'e'}


def pd_isna(v):
    try:
        return v is None or math.isnan(v)
    except Exception:
        return False


def nome_exibicao(nome):
    partes = nome.split()
    out = []
    for i, p in enumerate(partes):
        pl = p.lower()
        out.append(pl if (i > 0 and pl in CONECTIVOS) else pl.capitalize())
    return ' '.join(out)


def montar_dados_dashboard(meses_comuns, comps_mensais, comp_geral):
    por_unidade_mes = {}
    for mes in meses_comuns:
        comp = comps_mensais[mes]
        for _, row in comp.iterrows():
            nome = row['UNIDADE_CURTA']
            por_unidade_mes.setdefault(nome, {})[mes] = {
                'pontos': int(row['PONTOS']),
                'varPct': None if pd_isna(row['VAR_%']) else round(float(row['VAR_%']) * 100, 1),
                'abaixoMeta': bool(row['TOTAL_26'] < META_MINIMA),
            }

    ranking = []
    for _, row in comp_geral.iterrows():
        nome = row['UNIDADE_CURTA']
        meses_info = []
        for mes in meses_comuns:
            info = por_unidade_mes.get(nome, {}).get(mes, {'pontos': 0, 'varPct': None, 'abaixoMeta': True})
            meses_info.append({
                'mes': MES_DISPLAY[mes],
                'pontos': info['pontos'],
                'varPct': info['varPct'],
                'abaixoMeta': info['abaixoMeta'],
            })
        var_pct = row['VAR_%']
        ranking.append({
            'nome': nome_exibicao(nome),
            'varPct': None if pd_isna(var_pct) else round(float(var_pct) * 100, 1),
            'pontos': int(row['PONTOS']),
            'todosAbaixoMeta': all(m['abaixoMeta'] for m in meses_info),
            'meses': meses_info,
        })

    for i, r in enumerate(ranking):
        r['posicao'] = i + 1

    return {
        'meses': [MES_DISPLAY[m] for m in meses_comuns],
        'metaMinima': META_MINIMA,
        'ranking': ranking,
    }


def gerar_dashboard_html(dados):
    with open(TEMPLATE_PATH, encoding='utf-8') as f:
        template = f.read()
    with open(LOGO_LIGHT_B64_PATH, encoding='utf-8') as f:
        logo_light_b64 = f.read().strip()
    with open(LOGO_DARK_B64_PATH, encoding='utf-8') as f:
        logo_dark_b64 = f.read().strip()
    payload = json.dumps(dados, ensure_ascii=False)
    html = (template
            .replace('__DATA_JSON__', payload)
            .replace('__LOGO_LIGHT_B64__', logo_light_b64)
            .replace('__LOGO_DARK_B64__', logo_dark_b64))
    return html


def gerar_html_standalone(fragmento, out_path):
    titulo_m = re.search(r'<title>.*?</title>', fragmento, re.S)
    style_m = re.search(r'<style>.*?</style>', fragmento, re.S)
    resto = fragmento
    if titulo_m:
        resto = resto.replace(titulo_m.group(0), '')
    if style_m:
        resto = resto.replace(style_m.group(0), '')
    head = (titulo_m.group(0) if titulo_m else '') + '\n' + (style_m.group(0) if style_m else '')
    doc = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
{head}
</head>
<body>
{resto}
</body>
</html>
"""
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(doc)
    return out_path


def publicar_vercel():
    os.makedirs(os.path.join(DEPLOY_DIR, '.vercel'), exist_ok=True)
    shutil.copyfile(VERCEL_PROJECT_JSON, os.path.join(DEPLOY_DIR, '.vercel', 'project.json'))

    token = os.environ.get('VERCEL_TOKEN')
    if not token:
        raise RuntimeError('VERCEL_TOKEN não está definida.')

    for tentativa in (1, 2):
        resultado = subprocess.run(
            ['npx', 'vercel', '--prod', '--yes', '--token', token,
             '--scope', 'silvanopiacentine1-9451s-projects'],
            cwd=DEPLOY_DIR, capture_output=True, text=True,
        )
        print(resultado.stdout)
        print(resultado.stderr, file=sys.stderr)
        if resultado.returncode == 0:
            return
        print(f"AVISO: tentativa {tentativa} de publicação falhou.")
        if tentativa == 1:
            import time
            time.sleep(15)
    raise RuntimeError('Publicação no Vercel falhou (2 tentativas).')


def main():
    service = drive_sync.get_service()
    raiz_id = drive_sync.find_root_folder(service, 'Arrancada de Vendas')
    pasta_25 = drive_sync.find_child(service, raiz_id, '25', drive_sync.FOLDER_MIME)
    pasta_26 = drive_sync.find_child(service, raiz_id, '26', drive_sync.FOLDER_MIME)
    if not pasta_25 or not pasta_26:
        raise RuntimeError('Subpastas "25" e/ou "26" não encontradas dentro de "Arrancada de Vendas".')

    print('Baixando arquivos-fonte do Drive...')
    drive_sync.baixar_pasta(service, pasta_25['id'], os.path.join(TMP_ROOT, '25'))
    drive_sync.baixar_pasta(service, pasta_26['id'], os.path.join(TMP_ROOT, '26'))

    regulamento_item = drive_sync.find_child(service, raiz_id, REGULAMENTO_NOME_DRIVE)
    regulamento_local = None
    if regulamento_item:
        regulamento_local = os.path.join(TMP_ROOT, REGULAMENTO_NOME_DRIVE)
        drive_sync.baixar_arquivo(service, regulamento_item['id'], regulamento_item['mimeType'], regulamento_local)

    meses_comuns, dados_mes, comps_mensais, comp_geral = carregar_tudo()

    # Baixa a versão anterior do Excel (se existir) ANTES de gerar a nova — gerar_excel()
    # compara contra ela pra avisar se a atualização nova tem menos meses que a última
    # (ver "Proteção contra perda de dados" na memória do projeto). Sem isso, rodando
    # num container efêmero, essa checagem nunca teria nada pra comparar.
    excel_local = os.path.join(TMP_ROOT, EXCEL_NOME)
    excel_anterior_item = drive_sync.find_child(service, raiz_id, EXCEL_NOME)
    if excel_anterior_item:
        drive_sync.baixar_arquivo(service, excel_anterior_item['id'], excel_anterior_item['mimeType'], excel_local)

    gerar_excel(meses_comuns, dados_mes, comps_mensais, comp_geral, excel_local)
    print('Planilha gerada, enviando de volta pro Drive...')
    drive_sync.enviar_arquivo(
        service, raiz_id, EXCEL_NOME, excel_local,
        mime_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )

    # Sobe também o backup timestampado que gerar_excel() acabou de criar localmente
    # (senão ele seria perdido junto com o container efêmero no fim da rodada).
    if os.path.isdir(EXCEL_BACKUPS_DIR):
        backups_folder_id = drive_sync.achar_ou_criar_pasta(service, raiz_id, 'backups')
        for nome_arquivo in os.listdir(EXCEL_BACKUPS_DIR):
            caminho = os.path.join(EXCEL_BACKUPS_DIR, nome_arquivo)
            if os.path.isfile(caminho):
                drive_sync.enviar_arquivo(
                    service, backups_folder_id, nome_arquivo, caminho,
                    mime_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                )

    os.makedirs(DEPLOY_DIR, exist_ok=True)
    if regulamento_local:
        shutil.copyfile(regulamento_local, os.path.join(DEPLOY_DIR, REGULAMENTO_DEPLOY_NAME))

    dados_dashboard = montar_dados_dashboard(meses_comuns, comps_mensais, comp_geral)
    dados_dashboard['regulamentoUrl'] = REGULAMENTO_DEPLOY_NAME if regulamento_local else None
    fragmento = gerar_dashboard_html(dados_dashboard)
    gerar_html_standalone(fragmento, DEPLOY_INDEX)

    print('Publicando no Vercel...')
    publicar_vercel()

    print('Meses processados:', [MES_DISPLAY[m] for m in meses_comuns])
    print('Concluído com sucesso.')


if __name__ == '__main__':
    try:
        main()
    finally:
        shutil.rmtree(TMP_ROOT, ignore_errors=True)
