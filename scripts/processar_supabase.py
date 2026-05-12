"""
processar_supabase.py
Baixa as planilhas mais recentes do Supabase Storage, processa e grava
os KPIs no banco. Também regenera o dashboard.html local em paralelo.

Uso: python scripts/processar_supabase.py
Variáveis de ambiente necessárias:
  SUPABASE_URL          — URL do projeto Supabase
  SUPABASE_SERVICE_KEY  — Service role key (NÃO a anon key)
"""
import os
import sys
import json
import tempfile
from datetime import datetime, timezone

from supabase import create_client, Client

# Importa funções de processamento do script local
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from gerar_dashboard import (
    calcular, ler_planilha2, ler_eloca, ler_erp_crm,
    ler_cac, ler_faturamento, ler_crm, gerar_html, carregar_logo
)

SUPABASE_URL = os.environ.get('SUPABASE_URL', '').rstrip('/')
SUPABASE_KEY = os.environ.get('SUPABASE_SERVICE_KEY', '')


def get_sb() -> Client:
    if not SUPABASE_URL or not SUPABASE_KEY:
        sys.exit('ERRO: defina SUPABASE_URL e SUPABASE_SERVICE_KEY')
    return create_client(SUPABASE_URL, SUPABASE_KEY)


def baixar_arquivo(sb: Client, storage_path: str, dest: str):
    data = sb.storage.from_('planilhas').download(storage_path)
    with open(dest, 'wb') as f:
        f.write(data)


def baixar_historico(sb: Client, dest_dir: str) -> str | None:
    """Tenta baixar o arquivo histórico do bucket planilhas/historico/.
    Retorna o caminho local se bem-sucedido, ou None se não encontrado."""
    try:
        files = sb.storage.from_('planilhas').list('historico')
        xlsx_files = [f for f in (files or []) if f['name'].lower().endswith(('.xlsx', '.xls'))]
        if not xlsx_files:
            return None
        # Pega o mais recente por nome
        target = sorted(xlsx_files, key=lambda f: f['name'])[-1]
        storage_path = f"historico/{target['name']}"
        dest = os.path.join(dest_dir, target['name'])
        print(f'  [Histórico] Baixando do Supabase Storage: {target["name"]}')
        baixar_arquivo(sb, storage_path, dest)
        return dest
    except Exception as e:
        print(f'  [Histórico] Não encontrado no Storage: {e}')
        return None


def planilha_mais_recente(sb: Client, tipo: str) -> dict | None:
    resp = (sb.table('planilhas_importadas')
              .select('*')
              .eq('tipo', tipo)
              .order('data_planilha', desc=True)
              .limit(1)
              .execute())
    return resp.data[0] if resp.data else None


def ler_config(sb: Client) -> dict:
    resp = sb.table('configuracoes').select('chave,valor').execute()
    return {r['chave']: r['valor'] for r in resp.data}


def ler_propostas_sb(sb: Client) -> list:
    resp = (sb.table('propostas_manuais')
              .select('*')
              .eq('ativo', True)
              .execute())
    return resp.data


def build_manual(config: dict) -> dict:
    return {
        'marketing':    config.get('marketing', {}) or {},
        'inadimplencia': config.get('inadimplencia', {}) or {},
    }


def now_utc() -> str:
    return datetime.now(timezone.utc).isoformat()


def escrever_kpis(sb: Client, kpis: dict):
    ts = now_utc()

    # ── kpis_serie: uma linha por mês ─────────────────────────────────────
    for entry in kpis.get('serie', []):
        sb.table('kpis_serie').upsert({
            'ym': entry['ym'],
            'dados': entry,
            'processado_em': ts,
        }).execute()

    # ── pipeline ERP por mês ───────────────────────────────────────────────
    for entry in kpis.get('serie', []):
        ym = entry['ym']
        for status, qtd in (entry.get('pipeline') or {}).items():
            valor = float((entry.get('pipeline_val') or {}).get(status, 0))
            sb.table('pipeline_por_mes').upsert({
                'ym': ym, 'status': status, 'fonte': 'ERP',
                'qtd': int(qtd), 'valor': valor, 'processado_em': ts,
            }).execute()

    # ── pipeline CRM por mês ──────────────────────────────────────────────
    for ym, statuses in (kpis.get('crm_pipeline_mes') or {}).items():
        for status, v in statuses.items():
            sb.table('pipeline_por_mes').upsert({
                'ym': ym, 'status': status, 'fonte': 'CRM',
                'qtd': int(v['qtd']), 'valor': float(v['valor']),
                'processado_em': ts,
            }).execute()

    # ── top_clientes ──────────────────────────────────────────────────────
    sb.table('top_clientes').delete().neq('id', '00000000-0000-0000-0000-000000000000').execute()
    for c in kpis.get('top_clientes', []):
        sb.table('top_clientes').insert({
            'nome': c['nome'], 'valor_total': float(c['total']),
            'processado_em': ts,
        }).execute()

    # ── fat_fin ───────────────────────────────────────────────────────────
    fat_por_mes = (kpis.get('fat_fin') or {}).get('por_mes', {})
    for ym, info in fat_por_mes.items():
        for tipo, valor in (info.get('por_tipo') or {}).items():
            sb.table('fat_fin').upsert({
                'ym': ym, 'tipo': tipo, 'valor': float(valor),
                'processado_em': ts,
            }).execute()

    n = len(kpis.get('serie', []))
    print(f'  KPIs gravados: {n} meses no Supabase')


def marcar_processado(sb: Client, pid: str, registros: int, erro: str = None):
    sb.table('planilhas_importadas').update({
        'processado': erro is None,
        'processado_em': now_utc(),
        'registros_importados': registros,
        'erro': erro,
    }).eq('id', pid).execute()


def main():
    sb = get_sb()
    print('Conectado ao Supabase:', SUPABASE_URL)

    config          = ler_config(sb)
    propostas_sb    = ler_propostas_sb(sb)
    manual          = build_manual(config)

    # Converte propostas manuais do Supabase para o formato do calcular()
    prelim_crm = [{
        'ym':        p['ym'],
        'cliente':   p['cliente'],
        'vertical':  p['vertical'],
        'fase':      p['status'],
        'valor_total': float(p['valor']),
    } for p in propostas_sb]

    base      = os.path.dirname(os.path.abspath(__file__))
    proj_dir  = os.path.join(base, '..')
    dados_dir = os.path.join(proj_dir, 'dados')

    with tempfile.TemporaryDirectory() as tmpdir:
        # ── Arquivo histórico: tenta local primeiro, depois Supabase Storage ──
        hist_path = None
        if os.path.isdir(dados_dir):
            hist_files = [f for f in os.listdir(dados_dir)
                          if 'dash' in f.lower() and f.lower().endswith(('.xlsx', '.xls'))]
            if hist_files:
                hist_path = os.path.join(dados_dir, sorted(hist_files)[-1])
                print(f'  [Histórico] Local: {os.path.basename(hist_path)}')

        if not hist_path:
            hist_path = baixar_historico(sb, tmpdir)

        if not hist_path:
            sys.exit('ERRO: arquivo histórico não encontrado (local nem no Storage).'
                     ' Faça upload em planilhas/historico/ no Supabase.')
        arquivos = {}
        contagens = {}
        for tipo in ('ERP', 'Faturamento', 'CRM'):
            meta = planilha_mais_recente(sb, tipo)
            if meta:
                dest = os.path.join(tmpdir, meta['nome_arquivo'])
                baixar_arquivo(sb, meta['storage_path'], dest)
                arquivos[tipo] = {'path': dest, 'meta': meta}
                print(f'  [{tipo}] {meta["nome_arquivo"]}')

        print(f'[1/4] Histórico: {os.path.basename(hist_path)}')
        p2      = ler_planilha2(hist_path)
        erp_crm = ler_erp_crm(hist_path)
        cac     = ler_cac(hist_path)
        print(f'      {len(p2)} registros históricos, {len(erp_crm)} ERP_CRM')

        eloca_path = (arquivos.get('ERP') or {}).get('path')
        fat_path   = (arquivos.get('Faturamento') or {}).get('path')
        crm_path   = (arquivos.get('CRM') or {}).get('path')

        eloca      = ler_eloca(eloca_path)   if eloca_path else []
        faturamento= ler_faturamento(fat_path) if fat_path   else []
        crm        = ler_crm(crm_path)       if crm_path    else []
        contagens  = {'ERP': len(eloca), 'Faturamento': len(faturamento), 'CRM': len(crm)}

        print(f'[2/4] Calculando KPIs...')
        kpis = calcular(p2, eloca, erp_crm, cac,
                        faturamento=faturamento, manual=manual, crm=crm)

        print(f'[3/4] Gravando no Supabase...')
        escrever_kpis(sb, kpis)

        for tipo, info in arquivos.items():
            marcar_processado(sb, info['meta']['id'], contagens.get(tipo, 0))

        print(f'[4/4] Gerando dashboard.html local...')
        logo_b64  = carregar_logo(proj_dir)
        html      = gerar_html(kpis, logo_b64)
        out_path  = os.path.join(proj_dir, 'output', 'dashboard.html')
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(html)
        print(f'  Local: {out_path}')

    print('Processamento concluído.')


if __name__ == '__main__':
    main()
