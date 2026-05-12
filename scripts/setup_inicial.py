"""
setup_inicial.py
Executa UMA VEZ para preparar o ambiente web:
  1. Faz upload do arquivo histórico para Supabase Storage (planilhas/historico/)
  2. Processa todos os dados e grava KPIs no Supabase
  3. Cria o primeiro usuário admin (opcional)

Uso:
  python scripts/setup_inicial.py

Variáveis de ambiente (defina no terminal antes de rodar):
  set SUPABASE_URL=https://SEU_PROJETO.supabase.co
  set SUPABASE_SERVICE_KEY=sb_secret_...
"""
import os, sys, glob, subprocess

SUPABASE_URL = os.environ.get('SUPABASE_URL', '')
SUPABASE_KEY = os.environ.get('SUPABASE_SERVICE_KEY', '')

if not SUPABASE_URL or not SUPABASE_KEY:
    print('ERRO: defina SUPABASE_URL e SUPABASE_SERVICE_KEY antes de rodar.')
    print('  Windows: set SUPABASE_URL=https://... && set SUPABASE_SERVICE_KEY=sb_secret_...')
    print('  Mac/Linux: export SUPABASE_URL=https://... && export SUPABASE_SERVICE_KEY=sb_secret_...')
    sys.exit(1)

try:
    from supabase import create_client
except ImportError:
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'supabase'])
    from supabase import create_client

sb = create_client(SUPABASE_URL, SUPABASE_KEY)

base     = os.path.dirname(os.path.abspath(__file__))
proj_dir = os.path.join(base, '..')
dados_dir = os.path.join(proj_dir, 'dados')

# ── 1. Upload do arquivo histórico ────────────────────────────────────────────
print('\n[1/2] Upload do arquivo histórico para Supabase Storage...')
hist_files = sorted(glob.glob(os.path.join(dados_dir, '[Dd]ash*.xlsx')) +
                    glob.glob(os.path.join(dados_dir, '[Dd]ash*.xls')))

if not hist_files:
    print('  Nenhum arquivo "Dash*.xlsx" encontrado em dados/. Pulando.')
else:
    hist_local = hist_files[-1]
    nome = os.path.basename(hist_local)
    storage_path = f'historico/{nome}'
    print(f'  Enviando: {nome}')
    with open(hist_local, 'rb') as f:
        conteudo = f.read()
    try:
        sb.storage.from_('planilhas').upload(storage_path, conteudo,
            {'upsert': 'true', 'content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})
        print(f'  Salvo em planilhas/{storage_path}')
    except Exception as e:
        try:
            sb.storage.from_('planilhas').remove([storage_path])
            sb.storage.from_('planilhas').upload(storage_path, conteudo,
                {'content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})
            print(f'  Atualizado em planilhas/{storage_path}')
        except Exception as e2:
            print(f'  Aviso: {e2}')

# ── 2. Processar e gravar KPIs no Supabase ───────────────────────────────────
print('\n[2/2] Processando dados e gravando no Supabase...')
env = os.environ.copy()
result = subprocess.run(
    [sys.executable, os.path.join(base, 'processar_supabase.py')],
    env=env
)
if result.returncode == 0:
    print('\nSetup concluido! Dados carregados no Supabase.')
    print('   Proximo passo: acesse o Vercel e faca o deploy do repositorio.')
else:
    print('\nErro no processamento. Verifique os logs acima.')
