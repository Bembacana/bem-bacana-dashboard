/**
 * Edge Function: trigger-processing
 * Recebe uma requisição autenticada do dashboard e dispara
 * o workflow do GitHub Actions via repository_dispatch.
 *
 * Secrets necessários no Supabase (Settings > Edge Functions > Secrets):
 *   GITHUB_PAT   — Personal Access Token com permissão actions:write
 *   GITHUB_REPO  — ex: "BemBacanaLocacoes/bem-bacana-dashboard"
 */
import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
}

Deno.serve(async (req: Request) => {
  // Preflight CORS
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  try {
    // Verificar autenticação via Supabase Auth
    const authHeader = req.headers.get('Authorization')
    if (!authHeader) {
      return new Response(JSON.stringify({ error: 'Não autorizado' }), {
        status: 401, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      })
    }

    const supabase = createClient(
      Deno.env.get('SUPABASE_URL') ?? '',
      Deno.env.get('SUPABASE_ANON_KEY') ?? '',
      { global: { headers: { Authorization: authHeader } } }
    )

    const { data: { user }, error: authError } = await supabase.auth.getUser()
    if (authError || !user) {
      return new Response(JSON.stringify({ error: 'Sessão inválida' }), {
        status: 401, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      })
    }

    // Disparar GitHub Actions via repository_dispatch
    const githubPAT  = Deno.env.get('GITHUB_PAT') ?? ''
    const githubRepo = Deno.env.get('GITHUB_REPO') ?? ''

    if (!githubPAT || !githubRepo) {
      return new Response(JSON.stringify({ error: 'GitHub não configurado' }), {
        status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      })
    }

    const ghResponse = await fetch(
      `https://api.github.com/repos/${githubRepo}/dispatches`,
      {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${githubPAT}`,
          'Accept': 'application/vnd.github.v3+json',
          'Content-Type': 'application/json',
          'User-Agent': 'BemBacana-Dashboard',
        },
        body: JSON.stringify({
          event_type: 'processar-planilhas',
          client_payload: {
            triggered_by: user.email,
            triggered_at: new Date().toISOString(),
          }
        }),
      }
    )

    if (!ghResponse.ok) {
      const err = await ghResponse.text()
      console.error('GitHub error:', err)
      return new Response(JSON.stringify({ error: `GitHub: ${ghResponse.status}` }), {
        status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      })
    }

    // Registrar disparo na tabela de planilhas
    await supabase.from('planilhas_importadas')
      .update({ processado: false })
      .eq('processado', false)

    return new Response(
      JSON.stringify({ success: true, triggered_by: user.email }),
      { headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    )

  } catch (err) {
    console.error('Erro inesperado:', err)
    return new Response(JSON.stringify({ error: String(err) }), {
      status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' }
    })
  }
})
