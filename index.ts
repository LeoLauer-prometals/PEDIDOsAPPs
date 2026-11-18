// Edge Function "enviar-email-outlook" — Prometals Pedidos (Versão 6.2;
// Versão 6.3 adicionou suporte a e-mail em HTML — ver campo "html" abaixo)
//
// Por que essa function existe: até a Versão 6.1, a notificação por e-mail
// de "novo pedido"/"observação interna" era mandada pelo EmailJS, saindo
// pelo e-mail pessoal de quem configurou o serviço. Pedido do usuário: usar
// a mesma caixa de avisos (Outlook/Microsoft 365) que outro sistema da
// empresa já usa pra mandar e-mail — e do MESMO JEITO que esse outro
// sistema faz: chamando direto a API do Microsoft Graph (Outlook), com um
// aplicativo registrado no Azure AD (Tenant ID / Client ID / Client Secret).
//
// Essa chamada SÓ pode acontecer aqui, no servidor (nunca no navegador):
// o Client Secret dá, sozinho, poder de mandar e-mail por aquela caixa pra
// QUALQUER UM que o tiver — se ele fosse colado no HTML, qualquer pessoa que
// abrisse o código-fonte do site (F12 → ver código) poderia roubá-lo e
// mandar e-mail em nome da Prometals pra sempre, até alguém perceber e
// trocar o secret no Azure. Por isso ele mora só aqui, como "Secret" desta
// Edge Function (Supabase → Edge Functions → enviar-email-outlook →
// Secrets, ou Project Settings → Edge Functions → Secrets — ver o arquivo
// supabase-function-enviar-email-outlook.md pro passo a passo completo).
//
// Segurança: antes de mandar qualquer e-mail, a function confere (pelo
// token de login que o navegador manda sozinho) que quem está chamando é
// alguém realmente autenticado neste app — não precisa ser admin
// especificamente (qualquer vendedor(a) pode lançar um pedido e disparar o
// e-mail de aviso), só precisa estar logado de verdade.

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
}

function splitEnderecos(valor: unknown): string[] {
  if (!valor) return []
  return String(valor)
    .split(',')
    .map((e) => e.trim())
    .filter(Boolean)
}

Deno.serve(async (req: Request) => {
  // Preflight do navegador (CORS) — precisa responder OK sem checar nada.
  // IMPORTANTE: assim como a function "criar-conta", esta também precisa
  // estar com "Enforce JWT verification" DESLIGADO nas Settings dela no
  // painel do Supabase — senão o próprio Supabase bloqueia esse preflight
  // (e a chamada real) antes do código abaixo rodar, e o navegador mostra
  // só um erro genérico ("Failed to send a request to the Edge Function").
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  const jsonResponse = (body: unknown, status = 200) =>
    new Response(JSON.stringify(body), {
      status,
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
    })

  try {
    const authHeader = req.headers.get('Authorization')
    if (!authHeader) {
      return jsonResponse({ error: 'Não autenticado.' }, 401)
    }

    // Essas 2 variáveis já existem sozinhas em toda Edge Function do
    // Supabase — não precisa cadastrar nada manualmente em "Secrets" pra
    // elas (só as 4 do Azure/Outlook abaixo precisam ser cadastradas).
    const supabaseUrl = Deno.env.get('SUPABASE_URL')!
    const anonKey = Deno.env.get('SUPABASE_ANON_KEY')!

    // Confere quem está chamando — basta estar logado de verdade neste
    // app (não precisa ser admin: qualquer vendedor(a) lança pedido).
    const callerClient = createClient(supabaseUrl, anonKey, {
      global: { headers: { Authorization: authHeader } },
    })
    const { data: userData, error: userErr } = await callerClient.auth.getUser()
    if (userErr || !userData?.user) {
      return jsonResponse({ error: 'Sessão inválida — faça login de novo e tente outra vez.' }, 401)
    }

    const body = await req.json().catch(() => null)
    const assunto = body?.subject?.trim()
    const texto = body?.text
    const html = body?.html
    const destinatarios = splitEnderecos(body?.to)
    const copias = splitEnderecos(body?.cc)

    // "html" é opcional (Versão 6.3) — quando vem preenchido, o e-mail sai
    // formatado (contentType HTML); senão, continua em texto puro como
    // sempre foi. Mas precisa vir pelo menos um dos dois.
    if (!assunto || (!texto && !html) || !destinatarios.length) {
      return jsonResponse({ error: 'Preencha destinatário, assunto e o corpo do e-mail.' }, 400)
    }

    // Credenciais do app cadastrado no Azure AD (Microsoft Entra), com
    // permissão de aplicativo "Mail.Send" — cadastradas como Secrets desta
    // Edge Function, nunca no código. OUTLOOK_SENDER_EMAIL é a caixa de
    // avisos de onde o e-mail sai (precisa ter essa permissão concedida
    // pra ela, ou ser "Mail.Send" sem restrição de caixa).
    const tenantId = Deno.env.get('AZURE_TENANT_ID')
    const clientId = Deno.env.get('AZURE_CLIENT_ID')
    const clientSecret = Deno.env.get('AZURE_CLIENT_SECRET')
    const remetente = Deno.env.get('OUTLOOK_SENDER_EMAIL')
    if (!tenantId || !clientId || !clientSecret || !remetente) {
      return jsonResponse({
        error: 'Faltam credenciais do Outlook nos Secrets desta function (AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET / OUTLOOK_SENDER_EMAIL) — ver supabase-function-enviar-email-outlook.md.',
      }, 500)
    }

    // 1) Pega um token de aplicativo (client credentials) — vale por si só
    //    pra mandar e-mail pela caixa configurada, sem precisar de login de
    //    pessoa nenhuma (é exatamente o mesmo mecanismo do outro sistema da
    //    empresa que já usa essa caixa).
    const tokenResp = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials',
      }),
    })
    const tokenData = await tokenResp.json().catch(() => null)
    if (!tokenResp.ok || !tokenData?.access_token) {
      console.error('Outlook: falha ao pegar token do Azure AD.', tokenData)
      return jsonResponse({ error: 'Não consegui autenticar no Azure AD (confira Tenant ID / Client ID / Client Secret, e se a permissão Mail.Send tem consentimento de admin).' }, 502)
    }

    // 2) Manda o e-mail de verdade via Microsoft Graph, pela caixa
    //    configurada em OUTLOOK_SENDER_EMAIL.
    const graphResp = await fetch(
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(remetente)}/sendMail`,
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${tokenData.access_token}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          message: {
            subject: assunto,
            body: html
              ? { contentType: 'HTML', content: String(html) }
              : { contentType: 'Text', content: String(texto) },
            toRecipients: destinatarios.map((e) => ({ emailAddress: { address: e } })),
            ccRecipients: copias.map((e) => ({ emailAddress: { address: e } })),
          },
          saveToSentItems: true,
        }),
      },
    )

    if (graphResp.status !== 202) {
      const erroGraph = await graphResp.json().catch(() => null)
      console.error('Outlook: falha ao enviar via Microsoft Graph.', erroGraph)
      return jsonResponse({
        error: 'O Outlook recusou o envio: ' + (erroGraph?.error?.message || ('HTTP ' + graphResp.status)),
      }, 502)
    }

    return jsonResponse({ success: true })
  } catch (e) {
    return jsonResponse({ error: String((e as Error)?.message || e) }, 500)
  }
})
