const json = (res, status, body) => res.status(status).json(body)

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return json(res, 405, { error: "Metodo nao permitido" })
  }

  const supabaseUrl = process.env.SUPABASE_URL
  const anonKey = process.env.SUPABASE_ANON_KEY
  const serviceKey = process.env.SUPABASE_SERVICE_ROLE_KEY

  if (!supabaseUrl || !anonKey || !serviceKey) {
    return json(res, 500, { error: "Variaveis Supabase ausentes no servidor" })
  }

  const token = req.headers.authorization?.replace(/^Bearer\s+/i, "")
  if (!token) {
    return json(res, 401, { error: "Login obrigatorio" })
  }

  const { email, password } = req.body || {}
  if (!email || !password || password.length < 6) {
    return json(res, 400, { error: "Informe email e senha com pelo menos 6 caracteres" })
  }

  const userResp = await fetch(`${supabaseUrl}/auth/v1/user`, {
    headers: { apikey: anonKey, Authorization: `Bearer ${token}` },
  })
  const user = await userResp.json().catch(() => null)

  if (!userResp.ok || !user?.id) {
    return json(res, 401, { error: "Sessao invalida" })
  }

  const serviceHeaders = {
    apikey: serviceKey,
    Authorization: `Bearer ${serviceKey}`,
    "Content-Type": "application/json",
  }

  const profileResp = await fetch(
    `${supabaseUrl}/rest/v1/usuarios?id=eq.${encodeURIComponent(user.id)}&select=perfil,ativo&limit=1`,
    { headers: serviceHeaders }
  )
  const profiles = await profileResp.json().catch(() => [])
  const profile = Array.isArray(profiles) ? profiles[0] : null

  if (!profileResp.ok || profile?.perfil !== "admin" || profile?.ativo === false) {
    return json(res, 403, { error: "Apenas administradores podem criar usuarios" })
  }

  const createResp = await fetch(`${supabaseUrl}/auth/v1/admin/users`, {
    method: "POST",
    headers: serviceHeaders,
    body: JSON.stringify({ email, password, email_confirm: true }),
  })
  const created = await createResp.json().catch(() => null)

  if (!createResp.ok) {
    return json(res, createResp.status, {
      error: created?.msg || created?.message || "Erro ao criar usuario",
    })
  }

  return json(res, 200, { id: created.id, email: created.email })
}
