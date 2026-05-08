const TRACKING_URL =
  process.env.CORREIOS_TRACKING_URL ||
  "https://api.correios.com.br/srorastro/v1/objetos"

const json = (res, status, body) => res.status(status).json(body)
const cleanCode = value => String(value || "").trim().toUpperCase()
const asCodes = value => Array.isArray(value)
  ? value.map(cleanCode).filter(Boolean).slice(0, 50)
  : cleanCode(value) ? [cleanCode(value)] : []

const isCorreiosCarrier = value => {
  const v = String(value || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase()
  return v.includes("correio")
}

async function requireSession(req, allowedProfiles = ["admin", "logistica", "suporte"]) {
  const supabaseUrl = process.env.SUPABASE_URL
  const anonKey = process.env.SUPABASE_ANON_KEY
  const serviceKey = process.env.SUPABASE_SERVICE_ROLE_KEY
  const token = req.headers.authorization?.replace(/^Bearer\s+/i, "")

  if (!supabaseUrl || !anonKey || !serviceKey) {
    return { ok: false, status: 500, error: "Variaveis Supabase ausentes no servidor" }
  }
  if (!token) return { ok: false, status: 401, error: "Login obrigatorio" }

  const userResp = await fetch(`${supabaseUrl}/auth/v1/user`, {
    headers: { apikey: anonKey, Authorization: `Bearer ${token}` },
  })
  const user = await userResp.json().catch(() => null)
  if (!userResp.ok || !user?.id) return { ok: false, status: 401, error: "Sessao invalida" }

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
  if (!profileResp.ok || profile?.ativo === false || !allowedProfiles.includes(profile?.perfil)) {
    return { ok: false, status: 403, error: "Acesso sem permissao para Correios" }
  }

  return { ok: true, user, profile }
}

function latestEvent(obj) {
  const events = Array.isArray(obj?.eventos) ? obj.eventos : Array.isArray(obj?.eventosObjeto) ? obj.eventosObjeto : []
  return events[0] || null
}

function normalizeItem(body, fallbackCode) {
  const obj = Array.isArray(body?.objetos) ? body.objetos[0] : body?.objeto || body
  const event = latestEvent(obj)
  const unidade = event?.unidade?.endereco || event?.unidade || {}
  const cidade = unidade?.cidade || unidade?.municipio || ""
  const uf = unidade?.uf || ""
  const local = [cidade, uf].filter(Boolean).join(" / ")
  const status = event?.descricao || event?.status || event?.tipo || obj?.mensagem || ""
  const detalhe = event?.detalhe || event?.descricaoDetalhe || ""
  return {
    codigo: cleanCode(obj?.codObjeto || obj?.codigoObjeto || obj?.codigo || fallbackCode),
    status,
    detalhe,
    ultimaMov: event?.dtHrCriado || event?.dataHora || event?.data || "",
    local,
    raw: body,
  }
}

async function fetchTracking(code) {
  const url = `${TRACKING_URL.replace(/\/$/, "")}/${encodeURIComponent(code)}?resultado=T`
  const resp = await fetch(url, {
    method: "GET",
    headers: {
      Accept: "application/json",
      Authorization: `Bearer ${process.env.CORREIOS_TOKEN}`,
    },
  })
  const text = await resp.text()
  let data = {}
  try {
    data = text ? JSON.parse(text) : {}
  } catch {
    data = { message: text }
  }
  return { resp, data }
}

export default async function handler(req, res) {
  if (req.method !== "POST") return json(res, 405, { error: "Metodo nao permitido" })

  const session = await requireSession(req)
  if (!session.ok) return json(res, session.status, { error: session.error })
  if (!process.env.CORREIOS_TOKEN) return json(res, 500, { error: "CORREIOS_TOKEN ausente no servidor" })

  const body = req.body || {}
  if (body.transportadora && !isCorreiosCarrier(body.transportadora)) {
    return json(res, 400, { error: "Integracao limitada a pedidos Correios" })
  }

  const codigos = asCodes(body.codigos || body.codigo)
  if (!codigos.length) return json(res, 400, { error: "Informe codigo de rastreio Correios" })

  try {
    const items = []
    const errors = []
    for (const codigo of codigos) {
      const { resp, data } = await fetchTracking(codigo)
      if (resp.ok) items.push(normalizeItem(data, codigo))
      else errors.push({ codigo, status: resp.status, data })
    }
    return json(res, errors.length && !items.length ? 502 : 200, { items, errors })
  } catch (error) {
    return json(res, 502, { error: error.message || "Erro ao conectar Correios" })
  }
}
