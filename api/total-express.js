const TRACKING_URL =
  process.env.TOTAL_EXPRESS_TRACKING_URL ||
  "https://apis.totalexpress.com.br/ics-tracking-encomenda-lv/v1/tracking"
const TICKET_URL =
  process.env.TOTAL_EXPRESS_TICKET_URL ||
  "https://apis.totalexpress.com.br/ics-ticket-lv/v1/ticket"

const json = (res, status, body) => res.status(status).json(body)
const asArray = value => Array.isArray(value) ? value.filter(Boolean).map(String).slice(0, 50) : []
const isTotalExpressCarrier = value => {
  const v = String(value || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase()
  return v.includes("total") || v.includes("tex")
}
const toInt = value => {
  if (value === undefined || value === null || value === "") return undefined
  const n = Number(value)
  return Number.isFinite(n) ? n : undefined
}

function totalHeaders() {
  const headers = { "Content-Type": "application/json" }
  const customHeader = process.env.TOTAL_EXPRESS_AUTH_HEADER
  const customValue = process.env.TOTAL_EXPRESS_AUTH_VALUE
  if (customHeader && customValue) {
    headers[customHeader] = customValue
    return headers
  }

  const token = process.env.TOTAL_EXPRESS_TOKEN
  if (token) {
    const scheme = process.env.TOTAL_EXPRESS_AUTH_SCHEME || "Bearer"
    headers.Authorization = scheme ? `${scheme} ${token}` : token
  }
  return headers
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
    return { ok: false, status: 403, error: "Acesso sem permissao para Total Express" }
  }

  return { ok: true, user, profile }
}

function normalizeTrackingResponse(body) {
  const page = body?.data?.data && Array.isArray(body.data.data) ? body.data : body
  const items = Array.isArray(page?.data) ? page.data : []
  return {
    items,
    currentPage: page?.currentPage ?? page?.current_page ?? 1,
    lastPage: page?.lastPage ?? page?.last_page ?? 1,
    perPage: page?.perPage ?? 50,
    total: page?.total ?? items.length,
    raw: body,
  }
}

function ticketConfig(body) {
  const required = {
    criacaoFuncionarioId: toInt(process.env.TOTAL_EXPRESS_TICKET_CRIACAO_FUNCIONARIO_ID),
    categoriaId: toInt(body.categoriaId ?? process.env.TOTAL_EXPRESS_TICKET_CATEGORIA_ID),
    assuntoId: toInt(body.assuntoId ?? process.env.TOTAL_EXPRESS_TICKET_ASSUNTO_ID),
    responsavelGrupoNivelId: toInt(process.env.TOTAL_EXPRESS_TICKET_RESPONSAVEL_GRUPO_NIVEL_ID),
    criacaoGrupoId: toInt(process.env.TOTAL_EXPRESS_TICKET_CRIACAO_GRUPO_ID),
    responsavelGrupoId: toInt(process.env.TOTAL_EXPRESS_TICKET_RESPONSAVEL_GRUPO_ID),
    privacidadeId: toInt(process.env.TOTAL_EXPRESS_TICKET_PRIVACIDADE_ID),
    remetenteId: toInt(process.env.TOTAL_EXPRESS_TICKET_REMETENTE_ID),
  }
  const missing = Object.entries(required).filter(([, value]) => value === undefined).map(([key]) => key)
  return {
    missing,
    payload: {
      ...required,
      responsavelFuncionarioId: toInt(body.responsavelFuncionarioId ?? process.env.TOTAL_EXPRESS_TICKET_RESPONSAVEL_FUNCIONARIO_ID) ?? 0,
      descricao: body.descricao || "Acionamento aberto pela Central de Pedidos Saint Germain",
      origem: body.origem || "encomenda",
      origemId: String(body.origemId || ""),
      casoCritico: body.casoCritico ? 1 : 0,
      dataAgendamento: body.dataAgendamento || "",
      horaAgendamento: body.horaAgendamento || "",
      novaTentativa: body.novaTentativa || "",
      periodo1: body.periodo1 || "",
      periodo2: body.periodo2 || "",
    },
  }
}

async function forwardJson(url, payload) {
  const resp = await fetch(url, {
    method: "POST",
    headers: totalHeaders(),
    body: JSON.stringify(payload),
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

  if (!process.env.TOTAL_EXPRESS_TOKEN && !process.env.TOTAL_EXPRESS_AUTH_VALUE) {
    return json(res, 500, { error: "Credencial Total Express ausente no servidor" })
  }

  const body = req.body || {}
  const action = body.action || "tracking"
  if (body.transportadora && !isTotalExpressCarrier(body.transportadora)) {
    return json(res, 400, { error: "Integracao limitada a pedidos Total Express" })
  }

  try {
    if (action === "tracking") {
      const payload = {
        awbs: asArray(body.awbs),
        notasFiscais: asArray(body.notasFiscais),
        pedidos: asArray(body.pedidos),
        comprovanteEntrega: Boolean(body.comprovanteEntrega),
      }
      if (body.ediRecebimentoInicioData) payload.ediRecebimentoInicioData = body.ediRecebimentoInicioData
      if (body.ediRecebimentoFimData) payload.ediRecebimentoFimData = body.ediRecebimentoFimData
      if (!payload.awbs.length && !payload.notasFiscais.length && !payload.pedidos.length && !payload.ediRecebimentoInicioData) {
        return json(res, 400, { error: "Informe awbs, notasFiscais, pedidos ou ediRecebimentoInicioData" })
      }

      const page = Math.max(1, toInt(body.page) || 1)
      const { resp, data } = await forwardJson(`${TRACKING_URL}?page=${page}`, payload)
      return json(res, resp.status, resp.ok ? normalizeTrackingResponse(data) : data)
    }

    if (action === "ticket") {
      const cfg = ticketConfig(body)
      if (!cfg.payload.origemId) return json(res, 400, { error: "Informe origemId para abrir ticket" })
      if (cfg.missing.length) {
        return json(res, 500, { error: "Configuracao de ticket Total Express incompleta", missing: cfg.missing })
      }
      const { resp, data } = await forwardJson(TICKET_URL, cfg.payload)
      return json(res, resp.status, resp.ok ? { data } : data)
    }

    return json(res, 400, { error: "Acao Total Express invalida" })
  } catch (error) {
    return json(res, 502, { error: error.message || "Erro ao conectar Total Express" })
  }
}
