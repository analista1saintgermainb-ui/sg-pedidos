const BOXLINK_API = "https://api.boxlink.com.br"

const TRACKING_PATHS = [
  "/v2/tracking/ultima-ocorrencia",
  "/v2/tracking/ocorrencias",
  "/v2/tracking/periodo",
  "/rastreamento/v2/ultima-ocorrencia",
  "/v2/rastreamento/ultima-ocorrencia",
  "/rastreamento/ultima-ocorrencia",
]

const json = (res, status, body) => res.status(status).json(body)

export default async function handler(req, res) {
  const token = process.env.BOXLINK_TOKEN || req.headers.authorization?.replace(/^Bearer\s+/i, "")
  if (!token) {
    return json(res, 401, { error: "Token Boxlink ausente" })
  }

  const { dataHoraInicio, dataHoraFim, page = "0", size = "100" } = req.query
  if (!dataHoraInicio || !dataHoraFim) {
    return json(res, 400, { error: "Informe dataHoraInicio e dataHoraFim" })
  }

  const params = new URLSearchParams({ dataHoraInicio, dataHoraFim, page, size })
  let lastError = null

  for (const path of TRACKING_PATHS) {
    const upstream = await fetch(`${BOXLINK_API}${path}?${params.toString()}`, {
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    }).catch(error => ({ ok: false, status: 0, text: async () => error.message }))

    if (upstream.status === 404) continue

    if (!upstream.ok) {
      lastError = await upstream.text()
      continue
    }

    const data = await upstream.json()
    if (Array.isArray(data)) return json(res, 200, { items: data, hasMore: false })

    const items = data.content || data.data || data.envios || data.ocorrencias || []
    const hasMore = Number(page) < (data.totalPages || 1) - 1
    return json(res, 200, { items, hasMore })
  }

  return json(res, 502, { error: lastError || "Nenhum endpoint Boxlink respondeu" })
}
