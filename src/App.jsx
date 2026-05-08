import { useState, useEffect, useRef, useCallback } from 'react'
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell } from 'recharts'
import * as XLSX from 'xlsx'

// ─── Design System ────────────────────────────────────────────
// ─── Temas claro e escuro ─────────────────────────────────────
const CL = {
  brand:      "#050505", brandSoft:  "#141414",
  gold:       "#111111", goldLight:  "#2A2A2A", goldDim: "#5E5E5E",
  cream:      "#F6F5F1", creamDark:  "#ECEAE4",
  white:      "#FFFFFF", border:     "#DEDAD0", borderDark: "#C7C2B7",
  text1:      "#111111", text2:      "#4E4A44", text3: "#858078", text4: "#B8B2A8",
  red:        "#C0392B", redSoft:    "#F9ECEB", redBorder:  "#EBCBC8",
  green:      "#2E7D50", greenSoft:  "#EAF4EE", greenBorder:"#C0DCCB",
  amber:      "#595959", amberSoft:  "#F1F1EF", amberBorder:"#D8D8D2",
  blue:       "#1A5276", blueSoft:   "#EAF2FB", blueBorder: "#AACDE6",
}
const CD = {
  brand:      "#F7F6F1", brandSoft:  "#E3DFD4",
  gold:       "#F7F6F1", goldLight:  "#FFFFFF", goldDim: "#BDBDBD",
  cream:      "#080808", creamDark:  "#111111",
  white:      "#181818", border:     "#2A2A2A", borderDark: "#3A3A3A",
  text1:      "#F7F6F1", text2:      "#C9C3B8", text3: "#8A8378", text4: "#5E584F",
  red:        "#E05555", redSoft:    "#2A1212", redBorder:  "#4A2020",
  green:      "#4AB870", greenSoft:  "#0A2015", greenBorder:"#1A4030",
  amber:      "#D6D1C8", amberSoft:  "#1C1C1C", amberBorder:"#3A3A3A",
  blue:       "#5A9AD4", blueSoft:   "#0A1825", blueBorder: "#1A3A55",
}
// C é mutável — applyTheme() troca os valores em re-render
const C = {...CL}
function applyTheme(dark) {
  const src = dark ? CD : CL
  Object.keys(src).forEach(k => { C[k] = src[k] })
}
const shadow = {
  sm: "0 1px 0 rgba(17,17,17,0.06)",
  md: "0 10px 24px rgba(17,17,17,0.06)",
  lg: "0 18px 42px rgba(17,17,17,0.10)",
}
const getGlobalStyle = () => `
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: ${C.cream}; color: ${C.text1}; font-family: 'Inter', Arial, sans-serif; }
  ::-webkit-scrollbar { width: 4px; height: 4px; }
  ::-webkit-scrollbar-track { background: ${C.creamDark}; }
  ::-webkit-scrollbar-thumb { background: ${C.borderDark}; border-radius: 2px; }
  ::-webkit-scrollbar-thumb:hover { background: ${C.gold}; }
  select, input, textarea, button { font-family: 'Inter', Arial, sans-serif; }
  button { -webkit-font-smoothing: antialiased; }
  tr:hover td { background: ${C.creamDark} !important; }
`

// ─── Supabase ─────────────────────────────────────────────────
const SUPA_URL     = "https://jdiuuhfhsiymttxllssr.supabase.co"
const SUPA_KEY     = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpkaXV1aGZoc2l5bXR0eGxsc3NyIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzc3NTMyNTcsImV4cCI6MjA5MzMyOTI1N30.wNGhwh2bCF0HZSonn09S-15kEVAQGzEP1yWvRx3l_N4"
const SH  = { apikey: SUPA_KEY, "Content-Type": "application/json" }
const aSH = t => ({ ...SH, Authorization: `Bearer ${t}` })
const LOGIN_EMAIL_DOMAIN = "sg-pedidos.local"
const SESSION_KEY = "sg_pedidos_session"
const cleanLogin = value => String(value||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().trim().replace(/[^a-z0-9._-]/g,"")
const authEmailFromLogin = value => {
  const raw = String(value||"").trim().toLowerCase()
  if (raw.includes("@")) return raw
  const login = cleanLogin(raw)
  return login ? `${login}@${LOGIN_EMAIL_DOMAIN}` : ""
}
const displayLogin = email => {
  const raw = String(email||"")
  return raw.endsWith(`@${LOGIN_EMAIL_DOMAIN}`) ? raw.replace(`@${LOGIN_EMAIL_DOMAIN}`,"") : raw
}

async function signIn(login, password) {
  const email = authEmailFromLogin(login)
  const r = await fetch(`${SUPA_URL}/auth/v1/token?grant_type=password`, { method:"POST", headers:SH, body:JSON.stringify({email,password}) })
  const d = await r.json(); if (!r.ok) throw new Error(d.error_description||d.msg||"Erro ao fazer login"); return d
}
async function signOut(token) { await fetch(`${SUPA_URL}/auth/v1/logout`,{method:"POST",headers:aSH(token)}) }
async function createUser(email, password, token) {
  const r = await fetch("/api/create-user", {method:"POST",headers:{"Content-Type":"application/json",Authorization:`Bearer ${token}`},body:JSON.stringify({email,password})})
  const d = await r.json(); if (!r.ok) throw new Error(d.msg||d.message||"Erro ao criar usuário"); return d
}
async function totalExpressRequest(payload, token) {
  const r = await fetch("/api/total-express", {method:"POST",headers:{"Content-Type":"application/json",Authorization:`Bearer ${token}`},body:JSON.stringify(payload)})
  const d = await r.json().catch(()=>({}))
  if (!r.ok) throw new Error(d.error||d.message||d.descricao||"Erro Total Express")
  return d
}
async function dbLoadFast(token, onPartial) {
  let all = [], from = 0, step = 1000
  while (true) {
    const r = await fetch(`${SUPA_URL}/rest/v1/pedidos?select=*&order=id&limit=${step}&offset=${from}`, { headers: aSH(token) })
    if (!r.ok) throw new Error(await r.text())
    const data = await r.json()
    const chunk = data.map(row => ({ ...row.dados, id: row.id }))
    all = [...all, ...chunk]
    if (from === 0 && chunk.length > 0) onPartial(all)
    if (data.length < step) break
    from += step
  }
  return all
}
async function dbUpsert(rows, token) {
  for (let i = 0; i < rows.length; i += 200) {
    const batch = rows.slice(i, i+200)
    const payload = batch.map(r => ({id:Number(r.id),dados:r,updated_at:new Date().toISOString()}))
    const r2 = await fetch(`${SUPA_URL}/rest/v1/pedidos`,{method:"POST",headers:{...aSH(token),Prefer:"resolution=merge-duplicates,return=minimal"},body:JSON.stringify(payload)})
    if (!r2.ok) throw new Error(`${r2.status}: ${await r2.text()}`)
  }
}
async function dbDelete(id, token){ await fetch(`${SUPA_URL}/rest/v1/pedidos?id=eq.${id}`,{method:"DELETE",headers:aSH(token)}) }
async function dbClear(token){ await fetch(`${SUPA_URL}/rest/v1/pedidos?id=gte.0`,{method:"DELETE",headers:aSH(token)}) }
async function loadUsuarios(token){ const r = await fetch(`${SUPA_URL}/rest/v1/usuarios?select=*&order=created_at`,{headers:aSH(token)}); return r.ok?r.json():[] }
async function saveUsuario(u, token){ await fetch(`${SUPA_URL}/rest/v1/usuarios`,{method:"POST",headers:{...aSH(token),Prefer:"resolution=merge-duplicates,return=minimal"},body:JSON.stringify(u)}) }
async function deleteUsuario(id, token){ await fetch(`${SUPA_URL}/rest/v1/usuarios?id=eq.${id}`,{method:"DELETE",headers:aSH(token)}) }

// ─── Permissões ───────────────────────────────────────────────
const PERMS = {
  admin:    {tabs:["dashboard","logistica","suporte","devolucao","reenvio","arquivados","usuarios"],canImport:true,canDelete:true,canClear:true,canSendSupport:true,canOperate:true},
  logistica:{tabs:["dashboard","logistica"],canImport:true,canDelete:false,canClear:false,canSendSupport:true,canOperate:true},
  suporte:  {tabs:["suporte","devolucao","reenvio","arquivados"],canImport:false,canDelete:false,canClear:false,canSendSupport:false,canOperate:true},
  leitura:  {tabs:["dashboard","logistica","suporte","devolucao","reenvio","arquivados"],canImport:false,canDelete:false,canClear:false,canSendSupport:false,canOperate:false},
}

// ─── BUG FIX #1: ALERTA_DIAS movido para ANTES de QFILTERS ───
const ALERTA_DIAS = 7

// ─── Mapeamento de colunas expandido ─────────────────────────
const HEADER_MAP = {
  nuvem:        ["identificador ecommerce","id ecommerce","no nuvem","nuvem","pedido"],
  destinatario: [
    "destinatário nome","destinatario nome","nome destinatário","nome destinatario",
    "nome do destinatário","nome do destinatario","nome do cliente","nome do comprador",
    "nome comprador","comprador","destinatário","destinatario","nome do pedido","nome",
  ],
  transportadora:["estratégia de frete","estrategia de frete","transportadora","frete"],
  rastreio:     ["rastreador last mile","código de rastreio","codigo de rastreio","rastreio","last mile"],
  status:       ["situação","situacao","situac","status"],
  prazo:        ["prazo logístico","prazo logistico","prazo logico","prazo entrega","prazo de entrega","prazo previsto","prazo"],
  nf:           ["nº nota fiscal","no nota fiscal","nota fiscal","no nf","nf"],
  ultimaMov:    ["última movimentação","ultima movimentacao","ultima movimentação","última movimentacao","ultima mov","última mov","movimentacao","movimentação","data última ocorrência","data ultima ocorrencia","data ultima ocorrência","data última ocorrencia","ultima ocorrencia","última ocorrência"],
  cidade:       ["destinatário cidade","destinatario cidade","cidade destinatário","cidade destinatario","cidade"],
  uf:           ["destinatário uf","destinatario uf","uf destinatário","uf destinatario","uf","estado"],
  cep:          ["destinatário cep","destinatario cep","cep destinatário","cep destinatario","cep"],
  statusPrazo:  ["status prazo","statusprazo","prazo status","situacao prazo","situação prazo"],
  dataCriacao:  ["data criação envio","data criacao envio","data de criacao","data criacao","data envio"],
  email:        ["destinatário e-mail","destinatario e-mail","e-mail destinatário","email destinatario","e-mail do cliente","email do cliente","email cliente","e-mail","email"],
}
const norm    = s => (s||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().trim()
const findIdx = (hdrs, key) => {
  const aliases = HEADER_MAP[key] || []
  const nhdrs   = hdrs.map(norm)
  // 1ª — exact match em qualquer alias
  for (const v of aliases) {
    const i = nhdrs.findIndex(h => h === norm(v))
    if (i >= 0) return i
  }
  // 2ª — alias multi-palavra: cabeçalho contém a frase completa
  for (const v of aliases.filter(v => norm(v).includes(" "))) {
    const nv = norm(v)
    const i  = nhdrs.findIndex(h => h.includes(nv))
    if (i >= 0) return i
  }
  // 3ª — alias simples: cabeçalho COMEÇA COM o alias (evita "status prazo" casar com "prazo")
  for (const v of aliases.filter(v => !norm(v).includes(" "))) {
    const nv = norm(v)
    const i  = nhdrs.findIndex(h => h.startsWith(nv + " ") || h.startsWith(nv + "_") || h === nv)
    if (i >= 0) return i
  }
  // 4ª — último recurso: contains simples
  for (const v of aliases.filter(v => !norm(v).includes(" "))) {
    const nv = norm(v)
    const i  = nhdrs.findIndex(h => h.includes(nv))
    if (i >= 0) return i
  }
  return -1
}
const uniq   = arr => ["Todos",...Array.from(new Set(arr.filter(Boolean).sort()))]

// BUG FIX #1 (cont): QFILTERS agora pode referenciar ALERTA_DIAS sem erro
const QFILTERS = [
  {id:"todos",      label:"Todos"},
  {id:"urgente",    label:"Urgente"},
  {id:"extraviados",label:"Extraviados"},
  {id:"devolvidos", label:"Devolvidos"},
  {id:"vence_hoje", label:"Vence hoje"},
  {id:"vencidos",   label:"Vencidos"},
  {id:"parados",    label:`Parados +${ALERTA_DIAS}d`},
]
const PAGE_SIZE = 200

// ─── Helpers de cálculo ───────────────────────────────────────

// BUG FIX #4: parseStatusPrazo estava sendo chamada mas nunca definida
function parseStatusPrazo(raw) {
  if (!raw) return null
  const v = (raw||"").toLowerCase().trim()
  if (v.includes("antes")||v.includes("no prazo")||v==="ok"||v.includes("dentro")||v.includes("normal")) return true
  if (v.includes("atras")||v.includes("fora")||v==="vencido"||v.includes("atraso")) return false
  return null
}

// BUG FIX #2: isEntregue estava sem declaração function
function isEntregue(status) {
  const s = (status||"").toLowerCase()
  return s.includes("entregue")||s.includes("finaliz")||s.includes("entrega realizada")
}

function diasSemMov(ultimaMov) {
  if (!ultimaMov) return null
  const d = parsePrazo(ultimaMov)
  if (!d) return null
  const hoje = new Date(); hoje.setHours(0,0,0,0)
  return Math.floor((hoje - d) / 86400000)
}

function semMovInfo(ultimaMov) {
  const dias = diasSemMov(ultimaMov)
  if (dias === null) return null
  if (dias >= ALERTA_DIAS) return { dias, label:`${dias}d sem atualização`, alerta:true }
  if (dias >= 3)            return { dias, label:`${dias}d sem atualização`, alerta:false }
  return null
}

function calcMotivo(s) {
  const v = (s||"").toLowerCase()
  if (v.includes("extravia")||v.includes("perdid"))          return "Objeto extraviado"
  if (v.includes("devolv")||v.includes("recusa"))            return "Devolução / Recusa"
  if (v.includes("falha")||v.includes("tentativa"))          return "Falha na entrega"
  if (v.includes("atras"))                                   return "Atraso na entrega"
  if (v.includes("entregue")||v.includes("finaliz"))         return "Entrega concluída"
  if (v.includes("saiu"))                                    return "Saiu para entrega"
  if (v.includes("trânsito")||v.includes("transito"))        return "Em trânsito"
  if (v.includes("postado")||v.includes("coletado"))         return "Aguardando movimentação"
  if (v.includes("pendente")||v.includes("aguardando")||v.includes("processando")) return "Aguardando coleta"
  return "—"
}

function parsePrazo(v) {
  if (!v) return null
  const s = String(v).trim()
  // DD/MM/YYYY ou D/M/YYYY (padrão brasileiro)
  const br = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/)
  if (br) return new Date(+br[3], +br[2]-1, +br[1])
  // YYYY-MM-DD (ISO)
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/)
  if (iso) return new Date(+iso[1], +iso[2]-1, +iso[3])
  // DD-MM-YYYY
  const dmy = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/)
  if (dmy) return new Date(+dmy[3], +dmy[2]-1, +dmy[1])
  // DD.MM.YYYY
  const dot = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/)
  if (dot) return new Date(+dot[3], +dot[2]-1, +dot[1])
  // Excel serial (número inteiro como 45678)
  const num = parseFloat(s)
  if (!isNaN(num) && num > 30000 && num < 100000)
    return new Date(Math.round((num - 25569) * 86400000))
  return null
}

function calcUrg(prazo, status) {
  const s = (status||"").toLowerCase()
  // ── Alta: situações críticas ──────────────────────────────
  if (s.includes("extravia")||s.includes("perdid"))                          return "Alta"
  if (s.includes("devolv")||s.includes("recusa"))                            return "Alta"
  if (s.includes("problema_entrega")||s.includes("problema entrega"))        return "Alta"
  if (s.includes("falha")||s.includes("retido")||s.includes("apreend"))     return "Alta"
  // ── Baixa: entregue ou saindo para entrega ────────────────
  if (s.includes("entregue")||s.includes("finaliz"))                         return "Baixa"
  if (s.includes("saiu_para_entrega")||s.includes("saiu para entrega")||
      s.includes("saida_para_entrega")||s.includes("saiu"))                  return "Baixa"
  // ── Média: etapas intermediárias ──────────────────────────
  if (s.includes("aguardando_retirada")||s.includes("aguardando retirada"))  return "Média"
  if (s.includes("triado")||s.includes("triagem"))                           return "Média"
  if (s.includes("em_transito")||s.includes("em transito")||
      s.includes("trânsito")||s.includes("transito"))                        return "Média"
  if (s.includes("postado")||s.includes("coletado")||s.includes("colet"))   return "Média"
  if (s.includes("aguardando")||s.includes("processando"))                   return "Média"
  // ── Fallback: calcular pelo prazo logístico ───────────────
  const dt = parsePrazo(prazo)
  if (!dt) return "Média" // nunca retorna "—" — sem prazo = Média por padrão
  const h = new Date(); h.setHours(0,0,0,0)
  const d = Math.ceil((dt-h)/86400000)
  if (d<=1) return "Alta"; if (d<=3) return "Média"; return "Baixa"
}

function calcAcionar(urg, status) {
  const s = (status||"").toLowerCase()
  if (urg==="Alta"||s.includes("extravia")||s.includes("devolv")) return "Sim"
  if (urg==="Média") return "Avaliar"
  return "Não"
}

function slaInfo(prazo) {
  const dt = parsePrazo(prazo); if (!dt) return null
  const h = new Date(); h.setHours(0,0,0,0)
  const diff = Math.ceil((dt-h)/86400000)
  if (diff < 0)   return {label:`${Math.abs(diff)}d vencido`, color:C.red,   bg:C.redSoft,   bd:C.redBorder}
  if (diff === 0) return {label:"Vence hoje",                 color:C.amber, bg:C.amberSoft, bd:C.amberBorder}
  if (diff <= 3)  return {label:`${diff}d restantes`,         color:C.amber, bg:C.amberSoft, bd:C.amberBorder}
  return              {label:`${diff}d restantes`,             color:C.green, bg:C.greenSoft, bd:C.greenBorder}
}

function timeOpen(sentAt) {
  if (!sentAt) return null
  const ms = Date.now()-new Date(sentAt).getTime()
  const h = Math.floor(ms/3600000), d = Math.floor(h/24)
  if (d>0) return {label:`${d}d na fila`,alert:d>=2}
  if (h>0) return {label:`${h}h na fila`,alert:false}
  return {label:"< 1h",alert:false}
}

// BUG FIX #5 e #8: parseData corrigido — ix expandido + closing correto
function shouldAutoSendSupport(status) {
  const s = norm(status)
  return s.includes("extravia") ||
    s.includes("devolv") ||
    s.includes("devolucao") ||
    s.includes("recusa") ||
    (s.includes("aguardando") && s.includes("retirada"))
}

function applyAutoSupport(row, ts=new Date().toLocaleString("pt-BR"), sentAt=new Date().toISOString()) {
  if (!row || row.enviadoSuporte || row.atendimento==="Resolvido" || isEntregue(row.status) || !shouldAutoSendSupport(row.status)) return row
  return {
    ...row,
    enviadoSuporte: true,
    atendimento: "Aberto",
    sentAt,
    historico: [...(row.historico||[]), {acao:"Enviado automaticamente ao suporte por status critico", ts}],
  }
}

function parseData(text) {
  const sep = text.includes("\t")?"\t":text.includes(";")?";":","
  const lines = text.trim().split("\n").filter(l=>l.trim())
  if (!lines.length) return []
  const first = lines[0].split(sep).map(h=>h.trim().replace(/^["']|["']$/g,""))
  const isHdr = first.some(h=>["nuvem","destinat","identificador","ecommerce","rastreio","situac","status","frete"].some(k=>norm(h).includes(k)))
  const hdrs = isHdr ? first : []
  const data = isHdr ? lines.slice(1) : lines
  const ix = {
    nuvem:       isHdr?findIdx(hdrs,"nuvem")         :0,
    dest:        isHdr?findIdx(hdrs,"destinatario")  :1,
    transp:      isHdr?findIdx(hdrs,"transportadora"):2,
    rastreio:    isHdr?findIdx(hdrs,"rastreio")      :3,
    status:      isHdr?findIdx(hdrs,"status")        :4,
    prazo:       isHdr?findIdx(hdrs,"prazo")         :5,
    nf:          isHdr?findIdx(hdrs,"nf")            :6,
    ultimaMov:   isHdr?findIdx(hdrs,"ultimaMov")     :-1,
    // BUG FIX #8: campos novos mapeados
    cidade:      isHdr?findIdx(hdrs,"cidade")        :-1,
    uf:          isHdr?findIdx(hdrs,"uf")            :-1,
    cep:         isHdr?findIdx(hdrs,"cep")           :-1,
    statusPrazo: isHdr?findIdx(hdrs,"statusPrazo")   :-1,
    dataCriacao: isHdr?findIdx(hdrs,"dataCriacao")   :-1,
    email:       isHdr?findIdx(hdrs,"email")         :-1,
  }
  const g = (c,i) => i>=0&&i<c.length?c[i]:""

  // Fallback: detecta coluna de email pelo símbolo @ nas primeiras linhas de dados
  if (ix.email < 0 && data.length > 0) {
    const firstRows = data.slice(0, Math.min(10, data.length))
    outer: for (let col = 0; col < 50; col++) {
      for (const line of firstRows) {
        const cells = line.split(sep).map(v=>v.trim().replace(/^["']|["']$/g,""))
        const val = (cells[col]||"")
        if (val.includes("@") && val.includes(".")) {
          ix.email = col
          break outer
        }
      }
    }
  }
  return data.map((line,i) => {
    const c = line.split(sep).map(v=>v.trim().replace(/^["']|["']$/g,""))
    const status = g(c,ix.status), prazo = g(c,ix.prazo)
    const urg = calcUrg(prazo,status)
    const entregue = isEntregue(status)
    const spRaw = g(c,ix.statusPrazo)
    const spVal = parseStatusPrazo(spRaw)
    const dt = parsePrazo(prazo) // BUG FIX #8: dt declarado antes de uso em noPrazo
    // metrics.entrega_no_prazo: comparar apenas DATA (sem hora)
    // data_entrega = Data Última Ocorrência (só a data), prazo = Prazo Logístico (só a data)
    const dtEntregaRaw = g(c,ix.ultimaMov)
    const dtEntrega = isEntregue(status) && dtEntregaRaw ? parsePrazo(dtEntregaRaw) : null
    const dtEntregaDate = dtEntrega ? new Date(dtEntrega.getFullYear(), dtEntrega.getMonth(), dtEntrega.getDate()) : null
    const dtPrazoDate   = dt         ? new Date(dt.getFullYear(),       dt.getMonth(),       dt.getDate())        : null
    const noPrazo = spVal!==null ? spVal
      : (dtEntregaDate && dtPrazoDate) ? dtEntregaDate <= dtPrazoDate
      : null
    const row = {
      id: Date.now()+i,
      nuvem: g(c,ix.nuvem), destinatario: g(c,ix.dest),
      transportadora: g(c,ix.transp), rastreio: g(c,ix.rastreio),
      status, prazo, nf: g(c,ix.nf),
      ultimaMov:    g(c,ix.ultimaMov),
      cidade:       g(c,ix.cidade),
      uf:           g(c,ix.uf),
      cep:          g(c,ix.cep),
      statusPrazoRaw: spRaw,
      dataCriacao:  g(c,ix.dataCriacao),
      email:        g(c,ix.email),
      motivo:       calcMotivo(status),
      urgencia:     urg,
      acionar:      calcAcionar(urg,status),
      enviadoSuporte:false,
      atendimento:  entregue?"Resolvido":"Aberto",
      entregueNoPrazo: noPrazo,
      alertaStatus: null,
      fluxoEspecial: "",
      devolucaoStatus: "Aguardando tratativa",
      reenvioStatus: "Pendente",
      decisaoCliente: "",
      motivoDevolucao: "",
      novoPedido: "",
      novaTransportadora: "",
      novoRastreio: "",
      materialReenvio: "",
      obs:"", historico: entregue?[{acao:"Arquivado automaticamente — entrega concluída",ts:new Date().toLocaleString("pt-BR")}]:[],
      responsavel:"", sentAt:null, chamado:"", isNew:true,
    }
    return applyAutoSupport(row)
  }).filter(r=>{
    if (!r.nuvem&&!r.destinatario&&!r.nf) return false
    // transportadoras.ignore: "SP" (sigla de estado, não transportadora)
    const tr = (r.transportadora||"").trim()
    if (tr==="SP"||tr==="RJ"||tr==="MG"||tr==="RS"||tr.length<=2) r.transportadora=""
    return true
  }) // BUG FIX #5: filter movido para dentro da função
}

// BUG FIX (applyQF): adicionado case "parados" que estava faltando
function applyQF(rows, qf) {
  if (qf==="todos")       return rows
  if (qf==="urgente")     return rows.filter(r=>r.urgencia==="Alta")
  if (qf==="extraviados") return rows.filter(r=>(r.status||"").toLowerCase().includes("extravia"))
  if (qf==="devolvidos")  return rows.filter(r=>(r.status||"").toLowerCase().includes("devolv"))
  if (qf==="vence_hoje")  return rows.filter(r=>{const d=parsePrazo(r.prazo);if(!d)return false;const h=new Date();h.setHours(0,0,0,0);return Math.ceil((d-h)/86400000)===0})
  if (qf==="vencidos")    return rows.filter(r=>{const d=parsePrazo(r.prazo);if(!d)return false;const h=new Date();h.setHours(0,0,0,0);return d<h&&!isEntregue(r.status)})
  if (qf==="parados")     return rows.filter(r=>{const d=diasSemMov(r.ultimaMov);return d!==null&&d>=ALERTA_DIAS})
  return rows
}

function prioridadeOperacional(r) {
  const status = (r.status||"").toLowerCase()
  const semMov = diasSemMov(r.ultimaMov)
  const prazo = parsePrazo(r.prazo)
  const hoje = new Date(); hoje.setHours(0,0,0,0)
  const vencido = prazo && prazo < hoje && !isEntregue(r.status)
  const critica = r.urgencia==="Alta" || r.acionar==="Sim" || vencido || status.includes("extravia") || status.includes("devolv") || status.includes("recusa") || (semMov!==null&&semMov>=ALERTA_DIAS)
  if (!critica) return {level:"normal", label:"Monitorar", color:C.text3, bg:C.white, bd:C.border, left:C.borderDark}
  if (status.includes("extravia")) return {level:"critica", label:"Extravio", color:C.red, bg:C.redSoft, bd:C.redBorder, left:C.red}
  if (vencido) return {level:"critica", label:"Prazo vencido", color:C.red, bg:C.redSoft, bd:C.redBorder, left:C.red}
  if (status.includes("devolv")||status.includes("recusa")) return {level:"alta", label:"Devolucao", color:C.amber, bg:C.amberSoft, bd:C.amberBorder, left:C.amber}
  if (semMov!==null&&semMov>=ALERTA_DIAS) return {level:"alta", label:`Parado ${semMov}d`, color:C.amber, bg:C.amberSoft, bd:C.amberBorder, left:C.amber}
  return {level:"alta", label:"Prioridade alta", color:C.red, bg:C.redSoft, bd:C.redBorder, left:C.red}
}

// BUG FIX #6: applySortRows — removido trailing ", [rows,...])" corrompido
function applySortRows(rows, col, dir) {
  if (!col) return rows
  return [...rows].sort((a,b)=>{
    let va=a[col]||"", vb=b[col]||""
    if (col==="prazo"){va=parsePrazo(va)||new Date(0);vb=parsePrazo(vb)||new Date(0)}
    if (col==="urgencia"){const o={Alta:0,Média:1,Baixa:2,"—":3};va=o[va]??9;vb=o[vb]??9}
    const cmp = typeof va==="object"?va-vb:String(va).localeCompare(String(vb),"pt-BR")
    return dir==="asc"?cmp:-cmp
  })
}

function exportCSV(rows) {
  const h = ["No NUVEM","Destinatário","Cidade","UF","CEP","Transportadora","Cód. Rastreio","Status","Prazo","Motivo","Urgência","Acionar?","Suporte","Atendimento","Chamado","Responsável","Observações"]
  const e = v => `"${String(v||"").replace(/"/g,'""')}"`
  const csv = [h.map(e).join(";"),...rows.map(r=>[r.nuvem,r.destinatario,r.cidade,r.uf,r.cep,r.transportadora,r.rastreio,r.status,r.prazo,r.motivo,r.urgencia,r.acionar,r.enviadoSuporte?"Sim":"Não",r.atendimento,r.chamado,r.responsavel,r.obs].map(e).join(";"))].join("\n")
  const a = document.createElement("a")
  a.href = URL.createObjectURL(new Blob(["\uFEFF"+csv],{type:"text/csv;charset=utf-8;"}))
  a.download = "saint_germain_pedidos.csv"; a.click()
}

function getTemplate(r, ch, nomeAtendente) {
  const nome  = (r.destinatario||"").split(" ")[0]||"Cliente"
  const atend = (nomeAtendente||"").split(" ")[0]||"Time SG"
  const m     = (r.motivo||"").toLowerCase()
  const extrav = m.includes("extravia")
  const devolv = m.includes("devolu")||m.includes("recusa")
  const atraso = m.includes("atraso")
  const parado = diasSemMov(r.ultimaMov)!==null&&diasSemMov(r.ultimaMov)>=ALERTA_DIAS
  const tr = (r.transportadora||"").toLowerCase()
  const linkRastreio = tr.includes("correio")      ? "https://rastreamento.correios.com.br/app/index.php"
    : tr.includes("loggi")                          ? "https://www.loggi.com/rastreador/"
    : tr.includes("total")                          ? "https://totalconecta.totalexpress.com.br/rastreamento"
    : null // J&T, Estoca, Mandaê — sem link

  if (ch==="wpp") {
    if (extrav) return `Olá, ${nome}! Tudo bem? Me chamo ${atend} e faço parte do time de encantamento da SG 🤍\n\nEstou entrando em contato pois identificamos uma ocorrência no seu pedido *#${r.nuvem}*.\n\nSua encomenda está com o status de *objeto extraviado* junto à transportadora *${r.transportadora}*. Já acionamos nossa equipe para apurar o caso com urgência.\n\nRetornaremos com uma atualização em até *2 dias úteis*. Pedimos sinceras desculpas pelo transtorno! 🙏\n\nQualquer dúvida estamos à disposição 🤍\n${atend} — Time de Encantamento SG`
    if (devolv) return `Olá, ${nome}! Tudo bem? Me chamo ${atend} e faço parte do time de encantamento da SG 🤍\n\nEstou entrando em contato sobre seu pedido *#${r.nuvem}*.\n\nIdentificamos que sua encomenda foi *devolvida* ao nosso centro de distribuição após tentativas de entrega sem sucesso. 😔\n\nPara realizarmos um novo envio sem nenhum custo adicional, poderia confirmar seu endereço de entrega completo respondendo esta mensagem?\n\nEstamos aqui para resolver isso da melhor forma para você 🤍\n${atend} — Time de Encantamento SG`
    if (atraso) return `Olá, ${nome}! Tudo bem? Me chamo ${atend} e faço parte do time de encantamento da SG 🤍\n\nEstou entrando em contato sobre seu pedido *#${r.nuvem}*.\n\nIdentificamos um atraso na entrega pela transportadora *${r.transportadora}*. O prazo previsto era *${r.prazo||"—"}* e já estamos acompanhando de perto.\n\nAssim que tivermos uma atualização, te avisamos imediatamente! Pedimos desculpas pelo inconveniente 🙏\n\nQualquer dúvida estamos à disposição 🤍\n${atend} — Time de Encantamento SG`
    if (parado) return `Olá, ${nome}! Tudo bem? Me chamo ${atend} e faço parte do time de encantamento da SG 🤍\n\nEstou entrando em contato sobre seu pedido *#${r.nuvem}*.\n\nPercebemos que seu pedido está *em trânsito* com a transportadora *${r.transportadora}*, mas sem novas atualizações de rastreio nos últimos dias. Já estamos apurando a situação.\n\nRetornaremos em breve! Pedimos desculpas pela espera 🙏\n\nQualquer dúvida estamos à disposição 🤍\n${atend} — Time de Encantamento SG`
    return linkRastreio
      ? `Olá, ${nome}! Tudo bem? Me chamo ${atend} e faço parte do time de encantamento da SG 🤍\n\nVocê pode rastrear seu pedido *#${r.nuvem}* diretamente neste link:\n${linkRastreio}\n\nStatus atual: *${r.status}*${r.prazo?`\nPrazo previsto: *${r.prazo}*`:""}\n\nSe tiver qualquer dúvida, estamos à disposição! 🤍\n${atend} — Time de Encantamento SG`
      : `Olá, ${nome}! Tudo bem? Me chamo ${atend} e faço parte do time de encantamento da SG 🤍\n\nEstou entrando em contato sobre seu pedido *#${r.nuvem}*.\n\nStatus atual: *${r.status}*${r.prazo?`\nPrazo previsto: *${r.prazo}*`:""}\n\nQualquer dúvida estamos à disposição! 🤍\n${atend} — Time de Encantamento SG`
  } else {
    const assinatura = `Atenciosamente,\n${atend}\nTime de Encantamento — Saint Germain`
    const det = `• Pedido: #${r.nuvem}\n• Transportadora: ${r.transportadora}\n• Rastreio: ${r.rastreio||"—"}\n• Prazo previsto: ${r.prazo||"—"}`
    if (extrav) return `Assunto: Pedido #${r.nuvem} — Objeto Extraviado\n\nOlá, ${r.destinatario}! Tudo bem?\n\nMe chamo ${atend} e faço parte do time de encantamento da SG 🤍\n\nEntramos em contato sobre uma ocorrência no seu pedido.\n\n${det}\nStatus: Objeto extraviado\n\nJá acionamos nossa equipe de logística para apurar com urgência. Retornaremos em até 2 dias úteis.\n\nPedimos sinceras desculpas pelo transtorno.\n\n${assinatura}`
    if (devolv) return `Assunto: Pedido #${r.nuvem} — Devolução de Encomenda\n\nOlá, ${r.destinatario}! Tudo bem?\n\nMe chamo ${atend} e faço parte do time de encantamento da SG 🤍\n\nSua encomenda retornou ao nosso centro de distribuição após tentativas de entrega sem sucesso.\n\n${det}\n\nPara realizarmos um novo envio sem custo adicional, pedimos que confirme seu endereço respondendo a este chamado.\n\n${assinatura}`
    if (atraso) return `Assunto: Pedido #${r.nuvem} — Atraso na Entrega\n\nOlá, ${r.destinatario}! Tudo bem?\n\nMe chamo ${atend} e faço parte do time de encantamento da SG 🤍\n\nIdentificamos um atraso na entrega do seu pedido.\n\n${det}\n\nEstamos acompanhando o caso junto à transportadora e te manteremos informado(a).\n\nPedimos desculpas pelo inconveniente 🙏\n\n${assinatura}`
    if (parado) return `Assunto: Pedido #${r.nuvem} — Atualização de Rastreio\n\nOlá, ${r.destinatario}! Tudo bem?\n\nMe chamo ${atend} e faço parte do time de encantamento da SG 🤍\n\nSeu pedido está em trânsito, porém sem novas atualizações de rastreio nos últimos dias. Já estamos apurando com a transportadora.\n\n${det}\n\nRetornaremos em breve. Pedimos desculpas pela espera 🙏\n\n${assinatura}`
    return linkRastreio
      ? `Assunto: Pedido #${r.nuvem} — Rastreio\n\nOlá, ${r.destinatario}! Tudo bem?\n\nMe chamo ${atend} e faço parte do time de encantamento da SG 🤍\n\nVocê pode rastrear seu pedido neste link:\n${linkRastreio}\n\n${det}\nStatus atual: ${r.status}\n\nQualquer dúvida estamos à disposição! 🤍\n\n${assinatura}`
      : `Assunto: Pedido #${r.nuvem} — Atualização\n\nOlá, ${r.destinatario}! Tudo bem?\n\nMe chamo ${atend} e faço parte do time de encantamento da SG 🤍\n\nEntramos em contato para informar sobre seu pedido.\n\n${det}\nStatus atual: ${r.status}\n\nQualquer dúvida estamos à disposição! 🤍\n\n${assinatura}`
  }
}

// ─── Classificação automática de problemas (NOVO) ────────────
function classificarProblema(r) {
  const s = (r.status||"").toLowerCase()
  const dias = diasSemMov(r.ultimaMov)
  const dt = parsePrazo(r.prazo)
  const hoje = new Date(); hoje.setHours(0,0,0,0)
  const diasAtraso = dt ? Math.ceil((hoje-dt)/86400000) : 0

  if (s.includes("devolv")||s.includes("recusa"))                        return "DEVOLUCAO"
  if (s.includes("extravia")||s.includes("perdid"))                      return "EXTRAVIO"
  if (s.includes("falha")||s.includes("problema")||(s.includes("tent")&&s.includes("entrega"))) return "ENDERECO"
  if (s.includes("aguardando_retirada")||s.includes("aguardando retirada")||s.includes("retirada_transportadora")) return "AGUARDANDO"
  if (diasAtraso>0 && dias!==null && dias>4)                             return "POSSIVEL_EXTRAVIO"
  if (diasAtraso>0)                                                       return "ATRASO"
  return "OK"
}

// Links das transportadoras — abre direto no site ao acionar
const TRANSP_LINKS = {
  "correios":          "https://rastreamento.correios.com.br",
  "jadlog":            "https://www.jadlog.com.br/siteInstitucional/tracking.jad",
  "loggi":             "https://www.loggi.com/rastreador/",
  "total express":     "https://www.totalexpress.com.br/rastreio",
  "sequoia":           "https://rastreamento.sequoialog.com.br",
  "azul cargo":        "https://www.azulcargo.com.br/rastreio",
  "latam cargo":       "https://www.latamcargo.com",
  "jamef":             "https://www.jamef.com.br/rastreamento",
  "fedex":             "https://www.fedex.com/pt-br/tracking.html",
  "ups":               "https://www.ups.com/track",
  "dhl":               "https://www.dhl.com/br-pt/home/tracking.html",
  "tnt":               "https://www.tnt.com/express/pt_br/site/tracking.html",
  "braspress":         "https://www.braspress.com/rastreio",
  "rodonaves":         "https://www.rodonaves.com.br/rastreie-sua-carga",
  "gollog":            "https://gollog.com.br/rastreio",
  "rappi":             "https://www.rappi.com.br",
  "ifood":             "https://www.ifood.com.br",
  "kangu":             "https://kangu.com.br",
  "shein":             "https://www.shein.com.br",
  "shopee":            "https://shopee.com.br",
  "melhor envio":      "https://melhorenvio.com.br/rastreio",
  "frenet":            "https://www.frenet.com.br",
  "mandae":            "https://www.mandae.com.br/rastreio",
  "flash courier":     "https://www.flashcourier.com.br/rastreio",
  "tudo vai":          "https://www.tudovai.com.br",
}
function getTranspLink(transportadora) {
  if (!transportadora) return null
  const t = norm(transportadora)
  for (const [key, url] of Object.entries(TRANSP_LINKS)) {
    if (t.includes(norm(key))) return url
  }
  // Fallback: busca no Google pela transportadora
  return `https://www.google.com/search?q=${encodeURIComponent(transportadora + " rastreio contato")}`
}
function isTotalExpress(transportadora) {
  const t = norm(transportadora)
  return t.includes("total") || t.includes("tex")
}
function latestTotalTracking(item) {
  const events = Array.isArray(item?.tracking) ? item.tracking : []
  return events
    .filter(e => e?.data || e?.descricao)
    .sort((a,b)=>String(b.data||"").localeCompare(String(a.data||"")))[0] || null
}
function totalExpressPayloadForOrder(r) {
  if (r?.rastreio) return { awbs:[r.rastreio] }
  if (r?.nuvem) return { pedidos:[String(r.nuvem)] }
  if (r?.nf) return { notasFiscais:[String(r.nf)] }
  return null
}
function totalExpressUpdateFromItem(item) {
  const last = latestTotalTracking(item)
  const status = last?.descricao || ""
  const prazo = item?.previsaoEntregaAtualizada || item?.previsaoEntrega || ""
  const proof = item?.dadosRecebedor?.comprovanteEntrega?.urlArquivo || ""
  const urgencia = calcUrg(prazo, status)
  return {
    rastreio: item?.awb || item?.codigoBarra || "",
    status,
    prazo,
    ultimaMov: last?.data || "",
    motivo: calcMotivo(status),
    urgencia,
    acionar: calcAcionar(urgencia, status),
    totalExpressLastSync: new Date().toISOString(),
    comprovanteEntregaUrl: proof,
  }
}


const PROBLEMA_CONFIG = {
  ATRASO:            {label:"Atraso",           color:C.amber,bg:C.amberSoft, bd:C.amberBorder,icone:"⏰",sugestao:"Notificar cliente sobre o atraso na entrega"},
  POSSIVEL_EXTRAVIO: {label:"Possível Extravio", color:C.red,  bg:C.redSoft,  bd:C.redBorder,  icone:"🚨",sugestao:"Acionar transportadora imediatamente — possível extravio"},
  EXTRAVIO:          {label:"Objeto Extraviado", color:C.red,  bg:C.redSoft,  bd:C.redBorder,  icone:"🚨",sugestao:"Acionar transportadora imediatamente — confirmar localização do objeto"},
  AGUARDANDO:        {label:"Aguard. retirada transp.", color:C.blue, bg:C.blueSoft,  bd:C.blueBorder, icone:"🏭",sugestao:"Acionar transportadora para retirada no CD"},
  ENDERECO:          {label:"Problema na entrega", color:C.amber,bg:C.amberSoft, bd:C.amberBorder,icone:"📍",sugestao:"Confirmar dados de endereço com o cliente"},
  DEVOLUCAO:         {label:"Devolução",         color:C.amber,bg:C.amberSoft, bd:C.amberBorder,icone:"↩", sugestao:"Tratar reenvio ou reembolso com o cliente"},
  OK:                {label:"Sem pendências",    color:C.green,bg:C.greenSoft, bd:C.greenBorder,icone:"✓", sugestao:"Pedido sem problemas críticos identificados"},
}

// ─── Design Tokens para badges ───────────────────────────────
const urgStyles = {
  Alta:  {bg:C.redSoft,  color:C.red,  bd:C.redBorder,  dot:"#e74c3c"},
  Média: {bg:C.amberSoft,color:C.amber,bd:C.amberBorder, dot:"#F3D36B"},
  Baixa: {bg:C.greenSoft,color:C.green,bd:C.greenBorder, dot:"#27ae60"},
  "—":   {bg:C.creamDark,color:C.text3,bd:C.border,      dot:C.text4},
}
const acionStyles = {
  Sim:    {bg:C.redSoft,  color:C.red,  bd:C.redBorder},
  Avaliar:{bg:C.amberSoft,color:C.amber,bd:C.amberBorder},
  Não:    {bg:C.greenSoft,color:C.green,bd:C.greenBorder},
}
const atendStyles = {
  Aberto:         {bg:C.redSoft,  color:C.red,  bd:C.redBorder},
  "Em andamento": {bg:C.amberSoft,color:C.amber,bd:C.amberBorder},
  Resolvido:      {bg:C.greenSoft,color:C.green,bd:C.greenBorder},
}

// ─── Formatação de status (remove underscore, capitaliza) ────────
function formatStatus(s) {
  if (!s) return "—"
  const r = s.replace(/_/g," ")
    .replace(/em transito/i,"Em trânsito")
    .replace(/saiu para entrega/i,"Saiu para entrega")
    .replace(/aguardando retirada transportadora/i,"Aguard. retirada transp.")
    .replace(/problema entrega/i,"Problema na entrega")
    .replace(/sem informacao/i,"Sem informação")
    .replace(/triado/i,"Triado")
  return r.charAt(0).toUpperCase() + r.slice(1)
}

// ─── Componentes base ─────────────────────────────────────────
function Chip({val,styles}) {
  const s = styles[val]||{bg:C.creamDark,color:C.text3,bd:C.border}
  return <span style={{background:s.bg,color:s.color,border:`1px solid ${s.bd}`,borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:500,whiteSpace:"nowrap",letterSpacing:"0.02em"}}>{val}</span>
}

// BUG FIX #14: StatusBadge — bg:"#EBF5FB" era syntax error, corrigido para bg="#EBF5FB"
function StatusBadge({val}) {
  const s = (val||"").toLowerCase()
  let bg=C.creamDark, color=C.text3, bd=C.border
  if (s.includes("entregue")||s.includes("finaliz"))       {bg=C.greenSoft; color=C.green; bd=C.greenBorder}
  else if (s.includes("trânsito")||s.includes("transito")) {bg=C.blueSoft;  color=C.blue;  bd=C.blueBorder}
  else if (s.includes("saiu"))                             {bg="#EBF5FB";   color=C.blue;  bd=C.blueBorder}
  else if (s.includes("extravia"))                         {bg=C.redSoft;   color=C.red;   bd=C.redBorder}
  else if (s.includes("devolv"))                           {bg=C.amberSoft; color=C.amber; bd=C.amberBorder}
  return <span style={{background:bg,color,border:`1px solid ${bd}`,borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:500,whiteSpace:"nowrap",letterSpacing:"0.02em"}}>{formatStatus(val)}</span>
}

function SlaCell({prazo}) {
  const sla = slaInfo(prazo)
  return <div style={{lineHeight:1.6}}>
    <div style={{fontSize:11,color:C.text2,fontWeight:400}}>{prazo||"—"}</div>
    {sla&&<span style={{background:sla.bg,color:sla.color,border:`1px solid ${sla.bd}`,borderRadius:10,padding:"1px 7px",fontSize:10,fontWeight:500}}>{sla.label}</span>}
  </div>
}

// Situação Prazo: Antes do Prazo / No Prazo / Atraso
function SituacaoPrazoBadge({prazo, status, entregueNoPrazo}) {
  const dt = parsePrazo(prazo)
  if (entregueNoPrazo === true)  return <span style={{background:C.greenSoft,color:C.green,border:`1px solid ${C.greenBorder}`,borderRadius:10,padding:"2px 8px",fontSize:10,fontWeight:600}}>No Prazo</span>
  if (entregueNoPrazo === false) return <span style={{background:C.redSoft,color:C.red,border:`1px solid ${C.redBorder}`,borderRadius:10,padding:"2px 8px",fontSize:10,fontWeight:600}}>Atraso</span>
  if (!dt) return <span style={{background:C.creamDark,color:C.text4,border:`1px solid ${C.border}`,borderRadius:10,padding:"2px 8px",fontSize:10}}>—</span>
  // Comparar apenas data (sem hora)
  const hoje = new Date(); hoje.setHours(0,0,0,0)
  const prazoDate = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate())
  const diff = Math.ceil((prazoDate-hoje)/86400000)
  if (diff > 0)   return <span style={{background:C.greenSoft,color:C.green,border:`1px solid ${C.greenBorder}`,borderRadius:10,padding:"2px 8px",fontSize:10,fontWeight:600}}>Antes do Prazo</span>
  if (diff === 0) return <span style={{background:C.amberSoft,color:C.amber,border:`1px solid ${C.amberBorder}`,borderRadius:10,padding:"2px 8px",fontSize:10,fontWeight:600}}>No Prazo</span>
  return <span style={{background:C.redSoft,color:C.red,border:`1px solid ${C.redBorder}`,borderRadius:10,padding:"2px 8px",fontSize:10,fontWeight:600}}>Atraso</span>
}

function SemMovBadge({ultimaMov}) {
  const info = semMovInfo(ultimaMov)
  if (!info) return <span style={{fontSize:11,color:C.text4}}>{ultimaMov||"—"}</span>
  return <div>
    <div style={{fontSize:10,color:C.text3,marginBottom:2}}>{ultimaMov}</div>
    <span style={{background:info.alerta?C.redSoft:C.amberSoft,color:info.alerta?C.red:C.amber,border:`1px solid ${info.alerta?C.redBorder:C.amberBorder}`,borderRadius:10,padding:"1px 7px",fontSize:10,fontWeight:600,whiteSpace:"nowrap"}}>
      {info.alerta?"⚠ ":""}{info.label}
    </span>
  </div>
}

// BUG FIX #3: TimeOpenBadge estava sem declaração de função
function TimeOpenBadge({sentAt}) {
  const info = timeOpen(sentAt)
  if (!info) return null
  return <span style={{background:info.alert?C.redSoft:C.amberSoft,color:info.alert?C.red:C.amber,border:`1px solid ${info.alert?C.redBorder:C.amberBorder}`,borderRadius:10,padding:"2px 8px",fontSize:10,fontWeight:500}}>{info.label}</span>
}

function KpiCard({label,val,sub,accent}) {
  return <div style={{background:C.white,borderRadius:4,padding:"16px 18px",border:`1px solid ${accent?C.redBorder:C.border}`,borderLeft:`4px solid ${accent?C.red:C.brand}`,boxShadow:shadow.sm,position:"relative",overflow:"hidden"}}>
    <div style={{fontSize:10,color:C.text3,textTransform:"uppercase",marginBottom:12,fontWeight:700}}>{label}</div>
    <div style={{fontSize:30,fontWeight:800,color:accent?C.red:C.text1,lineHeight:1,marginBottom:7}}>{val}</div>
    {sub&&<div style={{fontSize:11,color:C.text3,fontWeight:500}}>{sub}</div>}
  </div>
}

function CopyBtn({text,label}) {
  const [ok,setOk] = useState(false)
  return <button onClick={()=>{navigator.clipboard.writeText(text);setOk(true);setTimeout(()=>setOk(false),2000)}}
    style={{background:ok?C.greenSoft:C.gold,border:`1px solid ${ok?C.greenBorder:C.goldDim}`,color:ok?C.green:C.white,borderRadius:6,padding:"5px 12px",fontSize:11,cursor:"pointer",fontWeight:500,whiteSpace:"nowrap",transition:"all .2s"}}>
    {ok?"✓ Copiado!":label||"Copiar"}
  </button>
}

function NavIcon({type, active}) {
  const stroke = active ? C.brand : C.text2
  const common = {width:22,height:22,display:"block"}
  const props = {stroke, strokeWidth:1.8, fill:"none", strokeLinecap:"round", strokeLinejoin:"round"}
  const icons = {
    dashboard:<svg style={common} viewBox="0 0 24 24"><path {...props} d="M4 13h6V4H4v9Zm10 7h6V4h-6v16ZM4 20h6v-3H4v3Z"/></svg>,
    logistica:<svg style={common} viewBox="0 0 24 24"><path {...props} d="M3 7h12v10H3zM15 10h3l3 3v4h-6zM7 20a2 2 0 1 0 0-4 2 2 0 0 0 0 4Zm10 0a2 2 0 1 0 0-4 2 2 0 0 0 0 4Z"/></svg>,
    suporte:<svg style={common} viewBox="0 0 24 24"><path {...props} d="M4 12a8 8 0 0 1 16 0v4a2 2 0 0 1-2 2h-2v-6h4M4 16a2 2 0 0 0 2 2h2v-6H4m8 8h2"/></svg>,
    devolucao:<svg style={common} viewBox="0 0 24 24"><path {...props} d="M9 7 5 11l4 4M5 11h9a5 5 0 0 1 0 10h-2"/></svg>,
    reenvio:<svg style={common} viewBox="0 0 24 24"><path {...props} d="M4 16v4h4M20 8V4h-4M5 19 19 5M8 5h5M16 19h-5"/></svg>,
    arquivados:<svg style={common} viewBox="0 0 24 24"><path {...props} d="M4 5h16v14H4zM8 12l3 3 5-6"/></svg>,
    usuarios:<svg style={common} viewBox="0 0 24 24"><path {...props} d="M16 11a4 4 0 1 0-8 0M4 21a8 8 0 0 1 16 0M18 8h3M19.5 6.5v3"/></svg>,
  }
  return icons[type] || icons.dashboard
}

function SortIcon({col,sortCol,sortDir}) {
  if (sortCol!==col) return <span style={{color:C.text4,fontSize:9,marginLeft:4}}>⇅</span>
  return <span style={{color:C.gold,fontSize:9,marginLeft:4}}>{sortDir==="asc"?"↑":"↓"}</span>
}

function Toast({toasts}) {
  return <div style={{position:"fixed",bottom:24,right:24,display:"flex",flexDirection:"column",gap:8,zIndex:9999,pointerEvents:"none"}}>
    {toasts.map(t=><div key={t.id} style={{background:t.type==="error"?C.red:t.type==="warn"?C.amber:C.green,color:C.white,borderRadius:10,padding:"12px 18px",fontSize:12,fontWeight:500,boxShadow:shadow.lg,maxWidth:320,lineHeight:1.5}}>{t.msg}</div>)}
  </div>
}

const getINP = () => ({borderRadius:4,border:`1px solid ${C.borderDark}`,padding:"9px 12px",fontSize:12,background:C.white,color:C.text1,outline:"none",transition:"border-color .2s, box-shadow .2s"})

function SGMonogram({size=44}) {
  return <div style={{width:size,height:size,display:"flex",alignItems:"center",justifyContent:"center",background:"#000",color:"#fff",fontFamily:"Georgia, 'Times New Roman', serif",fontSize:Math.round(size*0.62),lineHeight:1,fontWeight:400,letterSpacing:"-0.18em",paddingRight:Math.round(size*0.08)}}>SG</div>
}

function SGWordmark({dark=false, size=17}) {
  return <div style={{color:dark?"#fff":"#111",fontSize:size,textTransform:"uppercase",fontWeight:800,letterSpacing:"0.38em",lineHeight:1,whiteSpace:"nowrap"}}>SAINT GERMAIN</div>
}

// ─── Componentes de Suporte (NOVOS) ──────────────────────────

// HeaderProblema: identifica e destaca visualmente o tipo de problema
function HeaderProblema({r, onNotificou}) {
  const tipo = classificarProblema(r)
  const cfg  = PROBLEMA_CONFIG[tipo]
  const dias = diasSemMov(r.ultimaMov)
  const dt   = parsePrazo(r.prazo)
  const hoje = new Date(); hoje.setHours(0,0,0,0)
  const diasAtraso = dt ? Math.ceil((hoje-dt)/86400000) : 0
  const critico = tipo==="EXTRAVIO"||tipo==="POSSIVEL_EXTRAVIO"

  const metricas = [
    {lbl:"Transportadora",   val:r.transportadora||"—"},
    {lbl:"Última atualização",val:r.ultimaMov||"—"},
    diasAtraso>0 ? {lbl:"Dias de atraso",      val:`${diasAtraso}d`,       cor:C.red}  : null,
    dias!==null  ? {lbl:"Sem movimentação",     val:`${dias>=ALERTA_DIAS?"🔥 ":""}${dias}d`, cor:dias>=ALERTA_DIAS?C.red:C.amber} : null,
  ].filter(Boolean)

  return (
    <div style={{background:C.white,borderRadius:6,border:`1px solid ${cfg.bd}`,borderLeft:`4px solid ${cfg.color}`,padding:"10px 12px",marginBottom:10,boxShadow:shadow.sm}}>
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
        <span style={{background:cfg.bg,color:cfg.color,border:`1px solid ${cfg.bd}`,borderRadius:4,padding:"3px 8px",fontSize:10,fontWeight:800,textTransform:"uppercase"}}>{cfg.label}</span>
        <div style={{flex:1}}/>
        <Chip val={r.urgencia} styles={urgStyles}/>
        <Chip val={r.atendimento} styles={atendStyles}/>
      </div>
      <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
        {metricas.map(m=>(
          <div key={m.lbl} style={{background:C.cream,borderRadius:4,padding:"6px 9px",flex:1,minWidth:80,border:`1px solid ${C.border}`}}>
            <div style={{fontSize:8,color:C.text3,textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:2}}>{m.lbl}</div>
            <div style={{fontSize:12,fontWeight:600,color:m.cor||C.text1}}>{m.val}</div>
          </div>
        ))}
      </div>
      {r.alertaStatus&&(
        <div style={{marginTop:10,background:C.amberSoft,borderRadius:4,padding:"9px 12px",display:"flex",alignItems:"center",justifyContent:"space-between",gap:12,border:`1px solid ${C.amberBorder}`}}>
          <div style={{fontSize:11,color:C.amber,fontWeight:600}}>⚠ {r.alertaStatus}</div>
          <button onClick={onNotificou} style={{background:C.amber,border:"none",color:C.white,borderRadius:6,padding:"5px 12px",fontSize:10,cursor:"pointer",fontWeight:600,whiteSpace:"nowrap"}}>✓ Notifiquei</button>
        </div>
      )}
    </div>
  )
}

// SugestaoSistema: caixa de sugestão automática baseada no problema
function SugestaoSistema({r}) {
  const tipo = classificarProblema(r)
  const cfg  = PROBLEMA_CONFIG[tipo]
  if (tipo==="OK") return null
  return (
    <div style={{background:cfg.bg,border:`1px solid ${cfg.bd}`,borderRadius:10,padding:"10px 14px",marginBottom:14,display:"flex",alignItems:"center",gap:10}}>
      <span style={{fontSize:18}}>{cfg.icone}</span>
      <div>
        <div style={{fontSize:8,color:cfg.color,textTransform:"uppercase",letterSpacing:"0.12em",fontWeight:700,marginBottom:2}}>Sugestão do sistema</div>
        <div style={{fontSize:12,color:C.text1,fontWeight:500}}>{cfg.sugestao}</div>
      </div>
    </div>
  )
}

// Monta link mailto com assunto e corpo personalizados por tipo de problema
function buildMailto(r, nomeAtendente) {
  const email = (r.email||"").trim()
  if (!email) return null
  const tipo  = classificarProblema(r)
  const atend = (nomeAtendente||"").split(" ")[0] || "Time SG"
  const nome  = (r.destinatario||"").split(" ")[0] || "Cliente"

  const assuntos = {
    EXTRAVIO:          `Pedido #${r.nuvem} — Objeto Extraviado`,
    POSSIVEL_EXTRAVIO: `Pedido #${r.nuvem} — Pedido sem atualização`,
    DEVOLUCAO:         `Pedido #${r.nuvem} — Devolução de Encomenda`,
    ENDERECO:          `Pedido #${r.nuvem} — Tentativa de Entrega`,
    ATRASO:            `Pedido #${r.nuvem} — Atraso na Entrega`,
    OK:                `Pedido #${r.nuvem} — Atualização`,
  }
  const subject = assuntos[tipo] || `Pedido #${r.nuvem}`
  const body    = getTemplate(r, "zendesk", nomeAtendente)
    .replace(/^Assunto:.*\n\n/,"") // remove linha de assunto do template Zendesk

  return `mailto:${email}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`
}

// AcoesRapidas: botões de ação com feedback visual de loading
function AcoesRapidas({r, perms, nomeAtendente, onNotificar, onAcionarTransp, onReenvio, onResolver, onDevolver, onTotalTracking, onTotalTicket}) {
  const [loading, setLoading] = useState(null)
  const transpUrl = getTranspLink(r.transportadora)
  const mailtoUrl = buildMailto(r, nomeAtendente)
  const temEmail  = !!mailtoUrl
  const isTotal   = isTotalExpress(r.transportadora)

  const act = (key, fn) => async () => {
    setLoading(key); try { await fn() } finally { setLoading(null) }
  }

  if (!perms?.canOperate) return null

  const btnBase = {
    minHeight:42,
    borderRadius:6,
    padding:"10px 12px",
    fontSize:11,
    fontWeight:800,
    cursor:loading?"wait":"pointer",
    letterSpacing:"0.01em",
    transition:"all .2s",
    display:"flex",
    alignItems:"center",
    justifyContent:"center",
    gap:6,
    whiteSpace:"nowrap",
  }
  const btn = (key, label, style={}, onClick=null) => (
    <button key={key}
      onClick={onClick || act(key, {notif:onNotificar,transp:onAcionarTransp,resol:onResolver,dev:onDevolver}[key])}
      disabled={loading!==null}
      style={{...btnBase,opacity:loading!==null&&loading!==key?0.45:1,...style}}>
      {loading===key?"Aguarde...":label}
    </button>
  )

  return (
    <div style={{border:`1px solid ${C.border}`,background:C.cream,borderRadius:8,padding:10,marginBottom:10}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:10,marginBottom:8}}>
        <div style={{fontSize:8,color:C.text3,textTransform:"uppercase",letterSpacing:"0.14em",fontWeight:800}}>Acoes do atendimento</div>
        <div style={{fontSize:10,color:C.text4,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
          {temEmail?`Canal: ${r.email}`:"Canal: Zendesk / email pendente"} · {r.transportadora||"Transportadora nao informada"}
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,minmax(0,1fr))",gap:8}}>
        {temEmail ? (
          <a href={mailtoUrl}
            onClick={()=>{ onNotificar() }}
            target="_blank" rel="noopener noreferrer"
            style={{...btnBase,background:C.brand,color:C.white,border:"none",textDecoration:"none"}}>
            Notificar por Email
          </a>
        ) : (
          <button
            onClick={()=>{ onNotificar(); }}
            disabled={loading!==null}
            title="Adicione o email do cliente no arquivo para habilitar envio direto"
            style={{...btnBase,background:C.brand,color:C.white,border:"none",opacity:0.72}}>
            Notificar por Email
          </button>
        )}
        {isTotal
          ? btn("texTicket","Acionar Total Express",{background:C.blueSoft,color:C.blue,border:`1px solid ${C.blueBorder}`},act("texTicket", onTotalTicket))
          : btn("transp",`Acionar ${r.transportadora||"Transportadora"}`,{background:C.blueSoft,color:C.blue,border:`1px solid ${C.blueBorder}`},async()=>{
              setLoading("transp")
              try { await onAcionarTransp() } finally { setLoading(null) }
              if (transpUrl) window.open(transpUrl, "_blank", "noopener")
            })}
        {r.atendimento!=="Resolvido"&&btn("resol","Marcar como resolvido",{background:C.green,color:C.white,border:`1px solid ${C.green}`})}
        {btn("dev","Devolver",{background:C.white,color:C.text2,border:`1px solid ${C.borderDark}`})}
      </div>
    </div>
  )
}

// TimelineHistorico: histórico em ordem reversa com usuário + ação
function TimelineHistorico({historico, isOpen, onToggle}) {
  return (
    <div style={{background:C.white,borderRadius:12,border:`1px solid ${C.border}`,overflow:"hidden",boxShadow:shadow.sm}}>
      <button onClick={onToggle} style={{width:"100%",padding:"14px 18px",background:isOpen?C.cream:"transparent",border:"none",cursor:"pointer",display:"flex",justifyContent:"space-between",alignItems:"center",fontSize:11,fontWeight:500,color:C.text1,letterSpacing:"0.04em"}}>
        <span>Histórico{historico.length>0?` (${historico.length})`:""}</span>
        <span style={{fontSize:16,color:C.text4,fontWeight:300,width:24,textAlign:"center"}}>{isOpen?"−":"+"}</span>
      </button>
      {isOpen&&(
        <div style={{padding:16,borderTop:`1px solid ${C.border}`,background:C.cream}}>
          {historico.length===0
            ? <div style={{fontSize:12,color:C.text4}}>Nenhuma ação registrada.</div>
            : <div style={{display:"flex",flexDirection:"column",gap:0}}>
                {[...historico].reverse().map((h,i)=>(
                  <div key={i} style={{display:"flex",gap:12,alignItems:"flex-start",padding:"8px 0",borderBottom:i<historico.length-1?`1px solid ${C.border}55`:"none"}}>
                    <div style={{width:6,height:6,borderRadius:"50%",background:C.gold,flexShrink:0,marginTop:5}}/>
                    <div style={{flex:1}}>
                      <div style={{fontSize:10,color:C.text4,marginBottom:2}}>
                        {h.ts}
                        {h.usuario&&<span style={{marginLeft:6,color:C.gold,fontWeight:600}}>· {h.usuario}</span>}
                      </div>
                      <div style={{fontSize:11,color:C.text2,fontWeight:500}}>{h.acao}</div>
                    </div>
                  </div>
                ))}
              </div>
          }
        </div>
      )}
    </div>
  )
}

// ─── Login ────────────────────────────────────────────────────
function LoginScreen({onLogin}) {
  const [login,setLogin]=useState("")
  const [password,setPassword]=useState("")
  const [loading,setLoading]=useState(false)
  const [error,setError]=useState("")
  const handle = async e => {
    e.preventDefault(); setLoading(true); setError("")
    try { onLogin(await signIn(login,password)) }
    catch(err) { setError(err.message) }
    finally { setLoading(false) }
  }
  return (
    <div style={{minHeight:"100vh",background:C.brand,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",top:0,left:0,right:0,height:"1px",background:`linear-gradient(90deg,transparent,${C.gold},transparent)`}}/>
      <div style={{position:"absolute",bottom:0,left:0,right:0,height:"1px",background:`linear-gradient(90deg,transparent,${C.gold},transparent)`}}/>
      <div style={{position:"absolute",top:"50%",left:"50%",transform:"translate(-50%,-50%)",width:600,height:600,borderRadius:"50%",border:`1px solid ${C.gold}18`,pointerEvents:"none"}}/>
      <div style={{position:"absolute",top:"50%",left:"50%",transform:"translate(-50%,-50%)",width:400,height:400,borderRadius:"50%",border:`1px solid ${C.gold}12`,pointerEvents:"none"}}/>
      <div style={{textAlign:"center",marginBottom:48,position:"relative"}}>
        <div style={{display:"flex",justifyContent:"center",marginBottom:18}}><SGMonogram size={72}/></div>
        <SGWordmark dark size={24}/>
        <div style={{width:48,height:"1px",background:"#BDBDBD",margin:"16px auto"}}/>
        <div style={{fontSize:9,letterSpacing:"0.28em",color:"#BDBDBD",textTransform:"uppercase"}}>Central de Pedidos</div>
      </div>
      <div style={{background:C.white,borderRadius:16,padding:"40px 44px",width:"100%",maxWidth:400,boxShadow:shadow.lg,position:"relative"}}>
        <div style={{position:"absolute",top:0,left:"50%",transform:"translateX(-50%)",width:48,height:2,background:C.gold,borderRadius:1}}/>
        <form onSubmit={handle}>
          <div style={{marginBottom:20}}>
            <label style={{fontSize:9,letterSpacing:"0.14em",textTransform:"uppercase",color:C.text3,fontWeight:500,display:"block",marginBottom:7}}>Login</label>
            <input value={login} onChange={e=>setLogin(e.target.value)} type="text" placeholder="seu login" required autoCapitalize="none" autoComplete="username" style={{...getINP(),width:"100%",boxSizing:"border-box",fontSize:13}}/>
          </div>
          <div style={{marginBottom:28}}>
            <label style={{fontSize:9,letterSpacing:"0.14em",textTransform:"uppercase",color:C.text3,fontWeight:500,display:"block",marginBottom:7}}>Senha</label>
            <input value={password} onChange={e=>setPassword(e.target.value)} type="password" placeholder="••••••••" required autoComplete="current-password" style={{...getINP(),width:"100%",boxSizing:"border-box",fontSize:13}}/>
          </div>
          {error&&<div style={{background:C.redSoft,color:C.red,border:`1px solid ${C.redBorder}`,borderRadius:8,padding:"10px 14px",fontSize:12,marginBottom:20,lineHeight:1.5}}>{error}</div>}
          <button type="submit" disabled={loading} style={{width:"100%",background:loading?C.text4:C.brand,border:"none",color:C.white,borderRadius:8,padding:"13px 0",fontSize:11,fontWeight:500,cursor:loading?"not-allowed":"pointer",letterSpacing:"0.18em",textTransform:"uppercase",transition:"background .2s"}}>
            {loading?"Autenticando...":"Entrar"}
          </button>
        </form>
      </div>
      <div style={{marginTop:32,fontSize:9,letterSpacing:"0.2em",color:`${C.white}33`,textTransform:"uppercase"}}>Acesso restrito · Uso interno</div>
    </div>
  )
}

// ─── Painel de Usuários ───────────────────────────────────────
function UsuariosPanel({token,addToast}) {
  const [usuarios,setUsuarios]=useState([])
  const [loading,setLoading]=useState(true)
  const [form,setForm]=useState({login:"",nome:"",perfil:"logistica",senha:""})
  const [saving,setSaving]=useState(false)
  const LABEL={admin:"Admin",logistica:"Logística",suporte:"Suporte",leitura:"Somente leitura"}
  useEffect(()=>{loadUsuarios(token).then(d=>{setUsuarios(d);setLoading(false)})},[])
  const handleCreate = async e => {
    e.preventDefault(); setSaving(true)
    try {
      const authEmail = authEmailFromLogin(form.login)
      if (!authEmail) throw new Error("Informe um login valido")
      const auth = await createUser(authEmail,form.senha,token)
      await saveUsuario({id:auth.id,email:authEmail,nome:form.nome,perfil:form.perfil,ativo:true},token)
      addToast(`Usuario ${displayLogin(authEmail)} criado!`)
      setForm({login:"",nome:"",perfil:"logistica",senha:""})
      setUsuarios(await loadUsuarios(token))
    } catch(err) { addToast("Erro: "+err.message,"error") }
    finally { setSaving(false) }
  }
  const handleDelete = async id => {
    if (!confirm("Remover este usuário?")) return
    await deleteUsuario(id,token); setUsuarios(u=>u.filter(x=>x.id!==id)); addToast("Usuário removido","warn")
  }
  const handlePerfil = async (id,perfil) => {
    const u = usuarios.find(x=>x.id===id)
    await saveUsuario({...u,perfil},token)
    setUsuarios(prev=>prev.map(x=>x.id===id?{...x,perfil}:x)); addToast("Perfil atualizado")
  }
  return (
    <div style={{padding:"32px 40px",maxWidth:900}}>
      <div style={{marginBottom:28}}>
        <div style={{fontSize:9,letterSpacing:"0.18em",textTransform:"uppercase",color:C.gold,marginBottom:6}}>Administração</div>
        <div style={{fontSize:22,fontWeight:800,color:C.text1}}>Gestão de Usuários</div>
      </div>

      <div style={{background:C.white,borderRadius:14,border:`1px solid ${C.border}`,padding:28,marginBottom:24,boxShadow:shadow.sm}}>
        <div style={{fontSize:9,letterSpacing:"0.14em",textTransform:"uppercase",color:C.text3,marginBottom:20,fontWeight:500}}>Adicionar usuário</div>
        <form onSubmit={handleCreate}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:14}}>
            {[["Nome","nome","text","Nome completo"],["Login","login","text","ex: rodrigo"],["Senha inicial","senha","password","Minimo 6 caracteres"]].map(([lbl,key,type,ph])=>(
              <div key={key}>
                <label style={{fontSize:9,letterSpacing:"0.12em",textTransform:"uppercase",color:C.text3,fontWeight:500,display:"block",marginBottom:6}}>{lbl}</label>
                <input value={form[key]} onChange={e=>setForm(f=>({...f,[key]:key==="login"?cleanLogin(e.target.value):e.target.value}))} type={type} placeholder={ph} required={key!=="nome"} minLength={key==="senha"?6:undefined} autoCapitalize={key==="login"?"none":undefined} autoComplete={key==="login"?"username":undefined} style={{...getINP(),width:"100%",boxSizing:"border-box"}}/>
              </div>
            ))}
            <div>
              <label style={{fontSize:9,letterSpacing:"0.12em",textTransform:"uppercase",color:C.text3,fontWeight:500,display:"block",marginBottom:6}}>Perfil de acesso</label>
              <select value={form.perfil} onChange={e=>setForm(f=>({...f,perfil:e.target.value}))} style={{...getINP(),width:"100%",boxSizing:"border-box"}}>
                {Object.entries(LABEL).map(([k,v])=><option key={k} value={k}>{v}</option>)}
              </select>
            </div>
          </div>
          <button type="submit" disabled={saving} style={{background:saving?C.text4:C.brand,border:"none",color:C.white,borderRadius:8,padding:"10px 24px",fontSize:11,fontWeight:500,cursor:saving?"not-allowed":"pointer",letterSpacing:"0.1em",textTransform:"uppercase"}}>
            {saving?"Criando...":"+ Criar usuário"}
          </button>
        </form>
      </div>
      <div style={{background:C.white,borderRadius:14,border:`1px solid ${C.border}`,overflow:"hidden",boxShadow:shadow.sm}}>
        {loading?<div style={{padding:32,textAlign:"center",color:C.text4}}>Carregando...</div>:(
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead>
              <tr style={{background:C.brand}}>
                {["Nome","Login","Perfil","Ações"].map(h=><th key={h} style={{padding:"13px 18px",textAlign:"left",color:C.gold,fontWeight:400,fontSize:9,letterSpacing:"0.14em",textTransform:"uppercase"}}>{h}</th>)}
              </tr>
            </thead>
            <tbody>
              {usuarios.length===0?<tr><td colSpan={4} style={{padding:32,textAlign:"center",color:C.text4}}>Nenhum usuário cadastrado</td></tr>
              :usuarios.map((u,i)=>(
                <tr key={u.id} style={{background:i%2===0?C.white:C.cream,borderBottom:`1px solid ${C.border}`}}>
                  <td style={{padding:"12px 18px",color:C.text1,fontWeight:500}}>{u.nome||"—"}</td>
                  <td style={{padding:"12px 18px",color:C.text2}}>{displayLogin(u.email)}</td>
                  <td style={{padding:"12px 18px"}}>
                    <select value={u.perfil} onChange={e=>handlePerfil(u.id,e.target.value)} style={{...getINP(),padding:"5px 10px",fontSize:11}}>
                      {Object.entries(LABEL).map(([k,v])=><option key={k} value={k}>{v}</option>)}
                    </select>
                  </td>
                  <td style={{padding:"12px 18px"}}>
                    <button onClick={()=>handleDelete(u.id)} style={{background:C.redSoft,border:`1px solid ${C.redBorder}`,color:C.red,borderRadius:6,padding:"4px 12px",fontSize:11,cursor:"pointer",fontWeight:500}}>Remover</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  )
}

const SAMPLE=`Identificador Ecommerce;Destinatário Nome;Estratégia de Frete;Rastreador Last Mile;Situação;Prazo Logístico;Nº Nota Fiscal
12345;Ana Souza;Correios PAC;AA123456789BR;Em trânsito;05/05/2026;98765
12346;Carlos Lima;Jadlog;JD987654321;Extraviado;28/04/2026;98766
12347;Mariana Costa;Total Express;TE112233445;Entregue;01/05/2026;98767
12348;Fernando Silva;Correios SEDEX;AA223344556BR;Saiu para entrega;02/05/2026;98768
12349;Júlia Martins;Loggi;LG556677889;Devolvido;30/04/2026;98769`

// ─── App ──────────────────────────────────────────────────────
// ─── Kanban Suporte ───────────────────────────────────────────
const KANBAN_COLS = [
  {id:"DEVOLUCAO",  label:"Devolução",               color:C.amber, bg:C.amberSoft, bd:C.amberBorder, icone:"↩"},
  {id:"ENDERECO",   label:"Problema na entrega",      color:C.amber, bg:C.amberSoft, bd:C.amberBorder, icone:"📍"},
  {id:"AGUARDANDO", label:"Aguard. retirada transp.", color:C.blue,  bg:C.blueSoft,  bd:C.blueBorder,  icone:"🏭"},
  {id:"EXTRAVIO",   label:"Extravio",                 color:C.red,   bg:C.redSoft,   bd:C.redBorder,   icone:"🚨"},
  {id:"OK",         label:"Outros",                   color:C.text2, bg:C.creamDark, bd:C.border,      icone:"◎"},
]
function KanbanSuporteView({rows, onSelect, selSup, perms, upd, nomeAtendente}) {
  const cols = KANBAN_COLS.map(col=>({
    ...col,
    items: rows.filter(r=>{
      const t = classificarProblema(r)
      if (col.id==="OK") return t==="OK"||t==="ATRASO"||t==="POSSIVEL_EXTRAVIO"
      return t===col.id
    })
  }))
  return (
    <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:10,padding:"14px 18px",overflowX:"auto",height:"100%"}}>
      {cols.map(col=>(
        <div key={col.id} style={{background:col.bg,borderRadius:12,border:`1px solid ${col.bd}`,display:"flex",flexDirection:"column",minHeight:200}}>
          <div style={{padding:"10px 14px",borderBottom:`1px solid ${col.bd}`,display:"flex",alignItems:"center",gap:6}}>
            <span style={{fontSize:14}}>{col.icone}</span>
            <span style={{fontSize:10,fontWeight:700,color:col.color,letterSpacing:"0.08em",textTransform:"uppercase"}}>{col.label}</span>
            <span style={{marginLeft:"auto",background:col.color,color:"#fff",borderRadius:10,padding:"1px 8px",fontSize:10,fontWeight:700}}>{col.items.length}</span>
          </div>
          <div style={{flex:1,overflowY:"auto",padding:8,display:"flex",flexDirection:"column",gap:6}}>
            {col.items.length===0&&<div style={{textAlign:"center",padding:"24px 8px",color:C.text4,fontSize:11}}>Sem pedidos</div>}
            {col.items.map(r=>(
              <div key={r.id} onClick={()=>onSelect(r.id===selSup?null:r.id)}
                style={{background:r.id===selSup?C.gold+"22":C.white,borderRadius:10,padding:"10px 12px",cursor:"pointer",border:`1px solid ${r.id===selSup?col.bd:C.border}`,boxShadow:shadow.sm,transition:"all .15s"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
                  <span style={{fontWeight:700,fontSize:11,color:C.text1}}>{r.nuvem}</span>
                  <Chip val={r.urgencia} styles={urgStyles}/>
                </div>
                <div style={{fontSize:10,color:C.text2,marginBottom:4,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.destinatario}</div>
                <div style={{fontSize:10,color:C.text3,marginBottom:6}}>{r.transportadora}</div>
                <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
                  <Chip val={r.atendimento} styles={atendStyles}/>
                  <TimeOpenBadge sentAt={r.sentAt}/>
                </div>
              </div>
            ))}
          </div>
        </div>
      ))}
    </div>
  )
}


// ─── FilterBar — chips system ─────────────────────────────────
function FilterBar({search, onSearch, showFilters, onToggleFilters, filters, onClearAll, compact}) {
  const chipColors = {
    "Alta":           {color:C.red,   bg:C.redSoft,   bd:C.redBorder},
    "Média":          {color:C.amber, bg:C.amberSoft, bd:C.amberBorder},
    "Baixa":          {color:C.green, bg:C.greenSoft, bd:C.greenBorder},
    "Aberto":         {color:C.red,   bg:C.redSoft,   bd:C.redBorder},
    "Em andamento":   {color:C.amber, bg:C.amberSoft, bd:C.amberBorder},
    "Atraso":         {color:C.red,   bg:C.redSoft,   bd:C.redBorder},
    "Antes do Prazo": {color:C.green, bg:C.greenSoft, bd:C.greenBorder},
    "No Prazo":       {color:C.green, bg:C.greenSoft, bd:C.greenBorder},
    "Sim":            {color:C.red,   bg:C.redSoft,   bd:C.redBorder},
  }
  const activeFilters = filters.filter(f=>f.value!=="Todos")
  const totalAtivos = activeFilters.length

  return (
    <div style={{marginBottom:compact?8:14}}>
      {/* Linha 1: busca + botão */}
      <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:8}}>
        <input value={search} onChange={e=>onSearch(e.target.value)}
          placeholder={compact?"Buscar...":"Buscar pedido, destinatário, rastreio..."}
          style={{...getINP(),flex:1,fontSize:compact?11:12}}/>
        <button onClick={onToggleFilters}
          style={{background:showFilters||totalAtivos>0?C.brand:C.white,border:`1px solid ${showFilters||totalAtivos>0?C.brand:C.border}`,color:showFilters||totalAtivos>0?C.white:C.text2,borderRadius:8,padding:compact?"7px 14px":"9px 16px",fontSize:11,cursor:"pointer",fontWeight:500,display:"flex",alignItems:"center",gap:8,whiteSpace:"nowrap",transition:"all .15s"}}>
          <span>⚙</span>
          <span>Filtros</span>
          {totalAtivos>0&&<span style={{background:C.white,color:C.red,borderRadius:10,padding:"0px 7px",fontSize:10,fontWeight:700,lineHeight:1.6}}>{totalAtivos}</span>}
        </button>
      </div>

      {/* Chips de filtros ativos */}
      {activeFilters.length>0&&(
        <div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"center",marginBottom:8}}>
          {activeFilters.map(f=>{
            const s = chipColors[f.value]||{color:C.blue,bg:C.blueSoft,bd:C.blueBorder}
            return (
              <span key={f.key} style={{background:s.bg,color:s.color,border:`1px solid ${s.bd}`,borderRadius:20,padding:"3px 8px 3px 12px",fontSize:11,fontWeight:500,display:"flex",alignItems:"center",gap:5,whiteSpace:"nowrap"}}>
                {f.label}: <strong>{f.value}</strong>
                <span onClick={()=>f.setValue("Todos")} style={{cursor:"pointer",fontSize:15,lineHeight:1,opacity:0.6,marginLeft:2}}>×</span>
              </span>
            )
          })}
          <span onClick={onClearAll} style={{fontSize:11,color:C.red,cursor:"pointer",fontWeight:500,marginLeft:4}}>Limpar tudo</span>
        </div>
      )}

      {/* Painel de filtros — abre inline abaixo */}
      {showFilters&&(
        <div style={{background:C.white,border:`1px solid ${C.border}`,borderRadius:10,padding:"14px 16px",marginBottom:8,boxShadow:shadow.sm}}>
          <div style={{display:"grid",gridTemplateColumns:`repeat(${Math.min(filters.length,3)},1fr)`,gap:12}}>
            {filters.map(f=>(
              <div key={f.key}>
                <div style={{fontSize:9,color:C.text3,fontWeight:600,letterSpacing:"0.12em",textTransform:"uppercase",marginBottom:5}}>{f.label}</div>
                <select value={f.value} onChange={e=>f.setValue(e.target.value)}
                  style={{...getINP(),width:"100%",boxSizing:"border-box",fontSize:11}}>
                  {f.opts.map(o=><option key={o}>{o}</option>)}
                </select>
              </div>
            ))}
          </div>
          {totalAtivos>0&&(
            <button onClick={onClearAll} style={{marginTop:12,background:C.redSoft,border:`1px solid ${C.redBorder}`,color:C.red,borderRadius:7,padding:"6px 16px",fontSize:11,cursor:"pointer",fontWeight:500}}>
              × Limpar todos os filtros
            </button>
          )}
        </div>
      )}
    </div>
  )
}

function OperacaoEspecialPanel({type, rows, perms, upd, onCreateReenvio, onResolve}) {
  const [search,setSearch]=useState("")
  const isDev = type==="devolucao"
  const cfg = isDev
    ? {title:"Devolucao", empty:"Nenhum pedido em devolucao", statusKey:"devolucaoStatus", options:["Aguardando tratativa","Aguardando cliente","Em transporte reverso","Recebido no CD","Reenviar","Reembolsar"]}
    : {title:"Reenvio", empty:"Nenhum reenvio cadastrado", statusKey:"reenvioStatus", options:["Pendente","Pedido criado","Em separacao","Enviado","Concluido"]}
  const q = search.toLowerCase()
  const data = rows.filter(r=>!q||[r.nuvem,r.destinatario,r.rastreio,r.status,r.novoPedido,r.novaTransportadora,r.novoRastreio,r.materialReenvio,r.motivoDevolucao].some(v=>(v||"").toLowerCase().includes(q)))
  const pendentes = rows.filter(r=>(r[cfg.statusKey]||cfg.options[0])===cfg.options[0]).length
  const emAndamento = Math.max(0, rows.length - pendentes)
  const inputStyle = {...getINP(),width:"100%",boxSizing:"border-box",fontSize:11,padding:"7px 9px"}

  return (
    <div style={{padding:"24px 32px"}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12,marginBottom:18}}>
        <KpiCard label={`Total ${cfg.title}`} val={rows.length}/>
        <KpiCard label="Pendentes" val={pendentes} accent={pendentes>0}/>
        <KpiCard label="Em andamento" val={emAndamento}/>
      </div>
      <div style={{display:"flex",gap:10,alignItems:"center",marginBottom:14}}>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder={`Buscar em ${cfg.title.toLowerCase()}...`} style={{...getINP(),flex:1,padding:"10px 14px",boxSizing:"border-box",boxShadow:shadow.sm}}/>
      </div>
      <div style={{overflowX:"auto",borderRadius:12,border:`1px solid ${C.border}`,boxShadow:shadow.sm,background:C.white}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,tableLayout:"fixed",minWidth:isDev?1040:1120}}>
          <colgroup>
            <col style={{width:90}}/><col style={{width:170}}/><col style={{width:130}}/><col style={{width:150}}/>
            {!isDev&&<col style={{width:120}}/>}
            {!isDev&&<col style={{width:130}}/>}
            {!isDev&&<col style={{width:130}}/>}
            {!isDev&&<col style={{width:150}}/>}
            <col style={{width:180}}/>
            <col style={{width:150}}/><col style={{width:130}}/><col style={{width:210}}/>
          </colgroup>
          <thead>
            <tr>{["Pedido","Cliente","Transportadora","Status origem",...(!isDev?["Novo pedido","Nova transp.","Novo rastreio","Material"]:[]),"Motivo","Etapa","Responsavel","Acoes"].map(h=><th key={h} style={{padding:"12px 14px",textAlign:"left",color:C.gold,fontWeight:400,fontSize:9,letterSpacing:"0.14em",textTransform:"uppercase",background:C.brand,whiteSpace:"nowrap"}}>{h}</th>)}</tr>
          </thead>
          <tbody>
            {data.length===0?<tr><td colSpan={isDev?8:12} style={{padding:34,textAlign:"center",color:C.text4}}>{cfg.empty}</td></tr>
            :data.map((r,i)=>(
              <tr key={r.id} style={{background:i%2===0?C.white:C.cream,borderBottom:`1px solid ${C.border}`}}>
                <td style={{padding:"11px 14px",fontWeight:700,color:C.text1}}>{r.nuvem}</td>
                <td style={{padding:"11px 14px",color:C.text2,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={r.destinatario}>{r.destinatario||"—"}</td>
                <td style={{padding:"11px 14px",color:C.text2,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.transportadora||"—"}</td>
                <td style={{padding:"11px 14px"}}><StatusBadge val={r.status}/></td>
                {!isDev&&<td style={{padding:"9px 10px"}}><input value={r.novoPedido||""} onChange={e=>upd(r.id,{novoPedido:e.target.value})} placeholder="No novo pedido" style={inputStyle}/></td>}
                {!isDev&&<td style={{padding:"9px 10px"}}><input value={r.novaTransportadora||""} onChange={e=>upd(r.id,{novaTransportadora:e.target.value})} placeholder="Transportadora" style={inputStyle}/></td>}
                {!isDev&&<td style={{padding:"9px 10px"}}><input value={r.novoRastreio||""} onChange={e=>upd(r.id,{novoRastreio:e.target.value})} placeholder="Rastreio" style={inputStyle}/></td>}
                {!isDev&&<td style={{padding:"9px 10px"}}><input value={r.materialReenvio||""} onChange={e=>upd(r.id,{materialReenvio:e.target.value})} placeholder="Material enviado" style={inputStyle}/></td>}
                <td style={{padding:"11px 14px",color:C.text2,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={r.motivoDevolucao||r.motivo}>{r.motivoDevolucao||r.motivo||"—"}</td>
                <td style={{padding:"9px 10px"}}>
                  <select value={r[cfg.statusKey]||cfg.options[0]} onChange={e=>upd(r.id,{[cfg.statusKey]:e.target.value},{acao:`${cfg.title}: ${e.target.value}`})} disabled={!perms?.canOperate} style={inputStyle}>
                    {cfg.options.map(o=><option key={o}>{o}</option>)}
                  </select>
                </td>
                <td style={{padding:"9px 10px"}}><input value={r.responsavel||""} onChange={e=>upd(r.id,{responsavel:e.target.value})} placeholder="Responsavel" disabled={!perms?.canOperate} style={inputStyle}/></td>
                <td style={{padding:"9px 12px"}}>
                  <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                    {isDev&&perms?.canOperate&&<button onClick={()=>onCreateReenvio(r.id)} style={{background:C.amberSoft,border:`1px solid ${C.amberBorder}`,color:C.amber,borderRadius:7,padding:"6px 10px",fontSize:10,cursor:"pointer",fontWeight:600}}>Gerar reenvio</button>}
                    {perms?.canOperate&&<button onClick={()=>onResolve(r.id)} style={{background:C.gold,border:"none",color:C.white,borderRadius:7,padding:"6px 10px",fontSize:10,cursor:"pointer",fontWeight:600}}>Concluir</button>}
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  )
}

export default function App() {
  const [session,setSession]=useState(()=>{
    try {
      const raw = localStorage.getItem(SESSION_KEY)
      return raw ? JSON.parse(raw) : null
    } catch {
      return null
    }
  })
  const [perfil,setPerfil]=useState(null)
  const [nomeAtendente,setNomeAtendente]=useState("")
  const [loadingPerfil,setLoadingPerfil]=useState(false)
  const [rows,setRows]=useState([])
  const [tab,setTab]=useState(null)
  const [paste,setPaste]=useState("")
  const [importing,setImporting]=useState(false)
  const [loadingData,setLoadingData]=useState(false)
  const [initialDataLoaded,setInitialDataLoaded]=useState(false)
  const [compact,setCompact]=useState(false)
  const [navOpen,setNavOpen]=useState(false)
  const [dark,setDark]=useState(false)
  const [toasts,setToasts]=useState([])
  const [lSrch,setLSrch]=useState(""); const [lSt,setLSt]=useState("Todos")
  const [lTr,setLTr]=useState("Todos"); const [lUrg,setLUrg]=useState("Todos")
  const [lAc,setLAc]=useState("Todos"); const [lSitPrazo,setLSitPrazo]=useState("Todos"); const [qf,setQf]=useState("todos")
  const [lShowFilters,setLShowFilters]=useState(false)
  const [sResp,setSResp]=useState("Todos"); const [sShowFilters,setSShowFilters]=useState(false)
  const [lPage,setLPage]=useState(1)
  const [selIds,setSelIds]=useState(new Set())
  const [sortCol,setSortCol]=useState(null); const [sortDir,setSortDir]=useState("asc")
  const [sSrch,setSSrch]=useState(""); const [sAtend,setSAtend]=useState("Todos")
  const [sUrg,setSUrg]=useState("Todos")
  const [selSup,setSelSup]=useState(null)
  const [supView,setSupView]=useState('lista') // 'lista' | 'kanban'
  const [selSupIds,setSelSupIds]=useState(new Set())
  const [openHist,setOpenHist]=useState(false)
  const [aSrch,setASrch]=useState("")
  const [aPage,setAPage]=useState(1)
  const [syncStatus,setSyncStatus]=useState("idle")
  const [lastSync,setLastSync]=useState(null)
  const [countdown,setCountdown]=useState(10)
  const saveTimer=useRef(null); const fileRef=useRef()
  const token = session?.access_token

  const addToast = useCallback((msg,type="ok",ms=4000)=>{
    const id=Date.now(); setToasts(p=>[...p,{id,msg,type}]); setTimeout(()=>setToasts(p=>p.filter(t=>t.id!==id)),ms)
  },[])

  const loadUserProfile = useCallback(async data => {
    if (!data?.access_token) return
    setLoadingPerfil(true)
    try {
      const r = await fetch(`${SUPA_URL}/rest/v1/usuarios?id=eq.${data.user.id}&select=*`,{headers:aSH(data.access_token)})
      const arr = await r.json(); const p = arr[0]?.perfil||"leitura"
      setNomeAtendente(arr[0]?.nome||data.user?.email||"")
      setPerfil(p); setTab(PERMS[p].tabs[0])
    } catch(e) { setPerfil("leitura"); setTab("dashboard") }
    setLoadingPerfil(false)
  },[])

  const handleLogin = async data => {
    localStorage.setItem(SESSION_KEY, JSON.stringify(data))
    setSession(data)
    await loadUserProfile(data)
  }

  useEffect(()=>{
    if (session && !perfil && !loadingPerfil) loadUserProfile(session)
  },[session,perfil,loadingPerfil,loadUserProfile])

  const handleLogout = async () => {
    if (token) await signOut(token)
    localStorage.removeItem(SESSION_KEY)
    setSession(null); setPerfil(null); setRows([]); setTab(null)
  }
  const perms = perfil?PERMS[perfil]:null

  // BUG FIX #9: useEffect de carga inicial completamente reescrito
  useEffect(()=>{
    if (!token) { setInitialDataLoaded(false); return }
    setSyncStatus("loading")
    setLoadingData(true)
    setInitialDataLoaded(false)

    const fixRows = data => data.map(r=>({
      ...r, isNew:false,
      atendimento:    isEntregue(r.status)&&!r.enviadoSuporte?"Resolvido":r.atendimento,
      enviadoSuporte: isEntregue(r.status)&&!r.enviadoSuporte?false:r.enviadoSuporte,
    }))

    dbLoadFast(token, partial => {
      setRows(fixRows(partial))
      setLastSync(new Date())
      setLoadingData(false)
    }).then(data => {
      if (data.length>0) { setRows(fixRows(data)); setLastSync(new Date()) }
      setInitialDataLoaded(true)
      setSyncStatus("idle"); setLoadingData(false)
    }).catch(e => {
      setSyncStatus("error")
      addToast("Erro ao carregar: "+e.message,"error",8000)
      setInitialDataLoaded(true)
      setLoadingData(false)
    })
  },[token])

  // BUG FIX #10: polling usava dbLoad (inexistente) — corrigido para dbLoadFast
  useEffect(()=>{
    if (!token) return
    const poll = async () => {
      setCountdown(10)
      try {
        const remote = await dbLoadFast(token, ()=>{})
        if (remote.length>0) {
          let nc=0
          setRows(prev=>{
            const rm=new Map(remote.map(r=>[r.id,r]))
            const lm=new Map(prev.map(r=>[r.id,r]))
            const merged=[...rm.values()].map(r=>{
              const loc=lm.get(r.id)
              if (!loc){nc++;return{...r,isNew:true}}
              return loc.historico.length>=r.historico.length?loc:{...r,isNew:false}
            })
            prev.forEach(r=>{if (!rm.has(r.id))merged.push(r)}); return merged
          })
          if (nc>0) addToast(`${nc} pedido${nc>1?"s":""} atualizado${nc>1?"s":""} por outro usuário`,"warn")
          setLastSync(new Date())
        }
      } catch(e){}
    }
    const interval=setInterval(poll,10000)
    const cd=setInterval(()=>setCountdown(p=>p>0?p-1:10),1000)
    return ()=>{clearInterval(interval);clearInterval(cd)}
  },[token,addToast])

  useEffect(()=>{
    if (!token||rows.length===0) return
    if (saveTimer.current) clearTimeout(saveTimer.current)
    saveTimer.current=setTimeout(async()=>{
      setSyncStatus("saving")
      try { await dbUpsert(rows,token) }
      catch(e){setSyncStatus("error");addToast("Erro ao salvar: "+e.message,"error",8000);setTimeout(()=>setSyncStatus("idle"),4000);return}
      setLastSync(new Date()); setSyncStatus("saved"); setTimeout(()=>setSyncStatus("idle"),2500)
    },1200)
  },[rows,token,addToast])

  useEffect(()=>{if (!rows.some(r=>r.isNew))return;const t=setTimeout(()=>setRows(p=>p.map(r=>({...r,isNew:false}))),6000);return()=>clearTimeout(t)},[rows])
  useEffect(()=>setLPage(1),[lSrch,lSt,lTr,lUrg,lAc,lSitPrazo,qf,sortCol,sortDir])
  useEffect(()=>setAPage(1),[aSrch])
  useEffect(()=>{setSResp("Todos")},[]) // reset on mount
  const detailPanelRef = useRef(null)
  const queueRef = useRef(null)
  useEffect(()=>{
    setOpenHist(false)
    if (detailPanelRef.current) detailPanelRef.current.scrollTop = 0
    // Rola a fila para que o item selecionado fique no topo
    if (selSup && queueRef.current) {
      const el = queueRef.current.querySelector(`[data-id="${selSup}"]`)
      if (el) el.scrollIntoView({block:"start", behavior:"smooth"})
    }
    // Auto-preenche responsável com usuário logado se ainda estiver vazio
    if (selSup && nomeAtendente) {
      setRows(prev=>prev.map(r=>
        r.id===selSup && !r.responsavel ? {...r, responsavel:nomeAtendente} : r
      ))
    }
  },[selSup])

  // BUG FIX #7: doImport — removido trailing ", [rows,...])" corrompido
  const doImport = useCallback(txt=>{
    if (!perms?.canImport) return
    const parsed = parseData(txt)
    if (!parsed.length){addToast("Nenhum dado reconhecido.","error");return}
    let added=0,updated=0,skipped=0
    setRows(prev=>{
      const byNuvem=new Map(prev.map(r=>[r.nuvem,r])); const result=[...prev]
      for (const novo of parsed) {
        const existing=byNuvem.get(novo.nuvem)
        if (!existing){result.push(novo);added++}
        else if (norm(existing.status)===norm(novo.status)){
          if (isEntregue(novo.status)&&!existing.enviadoSuporte&&existing.atendimento!=="Resolvido"){
            const idx=result.findIndex(r=>r.nuvem===novo.nuvem)
            if (idx>=0){result[idx]={...existing,email:novo.email||existing.email,atendimento:"Resolvido",enviadoSuporte:false,historico:[...existing.historico,{acao:"Arquivado automaticamente — entrega concluída",ts:new Date().toLocaleString("pt-BR")}]};updated++}
          }else if (!existing.enviadoSuporte&&existing.atendimento!=="Resolvido"&&shouldAutoSendSupport(novo.status)){
            const idx=result.findIndex(r=>r.nuvem===novo.nuvem)
            if (idx>=0){
              result[idx]=applyAutoSupport({...existing,email:novo.email||existing.email,status:novo.status,urgencia:novo.urgencia,acionar:novo.acionar,motivo:novo.motivo})
              updated++
            }
          }else{
            // Mesmo status: só atualiza email se vier preenchido
            if (novo.email && !existing.email) {
              const idx=result.findIndex(r=>r.nuvem===novo.nuvem)
              if (idx>=0) result[idx]={...existing, email:novo.email}
            }
            skipped++
          }
        }else{
          const idx=result.findIndex(r=>r.nuvem===novo.nuvem)
          if (idx>=0){
            const alertaStatus=existing.enviadoSuporte&&norm(existing.status)!==norm(novo.status)?`Status atualizado: ${existing.status} → ${novo.status}`:existing.alertaStatus
            const spVal=parseStatusPrazo(novo.statusPrazoRaw)
            result[idx]={...novo,id:existing.id,obs:existing.obs,responsavel:existing.responsavel,chamado:existing.chamado,enviadoSuporte:existing.enviadoSuporte,atendimento:existing.enviadoSuporte?existing.atendimento:novo.atendimento,entregueNoPrazo:spVal!==null?spVal:novo.entregueNoPrazo,alertaStatus,historico:[...existing.historico,{acao:`Status atualizado: ${existing.status} → ${novo.status}`,ts:new Date().toLocaleString("pt-BR")}],isNew:true};updated++
            if (existing.atendimento==="Resolvido") result[idx]={...result[idx],atendimento:"Resolvido",enviadoSuporte:false}
            else result[idx]=applyAutoSupport(result[idx])
          }
        }
      }
      return result
    })
    setTimeout(()=>{
      const parts=[]
      if (added>0) parts.push(`${added} novo${added>1?"s":""}`)
      if (updated>0) parts.push(`${updated} atualizado${updated>1?"s":""}`)
      if (skipped>0) parts.push(`${skipped} ignorado${skipped>1?"s":""} (mesmo status)`)
      addToast(parts.join(" · ")||"Nenhuma alteração")
    },100)
    setPaste(""); setImporting(false)
  },[addToast,perms])

  const handleFile = e => {
    if (!perms?.canImport) return
    const f=e.target.files[0]; if (!f) return
    if (/\.(xlsx?)$/i.test(f.name)){const rd=new FileReader();rd.onload=ev=>{const wb=XLSX.read(new Uint8Array(ev.target.result),{type:"array"});doImport(XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]],{FS:";",blankrows:false}))};rd.readAsArrayBuffer(f)}
    else{const rd=new FileReader();rd.onload=ev=>doImport(ev.target.result);rd.readAsText(f,/\.csv$/i.test(f.name)?"windows-1252":"UTF-8")}
    e.target.value=""
  }

  const upd = (id,ch,hist) => setRows(prev=>prev.map(r=>{
    if (r.id!==id) return r
    const historico=hist?[...r.historico,{...hist,ts:new Date().toLocaleString("pt-BR")}]:r.historico
    return {...r,...ch,historico}
  }))
  const del = id => {if (!perms?.canDelete)return;setRows(prev=>prev.filter(r=>r.id!==id));dbDelete(id,token).catch(()=>{})}
  const toggleSel = id => setSelIds(prev=>{const n=new Set(prev);n.has(id)?n.delete(id):n.add(id);return n})
  const clearSel  = () => setSelIds(new Set())
  const bulkSend  = () => {
    if (!perms?.canSendSupport) return
    const ts=new Date().toLocaleString("pt-BR"),sentAt=new Date().toISOString()
    setRows(prev=>prev.map(r=>selIds.has(r.id)?{...r,enviadoSuporte:true,atendimento:"Aberto",sentAt,historico:[...r.historico,{acao:"Enviado ao suporte (lote)",ts}]}:r))
    addToast(`${selIds.size} pedido${selIds.size>1?"s":""} enviado${selIds.size>1?"s":""} ao suporte`); clearSel()
  }
  const toggleSelSup = id => setSelSupIds(prev=>{const n=new Set(prev);n.has(id)?n.delete(id):n.add(id);return n})
  const bulkArchive  = () => {
    if (!perms?.canOperate) return
    const ts=new Date().toLocaleString("pt-BR")
    setRows(prev=>prev.map(r=>selSupIds.has(r.id)?{...r,atendimento:"Resolvido",historico:[...r.historico,{acao:"Arquivado em lote",ts}]}:r))
    addToast(`${selSupIds.size} pedido${selSupIds.size>1?"s":""} finalizado${selSupIds.size>1?"s":""}`)
    if (selSupIds.has(selSup)) setSelSup(null); setSelSupIds(new Set())
  }
  const handleResolve   = id => {if (!perms?.canOperate)return;upd(id,{atendimento:"Resolvido"},{acao:"Atendimento finalizado",usuario:nomeAtendente});setSelSup(null);addToast("Pedido finalizado")}
  const handleCreateReenvio = id => {
    if (!perms?.canOperate) return
    const atual = rows.find(r=>r.id===id)
    upd(id,{fluxoEspecial:"reenvio",decisaoCliente:"Reenvio",motivoDevolucao:atual?.motivoDevolucao||atual?.motivo||"",reenvioStatus:"Pendente",devolucaoStatus:"Reenviar",atendimento:"Em andamento"},{acao:"Cliente optou por reenvio",usuario:nomeAtendente})
    setSelSup(null); addToast("Pedido movido para Reenvio")
  }
  const handleMarkDevolucao = id => {
    if (!perms?.canOperate) return
    const atual = rows.find(r=>r.id===id)
    upd(id,{fluxoEspecial:"devolucao",decisaoCliente:"Estorno / devolucao",motivoDevolucao:atual?.motivoDevolucao||atual?.motivo||"",devolucaoStatus:"Aguardando produto",atendimento:"Em andamento"},{acao:"Cliente optou por estorno/devolucao do produto",usuario:nomeAtendente})
    addToast("Pedido marcado em Devolucao")
  }
  const handleReturnLog = id => {if (!perms?.canOperate)return;upd(id,{enviadoSuporte:false,sentAt:null},{acao:"Devolvido à Logística",usuario:nomeAtendente});setSelSup(null)}
  const handleTotalExpressTracking = async id => {
    if (!perms?.canOperate) return
    const atual = rows.find(r=>r.id===id)
    if (!atual || !isTotalExpress(atual.transportadora)) {
      addToast("Acao disponivel apenas para pedidos Total Express","warn")
      return
    }
    const payload = totalExpressPayloadForOrder(atual)
    if (!payload) {
      addToast("Informe rastreio, pedido ou NF para consultar Total Express","error")
      return
    }
    const data = await totalExpressRequest({action:"tracking",transportadora:atual.transportadora,...payload,comprovanteEntrega:true}, token)
    const item = data.items?.[0]
    if (!item) {
      addToast("Total Express nao retornou encomenda para este pedido","warn")
      return
    }
    const ch = totalExpressUpdateFromItem(item)
    Object.keys(ch).forEach(k=>{ if (ch[k]==="" || ch[k]===undefined || ch[k]===null) delete ch[k] })
    upd(id,ch,{acao:`Total Express atualizado: ${ch.status||"sem nova ocorrencia"}`,usuario:nomeAtendente})
    addToast("Status Total Express atualizado")
  }
  const handleTotalExpressTicket = async id => {
    if (!perms?.canOperate) return
    const atual = rows.find(r=>r.id===id)
    if (!atual || !isTotalExpress(atual.transportadora)) {
      addToast("Ticket Total Express apenas para pedidos Total Express","warn")
      return
    }
    const origemId = atual.rastreio || atual.nuvem || atual.nf
    if (!origemId) {
      addToast("Informe rastreio, pedido ou NF antes de abrir ticket","error")
      return
    }
    const descricao = [
      `Pedido SG/Nuvem: ${atual.nuvem||"-"}`,
      `Cliente: ${atual.destinatario||"-"}`,
      `AWB/Rastreio: ${atual.rastreio||"-"}`,
      `NF: ${atual.nf||"-"}`,
      `Status atual: ${atual.status||"-"}`,
      `Motivo: ${atual.motivoDevolucao||atual.motivo||"Acionamento logistico"}`,
      `Responsavel interno: ${nomeAtendente||"-"}`,
    ].join("\n")
    const data = await totalExpressRequest({action:"ticket",transportadora:atual.transportadora,origem:"encomenda",origemId,descricao,casoCritico:atual.urgencia==="Alta"}, token)
    const ticketId = data?.data?.data?.id || data?.data?.id || data?.id || ""
    upd(id,{chamado:ticketId?`Total #${ticketId}`:(atual.chamado||"Ticket Total aberto"),totalExpressTicketId:ticketId},{acao:`Ticket Total Express aberto${ticketId?`: ${ticketId}`:""}`,usuario:nomeAtendente})
    addToast(ticketId?`Ticket Total #${ticketId} aberto`:"Ticket Total Express aberto")
  }
  const handleClearAll  = () => {
    if (!perms?.canClear) return
    if (!window.confirm("Isso removerá TODOS os pedidos da base de dados. Esta ação não pode ser desfeita. Confirmar?")) return
    setRows([]); dbClear(token).catch(()=>{}); addToast("Todos os dados foram removidos","warn")
  }
  const handleArchiveFromLog = id => {
    if (!perms?.canOperate) return
    upd(id, {atendimento:"Resolvido", enviadoSuporte:false}, {acao:"Arquivado pela Logística — entrega confirmada", usuario:nomeAtendente})
    addToast("Pedido finalizado")
  }
  const bulkArchiveFromLog = () => {
    if (!perms?.canOperate) return
    const ts = new Date().toLocaleString("pt-BR")
    setRows(prev=>prev.map(r=>selIds.has(r.id)?{...r,atendimento:"Resolvido",enviadoSuporte:false,historico:[...r.historico,{acao:"Arquivado em lote pela Logística",ts,usuario:nomeAtendente}]}:r))
    addToast(`${selIds.size} pedido${selIds.size>1?"s":""} finalizado${selIds.size>1?"s":""}`); clearSel()
  }
  const toggleSort = col => {if (sortCol===col)setSortDir(d=>d==="asc"?"desc":"asc");else{setSortCol(col);setSortDir("asc")}}

  const baseLog  = rows.filter(r=>!r.enviadoSuporte&&r.atendimento!=="Resolvido")
  const baseSup  = rows.filter(r=>r.enviadoSuporte&&r.atendimento!=="Resolvido"&&!r.fluxoEspecial)
  const baseDev  = rows.filter(r=>r.atendimento!=="Resolvido"&&r.fluxoEspecial!=="reenvio"&&(r.fluxoEspecial==="devolucao"||classificarProblema(r)==="DEVOLUCAO"))
  const baseReenvio = rows.filter(r=>r.atendimento!=="Resolvido"&&r.fluxoEspecial==="reenvio")
  const baseArch = rows.filter(r=>r.atendimento==="Resolvido")
  const prioridadeRows = baseLog.filter(r=>prioridadeOperacional(r).level!=="normal")
  const criticaRows = baseLog.filter(r=>prioridadeOperacional(r).level==="critica")
  const detail   = selSup?baseSup.find(r=>r.id===selSup):null
  const qCounts  = Object.fromEntries(QFILTERS.map(f=>[f.id,applyQF(baseLog,f.id).length]))
  const filteredLog = applySortRows(applyQF(baseLog,qf).filter(r=>{
    const q=lSrch.toLowerCase()
    return (!q||[r.nuvem,r.destinatario,r.transportadora,r.rastreio,r.status,r.motivo].some(v=>(v||"").toLowerCase().includes(q)))
      &&(lSt==="Todos"||r.status===lSt)&&(lTr==="Todos"||r.transportadora===lTr)
      &&(lUrg==="Todos"||r.urgencia===lUrg)&&(lAc==="Todos"||r.acionar===lAc)
      &&(lSitPrazo==="Todos"||(()=>{
        const dt=parsePrazo(r.prazo); if(!dt) return false
        const h=new Date(); h.setHours(0,0,0,0)
        const d=Math.ceil((dt-h)/86400000)
        if(lSitPrazo==="Atraso") return d<0
        if(lSitPrazo==="No Prazo") return d===0
        if(lSitPrazo==="Antes do Prazo") return d>0
        return true
      })())
  }),sortCol,sortDir)
  const totalPages = Math.max(1,Math.ceil(filteredLog.length/PAGE_SIZE))
  const safeP      = Math.min(lPage,totalPages)
  const pagedLog   = filteredLog.slice((safeP-1)*PAGE_SIZE,safeP*PAGE_SIZE)
  const respOpts = uniq(baseSup.map(r=>r.responsavel).filter(Boolean))
  const supRows = baseSup.filter(r=>{const q=sSrch.toLowerCase();return(!q||[r.nuvem,r.destinatario,r.rastreio,r.status].some(v=>(v||"").toLowerCase().includes(q)))&&(sAtend==="Todos"||r.atendimento===sAtend)&&(sUrg==="Todos"||r.urgencia===sUrg)&&(sResp==="Todos"||r.responsavel===sResp)}).sort((a,b)=>{const uo={Alta:0,Média:1,Baixa:2,"—":3},ao={Aberto:0,"Em andamento":1};return(uo[a.urgencia]-uo[b.urgencia])||(ao[a.atendimento]-ao[b.atendimento])})
  const archRows = baseArch.filter(r=>{const q=aSrch.toLowerCase();return!q||[r.nuvem,r.destinatario,r.transportadora,r.status].some(v=>(v||"").toLowerCase().includes(q))}).sort((a,b)=>{const ta=(a.historico.find(h=>h.acao&&(h.acao.includes("Resolvido")||h.acao.includes("Arquivado")))||{}).ts||"";const tb=(b.historico.find(h=>h.acao&&(h.acao.includes("Resolvido")||h.acao.includes("Arquivado")))||{}).ts||"";return tb.localeCompare(ta)})

  const stOpts=uniq(baseLog.map(r=>r.status)), trOpts=uniq(baseLog.map(r=>r.transportadora))
  const st={log:baseLog.length,alta:baseLog.filter(r=>r.urgencia==="Alta").length,acionar:baseLog.filter(r=>r.acionar==="Sim").length}
  const ss={total:baseSup.length,abertos:baseSup.filter(r=>r.atendimento==="Aberto").length,andamento:baseSup.filter(r=>r.atendimento==="Em andamento").length}
  const devStats={total:baseDev.length,pendentes:baseDev.filter(r=>(r.devolucaoStatus||"Aguardando tratativa")==="Aguardando tratativa").length}
  const reenvStats={total:baseReenvio.length,pendentes:baseReenvio.filter(r=>(r.reenvioStatus||"Pendente")==="Pendente").length}
  const arch=baseArch.length

  const entregues  = rows.filter(r=>isEntregue(r.status))
  const hoje       = new Date(); hoje.setHours(0,0,0,0)
  const parados    = baseLog.filter(r=>{const d=diasSemMov(r.ultimaMov);return d!==null&&d>=ALERTA_DIAS}).length

  // Calcula entregueNoPrazo on-the-fly (não confia no valor salvo, evita bug de import antigo)
  const calcNoPrazoLive = r => {
    const spVal = parseStatusPrazo(r.statusPrazoRaw)
    if (spVal !== null) return spVal
    if (!isEntregue(r.status)) return null
    const dtEntrega = parsePrazo(r.ultimaMov)
    const dtPrazo   = parsePrazo(r.prazo)
    if (!dtEntrega || !dtPrazo) return null
    const de = new Date(dtEntrega.getFullYear(), dtEntrega.getMonth(), dtEntrega.getDate())
    const dp = new Date(dtPrazo.getFullYear(),   dtPrazo.getMonth(),   dtPrazo.getDate())
    return de <= dp
  }

  const noPrazo    = entregues.filter(r=>calcNoPrazoLive(r)===true).length
  const pctNoPrazo = entregues.length>0?Math.round((noPrazo/entregues.length)*100):0

  // BUG FIX #11: const trStats={} estava faltando; BUG FIX #13: forEach sem trailing lixo
  const TRANSP_INVALIDAS = new Set(["SP","RJ","MG","RS","SC","PR","BA","GO","PE","CE","AM","PA","MT","MS","ES","RN","PI","AL","SE","TO","RO","AC","AP","RR","MA","PB","DF"])

  const trStats={}
  rows.forEach(r=>{
    const tr = (r.transportadora||"").trim()
    if (!tr || tr.length <= 2 || TRANSP_INVALIDAS.has(tr.toUpperCase())) return
    if (!trStats[tr]) trStats[tr]={total:0,entregues:0,noPrazo:0,foraPrazo:0,vencidos:0}
    const s=trStats[tr]; s.total++
    if (isEntregue(r.status)){
      s.entregues++
      const np = calcNoPrazoLive(r)
      if (np===true)  s.noPrazo++
      if (np===false) s.foraPrazo++
    } else {
      const d=parsePrazo(r.prazo); if(d&&d<hoje) s.vencidos++
    }
  })
  const trData    = Object.entries(trStats).map(([name,s])=>({name,total:s.total,entregues:s.entregues,noPrazo:s.noPrazo,foraPrazo:s.foraPrazo,vencidos:s.vencidos,pct:s.entregues>0?Math.round((s.noPrazo/s.entregues)*100):0})).sort((a,b)=>b.total-a.total).slice(0,8)
  const trBarData = trData.map(t=>({name:t.name,"No prazo":t.noPrazo,"Fora prazo":t.foraPrazo,"Vencidos":t.vencidos}))

  const ufStats={}
  rows.forEach(r=>{
    const uf=(r.uf||"").toUpperCase().trim(); if (!uf||uf.length>3) return
    if (!ufStats[uf]) ufStats[uf]={total:0,entregues:0,noPrazo:0}
    const s=ufStats[uf]; s.total++
    if (isEntregue(r.status)){s.entregues++;if(calcNoPrazoLive(r)===true)s.noPrazo++}
  })
  const ufData     = Object.entries(ufStats).map(([uf,s])=>({uf,total:s.total,entregues:s.entregues,noPrazo:s.noPrazo,pct:s.entregues>0?Math.round((s.noPrazo/s.entregues)*100):0})).sort((a,b)=>b.total-a.total).slice(0,15)
  const urgData    = ["Alta","Média","Baixa"].map(u=>({name:u,value:baseLog.filter(r=>r.urgencia===u).length,fill:urgStyles[u].dot})).filter(d=>d.value>0)
  const statusMap  = {}; rows.filter(r=>!isEntregue(r.status)).forEach(r=>{if(r.status)statusMap[r.status]=(statusMap[r.status]||0)+1})
  const statusData = Object.entries(statusMap).sort((a,b)=>b[1]-a[1]).slice(0,8).map(([name,value])=>({name,value}))

  if (!session) return <LoginScreen onLogin={handleLogin}/>
  if (loadingPerfil || !perfil) return (
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",minHeight:"100vh",background:C.brand,color:"#F7F6F1",fontSize:10,letterSpacing:"0.18em",textTransform:"uppercase"}}>
      Carregando central
    </div>
  )

  const showImp=importing&&perms?.tabs.some(t=>["logistica","dashboard"].includes(t))
  const pd=compact?5:9
  const PERFLABEL={admin:"Admin",logistica:"Logística",suporte:"Suporte",leitura:"Leitura"}
  const syncDot  = syncStatus==="error"?C.red:syncStatus==="saving"?C.gold:syncStatus==="saved"?"#27ae60":"#555"
  const syncText = syncStatus==="loading"?"Carregando...":syncStatus==="saving"?"Salvando...":syncStatus==="saved"?"Sincronizado ✓":syncStatus==="error"?"Erro":lastSync?`Sync em ${countdown}s`:""
  const TABS=[{key:"dashboard",label:"Dashboard",badge:null},{key:"logistica",label:"Logística",badge:st.acionar>0?st.acionar:null},{key:"suporte",label:"Suporte",badge:ss.abertos>0?ss.abertos:null},{key:"devolucao",label:"Devolução",badge:devStats.total>0?devStats.total:null},{key:"reenvio",label:"Reenvio",badge:reenvStats.pendentes>0?reenvStats.pendentes:null},{key:"arquivados",label:"Finalizados",badge:arch>0?arch:null},{key:"usuarios",label:"Usuários",badge:null}].filter(t=>perms?.tabs.includes(t.key))
  const TH  = {padding:`${compact?8:11}px 14px`,textAlign:"left",color:C.gold,fontWeight:400,fontSize:9,letterSpacing:"0.14em",textTransform:"uppercase",borderBottom:`1px solid #2A2A2A`,whiteSpace:"nowrap",background:C.brand,position:"sticky",top:0,zIndex:5,cursor:"pointer"}
  const THF = {...TH,cursor:"default"}

  return (
    <div style={{fontFamily:"'Inter',sans-serif",minHeight:"100vh",background:C.cream,color:C.text1,transition:"background .3s",paddingLeft:!showImp?(navOpen?254:52):0}}>
      <style>{getGlobalStyle()}</style>
      <Toast toasts={toasts}/>

      {/* ── HEADER ── */}
      <div style={{background:C.brand,padding:"0 28px",display:"flex",alignItems:"stretch",justifyContent:"space-between",borderBottom:`1px solid ${C.borderDark}`}}>
        <div style={{display:"flex",alignItems:"center",gap:22,padding:"12px 0"}}>
          <div style={{minWidth:184}}>
            <SGWordmark dark size={17}/>
            <div style={{fontSize:9,color:"#BDBDBD",textTransform:"uppercase",marginTop:7,fontWeight:700,letterSpacing:"0.16em"}}>Central de Pedidos</div>
          </div>
          <div style={{width:1,height:38,background:"#2B2B2B"}}/>
          <div style={{display:"flex",alignItems:"center",gap:7}}>
            <span style={{width:6,height:6,borderRadius:"50%",background:syncDot,display:"inline-block",boxShadow:`0 0 6px ${syncDot}66`}}/>
            <span style={{fontSize:10,color:syncDot,letterSpacing:"0.04em"}}>{syncText}</span>
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          {perms?.canImport&&!showImp&&(
            <>
              <button onClick={()=>setCompact(c=>!c)} style={{background:"transparent",border:`1px solid ${C.gold}44`,color:compact?C.gold:`${C.gold}88`,borderRadius:6,padding:"6px 12px",fontSize:10,cursor:"pointer",letterSpacing:"0.08em"}}>{compact?"⊞":"⊟"}</button>
              <button onClick={()=>{const hf=lSrch||lSt!=="Todos"||lTr!=="Todos"||lUrg!=="Todos"||lAc!=="Todos";exportCSV(hf&&tab==="logistica"?filteredLog:rows)}} style={{background:"transparent",border:`1px solid #444`,color:"#888",borderRadius:6,padding:"6px 14px",fontSize:10,cursor:"pointer",letterSpacing:"0.06em"}}>↓ Exportar</button>
              <button onClick={()=>setImporting(true)} style={{background:"transparent",border:`1px solid ${C.gold}`,color:C.gold,borderRadius:6,padding:"6px 14px",fontSize:10,cursor:"pointer",letterSpacing:"0.06em"}}>+ Importar</button>
            </>
          )}
          <button onClick={()=>{const nd=!dark;applyTheme(nd);setDark(nd)}} title="Alternar modo escuro" style={{background:"transparent",border:`1px solid #333`,color:dark?C.gold:C.text3,borderRadius:6,padding:"6px 10px",fontSize:12,cursor:"pointer"}}>{dark?"☀":"🌙"}</button>
        {perms?.canClear&&rows.length>0&&(
            <button onClick={handleClearAll} style={{background:"transparent",border:`1px solid #333`,color:"#555",borderRadius:6,padding:"6px 14px",fontSize:10,cursor:"pointer",letterSpacing:"0.06em"}}>
              Limpar tudo
            </button>
          )}
          <div style={{width:1,height:28,background:"#2A2A2A"}}/>
          <div style={{textAlign:"right"}}>
            <div style={{color:C.white,fontSize:11,letterSpacing:"0.02em"}}>{session.user?.email}</div>
            <div style={{color:C.gold,fontSize:9,letterSpacing:"0.1em",textTransform:"uppercase"}}>{PERFLABEL[perfil]||perfil}</div>
          </div>
          <button onClick={handleLogout} style={{background:"transparent",border:`1px solid #2A2A2A`,color:"#666",borderRadius:6,padding:"6px 12px",fontSize:10,cursor:"pointer",letterSpacing:"0.06em"}}>Sair</button>
        </div>
        <input ref={fileRef} type="file" accept=".csv,.txt,.tsv,.xls,.xlsx" style={{display:"none"}} onChange={handleFile}/>
      </div>

      {/* ── NAV ── */}
      {!showImp&&(
        <div style={{position:"fixed",left:0,top:0,bottom:0,width:navOpen?254:52,background:C.white,borderRight:`1px solid ${C.border}`,zIndex:50,display:"flex",flexDirection:"column",transition:"width .2s",boxShadow:shadow.sm}}>
          <button onClick={()=>setNavOpen(v=>!v)} title={navOpen?"Recolher menu":"Expandir menu"} style={{height:54,background:"transparent",border:"none",borderBottom:`1px solid ${C.border}`,color:C.text2,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>
            <svg width="22" height="22" viewBox="0 0 24 24"><path d={navOpen?"M6 6l12 12M18 6 6 18":"M4 7h16M4 12h16M4 17h16"} stroke="currentColor" strokeWidth="1.8" fill="none" strokeLinecap="round"/></svg>
          </button>
          <div style={{padding:navOpen?"10px 10px":"10px 5px",display:"flex",flexDirection:"column",gap:4}}>
          {TABS.map(t=>(
            <button key={t.key} onClick={()=>{setTab(t.key);if(t.key!=="suporte")setSelSup(null)}} title={t.label}
              style={{background:tab===t.key?C.creamDark:"transparent",border:"none",color:tab===t.key?C.text1:C.text2,padding:navOpen?"10px 12px":"10px 0",cursor:"pointer",fontSize:15,fontWeight:600,display:"flex",alignItems:"center",justifyContent:navOpen?"flex-start":"center",gap:12,transition:"all .2s",whiteSpace:"nowrap",borderRadius:7,position:"relative"}}>
              <span style={{width:24,display:"flex",alignItems:"center",justifyContent:"center"}}><NavIcon type={t.key} active={tab===t.key}/></span>
              {navOpen&&<span>{t.label}</span>}
              {t.badge!=null&&<span style={{marginLeft:navOpen?"auto":0,position:navOpen?"static":"absolute",top:4,right:3,background:C.red,color:C.white,borderRadius:6,padding:"1px 6px",fontSize:9,fontWeight:800}}>{t.badge}</span>}
            </button>
          ))}
          </div>
        </div>
      )}

      {/* ── IMPORT ── */}
      {showImp&&perms?.canImport&&(
        <div style={{padding:48,maxWidth:640,margin:"0 auto"}}>
          <div style={{textAlign:"center",marginBottom:36}}>
            <div style={{fontSize:9,letterSpacing:"0.22em",textTransform:"uppercase",color:C.gold,marginBottom:10}}>{importing?"Adicionar dados":"Bem-vindo"}</div>
            <div style={{fontSize:28,fontWeight:800,color:C.text1,marginBottom:8}}>Importe seus dados</div>
            <div style={{color:C.text3,fontSize:12,lineHeight:1.6}}>Aceita .csv (cp1252), .xls, .xlsx ou colagem direta do Excel</div>
          </div>
          <div style={{background:C.white,borderRadius:14,border:`1px solid ${C.border}`,padding:24,marginBottom:16,boxShadow:shadow.sm}}>
            <div style={{fontSize:9,color:C.text3,marginBottom:14,fontWeight:500,letterSpacing:"0.12em",textTransform:"uppercase"}}>Mapeamento automático de colunas</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
              {[["No NUVEM","Identificador Ecommerce"],["Destinatário","Destinatário Nome"],["Transportadora","Estratégia de Frete"],["Cód. Rastreio","Rastreador Last Mile"],["Status","Situação"],["Prazo","Prazo Logístico"],["Cidade / UF","Destinatário Cidade / UF"],["Status Prazo","Status Prazo"],["CEP","Destinatário CEP"],["Data Envio","Data Criação Envio"],["Email","E-mail Destinatário / Email Cliente"]].map(([c,a])=>(
                <div key={c} style={{background:C.cream,border:`1px solid ${C.border}`,borderRadius:8,padding:"8px 12px"}}>
                  <div style={{color:C.gold,fontWeight:500,marginBottom:2,fontSize:11}}>{c}</div>
                  <div style={{color:C.text4,fontSize:10}}>{a}</div>
                </div>
              ))}
            </div>
          </div>
          <textarea value={paste} onChange={e=>setPaste(e.target.value)} placeholder="Cole aqui os dados copiados do sistema..."
            style={{width:"100%",minHeight:120,borderRadius:10,border:`1px solid ${C.border}`,padding:14,fontSize:12,resize:"vertical",fontFamily:"monospace",boxSizing:"border-box",background:C.white,color:C.text1,lineHeight:1.6}}/>
          <div style={{display:"flex",gap:10,marginTop:12}}>
            <button onClick={()=>doImport(paste)} disabled={!paste.trim()} style={{flex:1,background:paste.trim()?C.brand:"#ccc",border:"none",color:C.white,borderRadius:8,padding:"12px 0",fontSize:11,fontWeight:500,cursor:paste.trim()?"pointer":"not-allowed",letterSpacing:"0.12em",textTransform:"uppercase"}}>Importar dados colados</button>
            <button onClick={()=>fileRef.current.click()} style={{flex:1,background:"transparent",border:`1px solid ${C.gold}`,color:C.gold,borderRadius:8,padding:"12px 0",fontSize:11,cursor:"pointer",letterSpacing:"0.08em"}}>Importar arquivo</button>
          </div>
          <div style={{textAlign:"center",marginTop:16,display:"flex",justifyContent:"center",gap:20}}>
            <button onClick={()=>doImport(SAMPLE)} style={{background:"transparent",border:"none",color:C.text4,fontSize:11,cursor:"pointer",textDecoration:"underline"}}>Carregar exemplo</button>
            {importing&&<button onClick={()=>setImporting(false)} style={{background:"transparent",border:"none",color:C.text4,fontSize:11,cursor:"pointer"}}>Cancelar</button>}
          </div>
        </div>
      )}

      {/* ── USUÁRIOS ── */}
      {tab==="usuarios"&&perfil==="admin"&&<UsuariosPanel token={token} addToast={addToast}/> }

      {/* ── DASHBOARD ── */}
      {tab==="dashboard"&&!showImp&&(
        <div style={{padding:"28px 40px"}}>
          <div style={{marginBottom:24}}>
            <div style={{fontSize:9,letterSpacing:"0.18em",textTransform:"uppercase",color:C.gold,marginBottom:4}}>Visão geral</div>
            <div style={{fontSize:22,fontWeight:800,color:C.text1}}>Dashboard Operacional</div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:12,marginBottom:20}}>
            {[
              {label:"Total de pedidos",        val:rows.length,     sub:"na base de dados"},
              {label:"Em logística",             val:st.log,          sub:`${st.acionar} acionam suporte`},
              {label:"No suporte",               val:ss.total,        sub:`${ss.abertos} abertos`,accent:ss.abertos>0},
              {label:`Parados +${ALERTA_DIAS}d`, val:parados,         sub:"sem movimentação",accent:parados>0},
              {label:"Entrega no prazo",         val:`${pctNoPrazo}%`,sub:`${noPrazo} de ${entregues.length} entregues`},
            ].map(k=><KpiCard key={k.label} {...k}/>)}
          </div>
          {rows.length===0?<div style={{textAlign:"center",padding:64,color:C.text4}}>Importe dados para visualizar os gráficos</div>:(
            <div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
                <div style={{background:C.white,borderRadius:14,border:`1px solid ${C.border}`,padding:20,boxShadow:shadow.sm}}>
                  <div style={{fontSize:9,fontWeight:500,color:C.text3,marginBottom:2,letterSpacing:"0.12em",textTransform:"uppercase"}}>Desempenho por transportadora</div>
                  <div style={{fontSize:9,color:C.text4,marginBottom:14}}>No prazo · Fora do prazo · Vencidos</div>
                  <ResponsiveContainer width="100%" height={220}>
                    <BarChart data={trBarData} layout="vertical" margin={{left:0,right:16,top:0,bottom:0}}>
                      <XAxis type="number" tick={{fontSize:9,fill:C.text4}} axisLine={false} tickLine={false}/>
                      <YAxis type="category" dataKey="name" width={90} tick={{fontSize:10,fill:C.text3}} axisLine={false} tickLine={false}/>
                      <Tooltip contentStyle={{fontSize:11,border:`1px solid ${C.border}`,borderRadius:8,boxShadow:shadow.md}}/>
                      <Bar dataKey="No prazo"   stackId="a" fill={C.green} name="No prazo"/>
                      <Bar dataKey="Fora prazo" stackId="a" fill={C.amber} name="Fora prazo"/>
                      <Bar dataKey="Vencidos"   stackId="a" fill={C.red}   name="Vencidos" radius={[0,4,4,0]}/>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
                <div style={{background:C.white,borderRadius:14,border:`1px solid ${C.border}`,padding:20,boxShadow:shadow.sm}}>
                  <div style={{fontSize:9,fontWeight:500,color:C.text3,marginBottom:14,letterSpacing:"0.12em",textTransform:"uppercase"}}>Urgência — pedidos em aberto</div>
                  <ResponsiveContainer width="100%" height={180}>
                    <PieChart>
                      <Pie data={urgData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={80} innerRadius={48} paddingAngle={3}>
                        {urgData.map((e,i)=><Cell key={i} fill={e.fill}/>)}
                      </Pie>
                      <Tooltip contentStyle={{fontSize:11,border:`1px solid ${C.border}`,borderRadius:8}}/>
                    </PieChart>
                  </ResponsiveContainer>
                  <div style={{display:"flex",justifyContent:"center",gap:14,flexWrap:"wrap"}}>
                    {urgData.map(e=><span key={e.name} style={{fontSize:10,color:C.text3,display:"flex",alignItems:"center",gap:5}}><span style={{width:8,height:8,borderRadius:"50%",background:e.fill,display:"inline-block"}}/>{e.name} ({e.value})</span>)}
                  </div>
                </div>
              </div>
              {ufData.length>0&&(
                <div style={{background:C.white,borderRadius:14,border:`1px solid ${C.border}`,overflow:"hidden",boxShadow:shadow.sm,marginBottom:16}}>
                  <div style={{padding:"18px 24px",borderBottom:`1px solid ${C.border}`,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                    <div style={{fontSize:9,fontWeight:500,color:C.text3,letterSpacing:"0.12em",textTransform:"uppercase"}}>Desempenho por estado (UF)</div>
                    <div style={{fontSize:10,color:C.text4}}>{ufData.length} estados com pedidos</div>
                  </div>
                  <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:0}}>
                    {ufData.map((u,i)=>(
                      <div key={u.uf} style={{padding:"14px 16px",borderRight:i%5!==4?`1px solid ${C.border}`:"none",borderBottom:i<ufData.length-5?`1px solid ${C.border}`:"none",background:i%2===0?C.white:C.cream}}>
                        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:6}}>
                          <span style={{fontSize:14,fontWeight:800,color:C.text1}}>{u.uf}</span>
                          <span style={{fontSize:10,fontWeight:600,color:u.pct>=80?C.green:u.pct>=60?C.amber:u.pct>0?C.red:C.text4}}>{u.pct>0?`${u.pct}%`:"—"}</span>
                        </div>
                        <div style={{fontSize:10,color:C.text3,marginBottom:4}}>{u.total} pedidos</div>
                        <div style={{height:4,background:C.creamDark,borderRadius:2,overflow:"hidden"}}>
                          <div style={{height:"100%",width:`${u.pct}%`,background:u.pct>=80?C.green:u.pct>=60?C.amber:C.red,borderRadius:2}}/>
                        </div>
                        <div style={{fontSize:9,color:C.text4,marginTop:4}}>{u.noPrazo} no prazo · {u.entregues} entregues</div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
              <div style={{background:C.white,borderRadius:14,border:`1px solid ${C.border}`,overflow:"hidden",boxShadow:shadow.sm,marginBottom:16}}>
                <div style={{padding:"18px 24px",borderBottom:`1px solid ${C.border}`}}>
                  <div style={{fontSize:9,fontWeight:500,color:C.text3,letterSpacing:"0.12em",textTransform:"uppercase"}}>Ranking de transportadoras — taxa de entrega no prazo</div>
                </div>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead>
                    <tr style={{background:C.cream}}>
                      {["Transportadora","Total","Entregues","No prazo","Fora prazo","Vencidos","Taxa SLA"].map(h=><th key={h} style={{padding:"9px 18px",textAlign:"left",fontSize:9,color:C.text3,fontWeight:500,letterSpacing:"0.1em",textTransform:"uppercase",borderBottom:`1px solid ${C.border}`}}>{h}</th>)}
                    </tr>
                  </thead>
                  <tbody>
                    {trData.map((t,i)=>(
                      <tr key={t.name} style={{background:i%2===0?C.white:C.cream,borderBottom:`1px solid ${C.border}88`}}>
                        <td style={{padding:"11px 18px",fontWeight:500,color:C.text1}}>{t.name}</td>
                        <td style={{padding:"11px 18px",color:C.text2}}>{t.total}</td>
                        <td style={{padding:"11px 18px",color:C.text2}}>{t.entregues}</td>
                        <td style={{padding:"11px 18px",color:C.green,fontWeight:500}}>{t.noPrazo}</td>
                        <td style={{padding:"11px 18px",color:C.amber,fontWeight:500}}>{t.foraPrazo}</td>
                        <td style={{padding:"11px 18px",color:C.red}}>{t.vencidos}</td>
                        <td style={{padding:"11px 18px"}}>
                          <div style={{display:"flex",alignItems:"center",gap:10}}>
                            <div style={{flex:1,height:5,background:C.creamDark,borderRadius:3,overflow:"hidden"}}>
                              <div style={{height:"100%",width:`${t.pct}%`,background:t.pct>=80?C.green:t.pct>=60?C.amber:C.red,borderRadius:3,transition:"width .6s"}}/>
                            </div>
                            <span style={{fontSize:11,fontWeight:600,color:t.pct>=80?C.green:t.pct>=60?C.amber:C.red,minWidth:34}}>{t.pct}%</span>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div style={{background:C.white,borderRadius:14,border:`1px solid ${C.border}`,padding:20,boxShadow:shadow.sm}}>
                <div style={{fontSize:9,fontWeight:500,color:C.text3,marginBottom:4,letterSpacing:"0.12em",textTransform:"uppercase"}}>Status dos pedidos em aberto</div>
                <ResponsiveContainer width="100%" height={Math.max(120,statusData.length*36)}>
                  <BarChart data={statusData} layout="vertical" margin={{left:0,right:24,top:4,bottom:0}}>
                    <XAxis type="number" tick={{fontSize:9,fill:C.text4}} axisLine={false} tickLine={false}/>
                    <YAxis type="category" dataKey="name" width={140} tick={{fontSize:10,fill:C.text3}} axisLine={false} tickLine={false}/>
                    <Tooltip contentStyle={{fontSize:11,border:`1px solid ${C.border}`,borderRadius:8}}/>
                    <Bar dataKey="value" fill={C.gold} radius={[0,5,5,0]} name="Pedidos"/>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          )}
        </div>
      )}

      {/* ── LOGÍSTICA ── */}
      {tab==="logistica"&&!showImp&&(
        <div style={{padding:"24px 32px"}}>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:20}}>
            <KpiCard label="Em logística"    val={st.log}/>
            <KpiCard label="Urgência alta"   val={st.alta}    accent={st.alta>0}/>
            <KpiCard label="CrÃ­ticos agora" val={criticaRows.length} accent={criticaRows.length>0}/>
            <KpiCard label="Acionar suporte" val={st.acionar} accent={st.acionar>0}/>
          </div>
          {prioridadeRows.length>0&&(
            <div style={{background:C.brand,border:`1px solid ${C.brand}`,borderLeft:`4px solid ${C.red}`,color:C.white,padding:"12px 16px",marginBottom:14,display:"flex",alignItems:"center",gap:12,boxShadow:shadow.md}}>
              <div style={{fontSize:11,fontWeight:800,textTransform:"uppercase"}}>{prioridadeRows.length} pedidos exigem tratativa</div>
              <div style={{fontSize:11,color:"#BDBDBD",flex:1}}>Prioridade por prazo vencido, extravio, devolucao, urgencia alta ou falta de movimentacao.</div>
              <button onClick={()=>{setQf("urgente");clearSel()}} style={{background:C.white,border:"none",color:C.brand,borderRadius:4,padding:"7px 12px",fontSize:10,fontWeight:800,cursor:"pointer"}}>Ver urgentes</button>
            </div>
          )}
          <div style={{display:"flex",gap:6,marginBottom:14,flexWrap:"wrap"}}>
            {QFILTERS.map(f=>(
              <button key={f.id} onClick={()=>{setQf(f.id);clearSel()}}
                style={{background:qf===f.id?C.brand:C.white,border:`1px solid ${qf===f.id?C.brand:C.border}`,color:qf===f.id?C.white:C.text2,borderRadius:20,padding:"5px 14px",fontSize:10,cursor:"pointer",fontWeight:qf===f.id?500:400,letterSpacing:"0.06em",boxShadow:qf===f.id?"none":shadow.sm,transition:"all .2s"}}>
                {f.label}{f.id!=="todos"?` (${qCounts[f.id]||0})`:""}
              </button>
            ))}
          </div>
          <FilterBar
            search={lSrch} onSearch={setLSrch}
            showFilters={lShowFilters} onToggleFilters={()=>setLShowFilters(v=>!v)}
            filters={[
              {key:"lSt",  label:"Status",        value:lSt,        setValue:setLSt,        opts:stOpts},
              {key:"lTr",  label:"Transportadora", value:lTr,        setValue:setLTr,        opts:trOpts},
              {key:"lUrg", label:"Urgência",       value:lUrg,       setValue:setLUrg,       opts:["Todos","Alta","Média","Baixa"]},
              {key:"lSit", label:"Situação prazo", value:lSitPrazo,  setValue:setLSitPrazo,  opts:["Todos","Antes do Prazo","No Prazo","Atraso"]},
              {key:"lAc",  label:"Acionar?",       value:lAc,        setValue:setLAc,        opts:["Todos","Sim","Avaliar","Não"]},
            ]}
            onClearAll={()=>{setLSt("Todos");setLTr("Todos");setLUrg("Todos");setLSitPrazo("Todos");setLAc("Todos");setLSrch("") ;setLShowFilters(false)}}
          />
          {perms?.canSendSupport&&selIds.size>0&&(
            <div style={{background:C.brand,borderRadius:10,padding:"12px 20px",marginBottom:14,display:"flex",alignItems:"center",gap:10,boxShadow:shadow.md}}>
              <span style={{color:"#888",fontSize:12,flex:1}}>{selIds.size} pedido{selIds.size>1?"s":""} selecionado{selIds.size>1?"s":""}</span>
              <button onClick={bulkSend} style={{background:C.gold,border:"none",color:C.white,borderRadius:7,padding:"8px 18px",fontSize:11,cursor:"pointer",fontWeight:500,letterSpacing:"0.08em"}}>Enviar ao Suporte ({selIds.size})</button>
              {perms?.canOperate&&<button onClick={bulkArchiveFromLog} style={{background:C.green,border:"none",color:C.white,borderRadius:7,padding:"8px 18px",fontSize:11,cursor:"pointer",fontWeight:500,letterSpacing:"0.08em"}}>✓ Finalizar ({selIds.size})</button>}
              <button onClick={clearSel} style={{background:"transparent",border:`1px solid #333`,color:"#666",borderRadius:7,padding:"8px 14px",fontSize:11,cursor:"pointer"}}>Cancelar</button>
            </div>
          )}
          <div style={{overflowX:"auto",overflowY:"auto",maxHeight:"54vh",borderRadius:12,border:`1px solid ${C.border}`,boxShadow:shadow.sm}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:compact?11:12,tableLayout:"fixed",minWidth:1100}}>
              <colgroup>
                <col style={{width:36}}/>
                <col style={{width:92}}/>
                <col style={{width:155}}/>
                <col style={{width:130}}/>
                <col style={{width:135}}/>
                <col style={{width:92}}/>
                <col style={{width:130}}/>
                <col style={{width:80}}/>
                <col style={{width:145}}/>
                <col style={{width:118}}/>
                <col style={{width:36}}/>
              </colgroup>
              <thead>
                <tr>
                  <th style={THF}>{perms?.canSendSupport&&<input type="checkbox" onChange={e=>e.target.checked?setSelIds(new Set(pagedLog.map(r=>r.id))):clearSel()} checked={selIds.size>0&&pagedLog.every(r=>selIds.has(r.id))} style={{cursor:"pointer",accentColor:C.gold}}/>}</th>
                  {[["nuvem","No NUVEM"],["destinatario","Destinatário"],["transportadora","Transportadora"],["status","Status"],["prazo","Prazo Logístico"],["situacaoPrazo","Situação Prazo"],["urgencia","Urgência"],["ultimaMov","Últ. Movimentação"]].map(([col,label])=>(
                    <th key={col} onClick={()=>toggleSort(col)} style={TH}>{label}<SortIcon col={col} sortCol={sortCol} sortDir={sortDir}/></th>
                  ))}
                  <th style={THF}>Ação</th>
                  <th style={THF}/>
                </tr>
              </thead>
              <tbody>
                {pagedLog.length===0?<tr><td colSpan={11} style={{textAlign:"center",padding:36,color:C.text4}}>Nenhum pedido encontrado</td></tr>
                :pagedLog.map((r,i)=>{
                  const pr = prioridadeOperacional(r)
                  return (
                  <tr key={r.id} style={{background:pr.level!=="normal"?pr.bg:r.isNew?`${C.gold}14`:i%2===0?C.white:C.cream,borderBottom:`1px solid ${pr.level!=="normal"?pr.bd:C.border}66`,borderLeft:`4px solid ${pr.left}`,outline:r.isNew?`1px solid ${C.gold}44`:"none"}}>
                    <td style={{padding:`${pd}px 8px`,textAlign:"center"}}>{perms?.canSendSupport&&<input type="checkbox" checked={selIds.has(r.id)} onChange={()=>toggleSel(r.id)} style={{cursor:"pointer",accentColor:C.gold}}/>}</td>
                    <td style={{padding:`${pd}px 14px`,fontWeight:800,color:C.text1,fontSize:11}}>
                      <div>{r.nuvem}</div>
                      {pr.level!=="normal"&&<div style={{display:"inline-flex",marginTop:4,background:C.white,border:`1px solid ${pr.bd}`,color:pr.color,borderRadius:3,padding:"2px 6px",fontSize:9,fontWeight:800,textTransform:"uppercase"}}>{pr.label}</div>}
                    </td>
                    <td style={{padding:`${pd}px 14px`,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:C.text1}} title={r.destinatario}>{r.destinatario}</td>
                    <td style={{padding:`${pd}px 14px`,color:C.text2,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={r.transportadora}>{r.transportadora}</td>
                    <td style={{padding:`${pd}px 10px`}}><StatusBadge val={r.status}/></td>
                    <td style={{padding:`${pd}px 14px`,fontSize:11,color:C.text2,whiteSpace:"nowrap"}}>{r.prazo||"—"}</td>
                    <td style={{padding:`${pd}px 10px`}}><SituacaoPrazoBadge prazo={r.prazo} status={r.status} entregueNoPrazo={r.entregueNoPrazo}/></td>
                    <td style={{padding:`${pd}px 10px`}}><Chip val={r.urgencia} styles={urgStyles}/></td>
                    <td style={{padding:`${pd}px 14px`}}><SemMovBadge ultimaMov={r.ultimaMov}/></td>
                    <td style={{padding:`${pd}px 8px`}}>{perms?.canSendSupport&&(
                      <div style={{display:"flex",gap:4}}>
                        <button onClick={()=>upd(r.id,{enviadoSuporte:true,atendimento:"Aberto",sentAt:new Date().toISOString()},{acao:"Enviado ao suporte"})} style={{flex:1,background:C.cream,border:`1px solid ${C.border}`,color:C.text2,borderRadius:6,padding:"4px 6px",fontSize:9,cursor:"pointer",fontWeight:500,whiteSpace:"nowrap"}}>Suporte →</button>
                        {perms?.canOperate&&<button onClick={()=>handleArchiveFromLog(r.id)} style={{flex:1,background:C.greenSoft,border:`1px solid ${C.greenBorder}`,color:C.green,borderRadius:6,padding:"4px 6px",fontSize:9,cursor:"pointer",fontWeight:600,whiteSpace:"nowrap"}}>✓ Arq.</button>}
                      </div>
                    )}</td>
                    <td style={{padding:`${pd}px 8px`,textAlign:"center"}}>{perms?.canDelete&&<button onClick={()=>del(r.id)} style={{background:"transparent",border:"none",color:C.text4,cursor:"pointer",fontSize:14}}>×</button>}</td>
                  </tr>
                )})}
              </tbody>
            </table>
          </div>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginTop:12}}>
            <div style={{fontSize:11,color:C.text4}}>{filteredLog.length===0?"Nenhum resultado":`Mostrando ${((safeP-1)*PAGE_SIZE)+1}–${Math.min(safeP*PAGE_SIZE,filteredLog.length)} de ${filteredLog.length} pedidos`}</div>
            {totalPages>1&&<div style={{display:"flex",gap:4,alignItems:"center"}}>
              <button onClick={()=>setLPage(n=>Math.max(1,n-1))} disabled={safeP===1} style={{...getINP(),padding:"5px 12px",cursor:safeP===1?"not-allowed":"pointer",opacity:safeP===1?0.4:1}}>‹</button>
              <span style={{fontSize:11,color:C.text3,padding:"0 10px"}}>{safeP} / {totalPages}</span>
              <button onClick={()=>setLPage(n=>Math.min(totalPages,n+1))} disabled={safeP===totalPages} style={{...getINP(),padding:"5px 12px",cursor:safeP===totalPages?"not-allowed":"pointer",opacity:safeP===totalPages?0.4:1}}>›</button>
            </div>}
          </div>
        </div>
      )}

      {/* ── SUPORTE ── */}
      {tab==="suporte"&&(
        <div style={{display:"flex",height:"calc(100vh - 110px)",overflow:"hidden"}}>

          {/* ── Fila lateral ── */}
          <div style={{width:detail?"360px":"100%",maxWidth:detail?"360px":"100%",borderRight:detail?`1px solid ${C.border}`:"none",display:"flex",flexDirection:"column",background:C.white,flexShrink:0}}>
            <div style={{padding:"16px 18px",borderBottom:`1px solid ${C.border}`,background:C.cream}}>
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8,marginBottom:14}}>
                {[{label:"Na fila",v:ss.total},{label:"Abertos",v:ss.abertos,red:true},{label:"Em andamento",v:ss.andamento}].map(s=>(
                  <div key={s.label} style={{textAlign:"center",padding:"10px 4px",borderRadius:10,background:C.white,border:`1px solid ${C.border}`,boxShadow:shadow.sm}}>
                    <div style={{fontSize:22,fontWeight:600,color:s.red&&s.v>0?C.red:C.gold,fontFamily:"'Cormorant Garamond',serif"}}>{s.v}</div>
                    <div style={{fontSize:8,color:C.text4,marginTop:3,letterSpacing:"0.1em",textTransform:"uppercase"}}>{s.label}</div>
                  </div>
                ))}
              </div>
              <div style={{display:"flex",gap:6,marginBottom:8}}>
              <button onClick={()=>setSupView(v=>v==="lista"?"kanban":"lista")} style={{background:supView==="kanban"?C.brand:C.white,border:`1px solid ${supView==="kanban"?C.brand:C.border}`,color:supView==="kanban"?C.white:C.text2,borderRadius:8,padding:"5px 12px",fontSize:10,cursor:"pointer",fontWeight:500,letterSpacing:"0.06em"}}>{supView==="kanban"?"☰ Lista":"⊞ Kanban"}</button>
            </div>
            <FilterBar
              search={sSrch} onSearch={setSSrch}
              showFilters={sShowFilters} onToggleFilters={()=>setSShowFilters(v=>!v)}
              filters={[
                {key:"at",   label:"Atendimento", value:sAtend, setValue:setSAtend, opts:["Todos","Aberto","Em andamento"]},
                {key:"urg",  label:"Urgência",    value:sUrg,   setValue:setSUrg,   opts:["Todos","Alta","Média","Baixa"]},
                {key:"resp", label:"Responsável", value:sResp,  setValue:setSResp,  opts:["Todos",...respOpts]},
              ]}
              onClearAll={()=>{setSAtend("Todos");setSUrg("Todos");setSResp("Todos");setSSrch("");setSShowFilters(false)}}
              compact
            />
              {perms?.canOperate&&selSupIds.size>0&&(
                <div style={{background:C.brand,borderRadius:8,padding:"9px 14px",display:"flex",alignItems:"center",gap:8}}>
                  <span style={{color:"#888",fontSize:11,flex:1}}>{selSupIds.size} selecionado{selSupIds.size>1?"s":""}</span>
                  <button onClick={bulkArchive} style={{background:C.gold,border:"none",color:C.white,borderRadius:6,padding:"5px 12px",fontSize:10,cursor:"pointer",fontWeight:500}}>Finalizar ({selSupIds.size})</button>
                  <button onClick={()=>setSelSupIds(new Set())} style={{background:"transparent",border:"none",color:"#666",fontSize:12,cursor:"pointer"}}>✕</button>
                </div>
              )}
            </div>
            <div ref={queueRef} style={{overflowY:"auto",flex:1}}>
              {supView==="kanban"?(
                <KanbanSuporteView rows={supRows} onSelect={setSelSup} selSup={selSup} perms={perms} upd={upd} nomeAtendente={nomeAtendente}/>
              ):ss.total===0
                ?<div style={{textAlign:"center",padding:"56px 20px",color:C.text4}}><div style={{fontSize:32,marginBottom:12,opacity:0.2}}>◎</div><div style={{fontSize:13}}>Fila vazia</div><div style={{fontSize:11,marginTop:4}}>Pedidos enviados da Logística aparecem aqui</div></div>
                :supRows.length===0?<div style={{textAlign:"center",padding:24,color:C.text4,fontSize:12}}>Nenhum resultado</div>
                :supRows.map(r=>{
                  const isSel = selSup===r.id
                  const tipo  = classificarProblema(r)
                  const cfg   = PROBLEMA_CONFIG[tipo]
                  const acColor = (tipo==="EXTRAVIO"||tipo==="POSSIVEL_EXTRAVIO")?C.red:r.urgencia==="Alta"?C.red:r.urgencia==="Média"?C.gold:C.green
                  return (
                    <div key={r.id} data-id={r.id} onClick={()=>setSelSup(isSel?null:r.id)}
                      style={{padding:"12px 16px 12px 14px",cursor:"pointer",borderBottom:`1px solid ${C.border}55`,borderLeft:`3px solid ${acColor}`,background:isSel?C.amberSoft:r.alertaStatus?`${C.amber}08`:C.white,transition:"background .15s",display:"flex",alignItems:"flex-start",gap:10}}>
                      {perms?.canOperate&&<input type="checkbox" checked={selSupIds.has(r.id)} onClick={e=>e.stopPropagation()} onChange={()=>toggleSelSup(r.id)} style={{marginTop:3,cursor:"pointer",accentColor:C.gold,flexShrink:0}}/>}
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:3}}>
                          <span style={{fontWeight:600,fontSize:12,color:C.text1}}>{r.nuvem}</span>
                          <TimeOpenBadge sentAt={r.sentAt}/>
                        </div>
                        {/* Badge de tipo de problema na fila */}
                        {tipo!=="OK"&&(
                          <div style={{background:cfg.bg,border:`1px solid ${cfg.bd}`,borderRadius:6,padding:"2px 8px",fontSize:9,color:cfg.color,marginBottom:4,display:"inline-flex",alignItems:"center",gap:3,fontWeight:700,letterSpacing:"0.04em"}}>
                            {cfg.icone} {cfg.label}
                          </div>
                        )}
                        {r.alertaStatus&&<div style={{background:C.amberSoft,border:`1px solid ${C.amberBorder}`,borderRadius:5,padding:"2px 8px",fontSize:10,color:C.amber,marginBottom:4,display:"flex",alignItems:"center",gap:4}}><span>⚠</span><span style={{fontWeight:500}}>Status alterado!</span></div>}
                        <div style={{fontSize:11,color:C.text2,marginBottom:5,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.destinatario}</div>
                        <div style={{display:"flex",gap:4,flexWrap:"wrap",marginBottom:4}}><StatusBadge val={r.status}/><Chip val={r.urgencia} styles={urgStyles}/><Chip val={r.atendimento} styles={atendStyles}/></div>
                        {r.prazo&&<div style={{marginBottom:2}}><SlaCell prazo={r.prazo}/></div>}
                        {r.responsavel&&<div style={{fontSize:9,color:C.text4,marginTop:2,letterSpacing:"0.06em"}}>RESP: {r.responsavel}</div>}
                      </div>
                    </div>
                  )
                })
              }
            </div>
          </div>
          {detail?(
            <div ref={detailPanelRef} style={{flex:1,display:"flex",flexDirection:"column",overflowY:"auto",background:C.cream,height:"100%"}}>

              {/* BLOCO 1 — TOPO STICKY: título + HeaderProblema + AcoesRapidas */}
              <div style={{background:C.white,padding:"12px 18px",borderBottom:`1px solid ${C.border}`,position:"sticky",top:0,zIndex:5,boxShadow:shadow.sm}}>
                <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:10}}>
                  <div>
                    <div style={{fontSize:9,color:C.text3,textTransform:"uppercase",marginBottom:4,fontWeight:800}}>Pedido em atendimento</div>
                    <div style={{fontSize:18,fontWeight:800,color:C.text1,marginBottom:2}}>{detail.destinatario}</div>
                    <div style={{fontSize:11,color:C.text3}}>#{detail.nuvem} · {detail.transportadora}</div>
                  </div>
                  <button onClick={()=>setSelSup(null)} style={{background:"transparent",border:`1px solid ${C.border}`,color:C.text4,cursor:"pointer",fontSize:16,width:32,height:32,display:"flex",alignItems:"center",justifyContent:"center",borderRadius:8,flexShrink:0}}>×</button>
                </div>

                {/* HeaderProblema: classificação visual + métricas */}
                <HeaderProblema
                  r={detail}
                  onNotificou={()=>upd(detail.id,{alertaStatus:null},{acao:"Alerta dispensado — cliente notificado",usuario:nomeAtendente})}
                />

                {/* AcoesRapidas: botões de ação com loading */}
                <AcoesRapidas
                  r={detail}
                  perms={perms}
                  nomeAtendente={nomeAtendente}
                  onNotificar={()=>{
                    upd(detail.id,{atendimento:detail.atendimento==="Aberto"?"Em andamento":detail.atendimento},{acao:"Cliente notificado",usuario:nomeAtendente})
                    addToast("Notificacao registrada")
                  }}
                  onAcionarTransp={()=>{
                    upd(detail.id,{atendimento:detail.atendimento==="Aberto"?"Em andamento":detail.atendimento},{acao:`Transportadora acionada: ${detail.transportadora}`,usuario:nomeAtendente})
                    addToast(`🚛 ${detail.transportadora} acionada`)
                  }}
                  onReenvio={()=>{
                    handleCreateReenvio(detail.id)
                  }}
                  onResolver={()=>handleResolve(detail.id)}
                  onDevolver={()=>handleReturnLog(detail.id)}
                  onTotalTracking={()=>handleTotalExpressTracking(detail.id)}
                  onTotalTicket={()=>handleTotalExpressTicket(detail.id)}
                />
              </div>

              {/* BLOCO 2 + 3 — CONTEÚDO ROLÁVEL */}
              <div style={{padding:"18px 22px",flex:1}}>

                {/* SugestaoSistema: sugestão automática */}
                <SugestaoSistema r={detail}/>

                {perms?.canOperate&&(
                  <div style={{background:C.white,borderRadius:10,border:`1px solid ${C.border}`,padding:14,marginBottom:14,boxShadow:shadow.sm}}>
                    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:12,marginBottom:10}}>
                      <div>
                        <div style={{fontSize:8,color:C.text3,textTransform:"uppercase",letterSpacing:"0.14em",fontWeight:800,marginBottom:3}}>Decisao do cliente</div>
                        <div style={{fontSize:11,color:C.text4}}>Registre o motivo e direcione o pedido para o fluxo correto.</div>
                      </div>
                      {detail.decisaoCliente&&<Chip val={detail.decisaoCliente} styles={{[detail.decisaoCliente]:{bg:C.cream,color:C.text1,bd:C.borderDark}}}/>}
                    </div>
                    <textarea value={detail.motivoDevolucao||detail.motivo||""} onChange={e=>upd(detail.id,{motivoDevolucao:e.target.value})} placeholder="Motivo da devolucao, estorno ou reenvio..." rows={3}
                      style={{width:"100%",borderRadius:6,border:`1px solid ${C.borderDark}`,padding:"10px 12px",fontSize:12,resize:"vertical",fontFamily:"inherit",background:C.white,color:C.text1,boxSizing:"border-box",lineHeight:1.5,marginBottom:10}}/>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr auto",gap:8}}>
                      <button onClick={()=>handleMarkDevolucao(detail.id)} style={{background:detail.decisaoCliente==="Estorno / devolucao"?C.brand:C.white,border:`1px solid ${C.borderDark}`,color:detail.decisaoCliente==="Estorno / devolucao"?C.white:C.text1,borderRadius:6,padding:"10px 12px",fontSize:11,cursor:"pointer",fontWeight:800}}>Estorno / devolver produto</button>
                      <button onClick={()=>handleCreateReenvio(detail.id)} style={{background:detail.decisaoCliente==="Reenvio"?C.brand:C.white,border:`1px solid ${C.borderDark}`,color:detail.decisaoCliente==="Reenvio"?C.white:C.text1,borderRadius:6,padding:"10px 12px",fontSize:11,cursor:"pointer",fontWeight:800}}>Reenvio</button>
                      <button onClick={()=>handleResolve(detail.id)} style={{background:C.green,border:`1px solid ${C.green}`,color:C.white,borderRadius:6,padding:"10px 14px",fontSize:11,cursor:"pointer",fontWeight:800}}>Finalizar</button>
                    </div>
                  </div>
                )}

                {/* Grid 3 colunas: LOGÍSTICA | CLIENTE | PRAZO */}
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:14}}>

                  {/* LOGÍSTICA */}
                  <div style={{background:C.white,borderRadius:12,border:`1px solid ${C.border}`,padding:"14px 16px",boxShadow:shadow.sm}}>
                    <div style={{fontSize:8,color:C.gold,textTransform:"uppercase",letterSpacing:"0.16em",fontWeight:700,marginBottom:10,paddingBottom:6,borderBottom:`1px solid ${C.border}`}}>Logística</div>
                    {[
                      ["Transportadora", detail.transportadora||"—"],
                      ["Cód. Rastreio",  detail.rastreio?<span style={{fontFamily:"monospace",fontSize:10}}>{detail.rastreio}</span>:"—"],
                      ["Status",         <StatusBadge val={detail.status}/>],
                      ["Última atualiz.",<SemMovBadge ultimaMov={detail.ultimaMov}/>],
                    ].map(([lbl,val])=>(
                      <div key={lbl} style={{marginBottom:8}}>
                        <div style={{fontSize:8,color:C.text4,textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:2}}>{lbl}</div>
                        <div style={{fontSize:11,color:C.text1}}>{val}</div>
                      </div>
                    ))}
                  </div>

                  {/* CLIENTE */}
                  <div style={{background:C.white,borderRadius:12,border:`1px solid ${C.border}`,padding:"14px 16px",boxShadow:shadow.sm}}>
                    <div style={{fontSize:8,color:C.gold,textTransform:"uppercase",letterSpacing:"0.16em",fontWeight:700,marginBottom:10,paddingBottom:6,borderBottom:`1px solid ${C.border}`}}>Cliente</div>
                    {[
                      ["Nome",     detail.destinatario||"—"],
                      ["Cidade/UF",[detail.cidade,detail.uf].filter(Boolean).join(" / ")||"—"],
                      ["CEP",      detail.cep||"—"],
                      ["Motivo",   detail.motivo||"—"],
                    ].map(([lbl,val])=>(
                      <div key={lbl} style={{marginBottom:8}}>
                        <div style={{fontSize:8,color:C.text4,textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:2}}>{lbl}</div>
                        <div style={{fontSize:11,color:C.text1}}>{val}</div>
                      </div>
                    ))}
                  </div>

                  {/* PRAZO */}
                  <div style={{background:C.white,borderRadius:12,border:`1px solid ${C.border}`,padding:"14px 16px",boxShadow:shadow.sm}}>
                    <div style={{fontSize:8,color:C.gold,textTransform:"uppercase",letterSpacing:"0.16em",fontWeight:700,marginBottom:10,paddingBottom:6,borderBottom:`1px solid ${C.border}`}}>Prazo</div>
                    {(()=>{
                      const dt2=parsePrazo(detail.prazo)
                      const h2=new Date(); h2.setHours(0,0,0,0)
                      const diasAtraso=dt2?Math.ceil((h2-dt2)/86400000):0
                      return <>
                        <div style={{marginBottom:8}}>
                          <div style={{fontSize:8,color:C.text4,textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:2}}>Prazo logístico</div>
                          <SlaCell prazo={detail.prazo}/>
                        </div>
                        <div style={{marginBottom:8}}>
                          <div style={{fontSize:8,color:C.text4,textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:2}}>Status prazo</div>
                          <div style={{fontSize:11,color:C.text1}}>{detail.statusPrazoRaw||"—"}</div>
                        </div>
                        {diasAtraso>0&&(
                          <div style={{marginBottom:8}}>
                            <div style={{fontSize:8,color:C.text4,textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:2}}>Dias de atraso</div>
                            <span style={{background:C.redSoft,color:C.red,border:`1px solid ${C.redBorder}`,borderRadius:10,padding:"2px 8px",fontSize:11,fontWeight:700}}>{diasAtraso}d atrasado</span>
                          </div>
                        )}
                        <div>
                          <div style={{fontSize:8,color:C.text4,textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:2}}>Data de envio</div>
                          <div style={{fontSize:11,color:C.text2}}>{detail.dataCriacao||"—"}</div>
                        </div>
                      </>
                    })()}
                  </div>
                </div>

                {/* Operação: Responsável + Chamado + Obs */}
                {perms?.canOperate&&(
                  <div style={{background:C.white,borderRadius:12,border:`1px solid ${C.border}`,padding:"14px 16px",marginBottom:14,boxShadow:shadow.sm}}>
                    <div style={{fontSize:8,color:C.gold,textTransform:"uppercase",letterSpacing:"0.16em",fontWeight:700,marginBottom:12,paddingBottom:6,borderBottom:`1px solid ${C.border}`}}>Operação</div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
                      {[["Responsável","responsavel","Nome do responsável..."],["Nº Chamado Zendesk","chamado","Ex: #45821"]].map(([lbl,key,ph])=>(
                        <div key={key}>
                          <div style={{fontSize:8,color:C.text4,textTransform:"uppercase",letterSpacing:"0.12em",marginBottom:5,fontWeight:500}}>{lbl}</div>
                          <input value={detail[key]||""} onChange={e=>upd(detail.id,{[key]:e.target.value})} placeholder={ph} style={{...getINP(),width:"100%",boxSizing:"border-box"}}/>
                        </div>
                      ))}
                    </div>
                    <div>
                      <div style={{fontSize:8,color:C.text4,textTransform:"uppercase",letterSpacing:"0.12em",marginBottom:5,fontWeight:500}}>Observações</div>
                      <textarea value={detail.obs||""} onChange={e=>upd(detail.id,{obs:e.target.value})} placeholder="Anotações do atendimento..." rows={2}
                        style={{width:"100%",borderRadius:8,border:`1px solid ${C.border}`,padding:"10px 12px",fontSize:12,resize:"vertical",fontFamily:"inherit",background:C.white,color:C.text1,boxSizing:"border-box",lineHeight:1.6}}/>
                    </div>
                  </div>
                )}

                {/* Timeline de Histórico */}
                <TimelineHistorico historico={detail.historico} isOpen={openHist} onToggle={()=>setOpenHist(v=>!v)}/>
              </div>
            </div>
          ):ss.total>0?(
            <div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center",background:C.cream}}>
              <div style={{textAlign:"center"}}>
                <div style={{width:48,height:1,background:C.border,margin:"0 auto 20px"}}/>
                <div style={{fontSize:10,letterSpacing:"0.18em",textTransform:"uppercase",color:C.text4}}>Selecione um pedido para atender</div>
                <div style={{fontSize:9,color:C.text4,marginTop:6,letterSpacing:"0.06em"}}>ou selecione vários para finalizar em lote</div>
              </div>
            </div>
          ):null}
        </div>
      )}

      {/* ── FINALIZADOS ── */}
      {tab==="devolucao"&&!showImp&&(
        <OperacaoEspecialPanel
          type="devolucao"
          rows={baseDev}
          perms={perms}
          upd={upd}
          onCreateReenvio={handleCreateReenvio}
          onResolve={handleResolve}
        />
      )}

      {tab==="reenvio"&&!showImp&&(
        <OperacaoEspecialPanel
          type="reenvio"
          rows={baseReenvio}
          perms={perms}
          upd={upd}
          onCreateReenvio={handleCreateReenvio}
          onResolve={handleResolve}
        />
      )}

      {tab==="arquivados"&&(
        <div style={{padding:"24px 32px"}}>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12,marginBottom:20}}>
            <KpiCard label="Total finalizados" val={arch}/>
            <KpiCard label="Resolvidos hoje" val={rows.filter(r=>{if(r.atendimento!=="Resolvido")return false;const h=r.historico.find(x=>x.acao&&(x.acao.includes("Resolvido")||x.acao.includes("Arquivado")));return h&&h.ts&&h.ts.startsWith(new Date().toLocaleDateString("pt-BR"))}).length}/>
            <KpiCard label="Com observações" val={baseArch.filter(r=>r.obs&&r.obs.trim()).length}/>
          </div>
          {arch===0?<div style={{textAlign:"center",padding:"56px 0",color:C.text4}}><div style={{fontSize:32,marginBottom:12,opacity:0.2}}>◎</div><div style={{fontSize:14}}>Nenhum atendimento finalizado</div></div>:(
            <div>
              <div style={{marginBottom:14}}><input value={aSrch} onChange={e=>setASrch(e.target.value)} placeholder="Buscar nos finalizados..." style={{...getINP(),width:"100%",padding:"10px 14px",boxSizing:"border-box",boxShadow:shadow.sm}}/></div>
              {(()=>{
                const aTotalPages = Math.max(1,Math.ceil(archRows.length/PAGE_SIZE))
                const aSafeP = Math.min(aPage,aTotalPages)
                const pagedArch = archRows.slice((aSafeP-1)*PAGE_SIZE, aSafeP*PAGE_SIZE)
                return <>
                  <div style={{overflowX:"auto",overflowY:"auto",maxHeight:"60vh",borderRadius:12,border:`1px solid ${C.border}`,boxShadow:shadow.sm}}>
                    <table style={{width:"100%",borderCollapse:"collapse",fontSize:compact?11:12,tableLayout:"fixed",minWidth:900}}>
                      <colgroup><col style={{width:90}}/><col style={{width:150}}/><col style={{width:110}}/><col style={{width:110}}/><col style={{width:120}}/><col style={{width:70}}/><col style={{width:110}}/><col style={{width:96}}/><col style={{width:96}}/><col style={{width:90}}/></colgroup>
                      <thead><tr>{["No NUVEM","Destinatário","Transportadora","Status","Motivo","Urgência","Prazo / SLA","Chamado","Responsável","Ações"].map(h=><th key={h} style={THF}>{h}</th>)}</tr></thead>
                      <tbody>
                        {pagedArch.length===0?<tr><td colSpan={10} style={{textAlign:"center",padding:32,color:C.text4}}>Nenhum resultado</td></tr>
                        :pagedArch.map((r,i)=>(
                          <tr key={r.id} style={{background:i%2===0?C.white:C.cream,borderBottom:`1px solid ${C.border}55`}}>
                            <td style={{padding:`${pd}px 14px`,fontWeight:600,color:C.text3,fontSize:11}}>{r.nuvem}</td>
                            <td style={{padding:`${pd}px 14px`,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:C.text1}} title={r.destinatario}>{r.destinatario}</td>
                            <td style={{padding:`${pd}px 14px`,color:C.text2,overflow:"hidden",textOverflow:"ellipsis"}}>{r.transportadora}</td>
                            <td style={{padding:`${pd}px 14px`}}><StatusBadge val={r.status}/></td>
                            <td style={{padding:`${pd}px 14px`,color:C.text3,fontSize:10,overflow:"hidden",textOverflow:"ellipsis"}} title={r.motivo}>{r.motivo}</td>
                            <td style={{padding:`${pd}px 14px`}}><Chip val={r.urgencia} styles={urgStyles}/></td>
                            <td style={{padding:`${pd}px 14px`}}><SlaCell prazo={r.prazo}/></td>
                            <td style={{padding:`${pd}px 14px`,color:C.text3,fontSize:11}}>{r.chamado||"—"}</td>
                            <td style={{padding:`${pd}px 14px`,color:C.text3,fontSize:11,overflow:"hidden",textOverflow:"ellipsis"}}>{r.responsavel||"—"}</td>
                            <td style={{padding:`${pd}px 14px`}}>{perms?.canOperate&&<button onClick={()=>upd(r.id,{atendimento:"Em andamento"},{acao:"Reaberto dos finalizados",usuario:nomeAtendente})} style={{background:C.cream,border:`1px solid ${C.border}`,color:C.text2,borderRadius:6,padding:"4px 12px",fontSize:10,cursor:"pointer",fontWeight:500}}>Reabrir</button>}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginTop:10}}>
                    <div style={{fontSize:11,color:C.text4}}>{archRows.length===0?"Nenhum resultado":`Mostrando ${((aSafeP-1)*PAGE_SIZE)+1}–${Math.min(aSafeP*PAGE_SIZE,archRows.length)} de ${archRows.length} finalizados`}</div>
                    {aTotalPages>1&&<div style={{display:"flex",gap:4,alignItems:"center"}}>
                      <button onClick={()=>setAPage(n=>Math.max(1,n-1))} disabled={aSafeP===1} style={{...getINP(),padding:"5px 12px",cursor:aSafeP===1?"not-allowed":"pointer",opacity:aSafeP===1?0.4:1}}>‹</button>
                      <span style={{fontSize:11,color:C.text3,padding:"0 10px"}}>{aSafeP} / {aTotalPages}</span>
                      <button onClick={()=>setAPage(n=>Math.min(aTotalPages,n+1))} disabled={aSafeP===aTotalPages} style={{...getINP(),padding:"5px 12px",cursor:aSafeP===aTotalPages?"not-allowed":"pointer",opacity:aSafeP===aTotalPages?0.4:1}}>›</button>
                    </div>}
                  </div>
                </>
              })()}
            </div>
          )}
        </div>
      )}
    </div>
  )
}
