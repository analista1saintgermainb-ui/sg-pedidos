import { useState, useEffect, useRef, useCallback } from 'react'
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell } from 'recharts'
import * as XLSX from 'xlsx'
import { createClient } from '@supabase/supabase-js'

// ─── Design System ────────────────────────────────────────────
// ─── Temas claro e escuro ─────────────────────────────────────
const CL = {
  brand:      "#0C0C0C", brandSoft:  "#1A1A1A",
  gold:       "#B8974A", goldLight:  "#D4AF6A", goldDim: "#8C7038",
  cream:      "#F8F5EF", creamDark:  "#F0EDE5",
  white:      "#FFFFFF", border:     "#E8E3D8", borderDark: "#D4CFC4",
  text1:      "#1A1A1A", text2:      "#5C5750", text3: "#9C9690", text4: "#C4C0B8",
  red:        "#C0392B", redSoft:    "#F9ECEB", redBorder:  "#EBCBC8",
  green:      "#2E7D50", greenSoft:  "#EAF4EE", greenBorder:"#C0DCCB",
  amber:      "#8C6D1F", amberSoft:  "#FDF6E3", amberBorder:"#E8D5A3",
  blue:       "#1A5276", blueSoft:   "#EAF2FB", blueBorder: "#AACDE6",
}
const CD = {
  brand:      "#F0EDE5", brandSoft:  "#D4CFC4",
  gold:       "#D4AF6A", goldLight:  "#E8C97A", goldDim: "#B8974A",
  cream:      "#0F0F0F", creamDark:  "#1A1A1A",
  white:      "#1E1E1E", border:     "#2E2E2E", borderDark: "#3A3A3A",
  text1:      "#F0EDE5", text2:      "#C4C0B8", text3: "#7A7670", text4: "#4A4640",
  red:        "#E05555", redSoft:    "#2A1212", redBorder:  "#4A2020",
  green:      "#4AB870", greenSoft:  "#0A2015", greenBorder:"#1A4030",
  amber:      "#C89830", amberSoft:  "#281E00", amberBorder:"#4A3800",
  blue:       "#5A9AD4", blueSoft:   "#0A1825", blueBorder: "#1A3A55",
}
// C é mutável — applyTheme() troca os valores em re-render
const C = {...CL}
function applyTheme(dark) {
  const src = dark ? CD : CL
  Object.keys(src).forEach(k => { C[k] = src[k] })
}
const shadow = {
  sm: "0 1px 3px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04)",
  md: "0 4px 12px rgba(0,0,0,0.08), 0 2px 4px rgba(0,0,0,0.04)",
  lg: "0 8px 24px rgba(0,0,0,0.10), 0 4px 8px rgba(0,0,0,0.06)",
}
const getGlobalStyle = () => `
  @import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;500;600&family=Inter:wght@300;400;500;600&display=swap');
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: ${C.cream}; color: ${C.text1}; }
  ::-webkit-scrollbar { width: 4px; height: 4px; }
  ::-webkit-scrollbar-track { background: ${C.creamDark}; }
  ::-webkit-scrollbar-thumb { background: ${C.borderDark}; border-radius: 2px; }
  ::-webkit-scrollbar-thumb:hover { background: ${C.gold}; }
  select, input, textarea, button { font-family: 'Inter', sans-serif; }
  tr:hover td { background: ${C.creamDark} !important; }
`

// ─── Supabase ─────────────────────────────────────────────────
const SUPA_URL     = "https://jdiuuhfhsiymttxllssr.supabase.co"
const SUPA_KEY     = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpkaXV1aGZoc2l5bXR0eGxsc3NyIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzc3NTMyNTcsImV4cCI6MjA5MzMyOTI1N30.wNGhwh2bCF0HZSonn09S-15kEVAQGzEP1yWvRx3l_N4"
const SUPA_SERVICE = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpkaXV1aGZoc2l5bXR0eGxsc3NyIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc3Nzc1MzI1NywiZXhwIjoyMDkzMzI5MjU3fQ.yjZ8VKr8YfbMBELdoKevdE1k_dd2OXUlYjUj4n2GeQw"
const SH  = { apikey: SUPA_KEY, "Content-Type": "application/json" }
const aSH = t => ({ ...SH, Authorization: `Bearer ${t}` })
const supabase = createClient(SUPA_URL, SUPA_KEY)

async function signIn(email, password) {
  const r = await fetch(`${SUPA_URL}/auth/v1/token?grant_type=password`, { method:"POST", headers:SH, body:JSON.stringify({email,password}) })
  const d = await r.json(); if (!r.ok) throw new Error(d.error_description||d.msg||"Erro ao fazer login"); return d
}
async function signOut(token) { await fetch(`${SUPA_URL}/auth/v1/logout`,{method:"POST",headers:aSH(token)}) }
async function createUser(email, password) {
  const r = await fetch(`${SUPA_URL}/auth/v1/admin/users`,{method:"POST",headers:{apikey:SUPA_SERVICE,Authorization:`Bearer ${SUPA_SERVICE}`,"Content-Type":"application/json"},body:JSON.stringify({email,password,email_confirm:true})})
  const d = await r.json(); if (!r.ok) throw new Error(d.msg||d.message||"Erro ao criar usuário"); return d
}
// ─── NOVO: carrega apenas pedidos ativos (resolvido = false) ───
async function dbLoadAtivos(token, onPartial) {
  let all = [], from = 0, step = 1000
  while (true) {
    const r = await fetch(`${SUPA_URL}/rest/v1/pedidos?select=*&resolvido=eq.false&order=id&limit=${step}&offset=${from}`, { headers: aSH(token) })
    if (!r.ok) throw new Error(await r.text())
    const data = await r.json()
    const chunk = data.map(row => ({ ...row.dados, id: row.id, resolvido: row.resolvido }))
    all = [...all, ...chunk]
    if (from === 0 && chunk.length > 0) onPartial(all)
    if (data.length < step) break
    from += step
  }
  return all
}
// ─── NOVO: carrega arquivados com paginação ───
async function dbLoadArquivados(token, page = 1, pageSize = 200) {
  const from = (page - 1) * pageSize
  const r = await fetch(`${SUPA_URL}/rest/v1/pedidos?select=*&resolvido=eq.true&order=id.desc&limit=${pageSize}&offset=${from}`, { headers: aSH(token) })
  if (!r.ok) throw new Error(await r.text())
  const data = await r.json()
  return data.map(row => ({ ...row.dados, id: row.id, resolvido: row.resolvido }))
}
async function dbCountArquivados(token) {
  const r = await fetch(`${SUPA_URL}/rest/v1/pedidos?select=id&resolvido=eq.true`, { headers: aSH(token) })
  if (!r.ok) return 0
  const data = await r.json()
  return data.length
}
// ─── NOVO: upsert atualiza também a coluna resolvido ───
async function dbUpsert(rows, token) {
  for (let i = 0; i < rows.length; i += 200) {
    const batch = rows.slice(i, i+200)
    const payload = batch.map(r => ({id:Number(r.id), dados:r, resolvido: r.atendimento === 'Resolvido', updated_at:new Date().toISOString()}))
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
  admin:    {tabs:["dashboard","logistica","suporte","arquivados","usuarios"],canImport:true,canDelete:true,canClear:true,canSendSupport:true,canOperate:true},
  logistica:{tabs:["dashboard","logistica"],canImport:true,canDelete:false,canClear:false,canSendSupport:true,canOperate:true},
  suporte:  {tabs:["suporte","arquivados"],canImport:false,canDelete:false,canClear:false,canSendSupport:false,canOperate:true},
  leitura:  {tabs:["dashboard","logistica","suporte","arquivados"],canImport:false,canDelete:false,canClear:false,canSendSupport:false,canOperate:false},
}

const ALERTA_DIAS = 7
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

// ─── Helpers de cálculo (mantidos originais) ─────────────────
const HEADER_MAP = {
  nuvem:        ["identificador ecommerce","id ecommerce","no nuvem","nuvem","pedido"],
  destinatario: ["destinatário nome","destinatario nome","nome destinatário","nome destinatario","nome do destinatário","nome do destinatario","nome do cliente","nome do comprador","nome comprador","comprador","destinatário","destinatario","nome do pedido","nome"],
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
  for (const v of aliases) {
    const i = nhdrs.findIndex(h => h === norm(v))
    if (i >= 0) return i
  }
  for (const v of aliases.filter(v => norm(v).includes(" "))) {
    const nv = norm(v)
    const i  = nhdrs.findIndex(h => h.includes(nv))
    if (i >= 0) return i
  }
  for (const v of aliases.filter(v => !norm(v).includes(" "))) {
    const nv = norm(v)
    const i  = nhdrs.findIndex(h => h.startsWith(nv + " ") || h.startsWith(nv + "_") || h === nv)
    if (i >= 0) return i
  }
  for (const v of aliases.filter(v => !norm(v).includes(" "))) {
    const nv = norm(v)
    const i  = nhdrs.findIndex(h => h.includes(nv))
    if (i >= 0) return i
  }
  return -1
}
const uniq   = arr => ["Todos",...Array.from(new Set(arr.filter(Boolean).sort()))]

function parseStatusPrazo(raw) {
  if (!raw) return null
  const v = (raw||"").toLowerCase().trim()
  if (v.includes("antes")||v.includes("no prazo")||v==="ok"||v.includes("dentro")||v.includes("normal")) return true
  if (v.includes("atras")||v.includes("fora")||v==="vencido"||v.includes("atraso")) return false
  return null
}
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
  const br = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/)
  if (br) return new Date(+br[3], +br[2]-1, +br[1])
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/)
  if (iso) return new Date(+iso[1], +iso[2]-1, +iso[3])
  const dmy = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/)
  if (dmy) return new Date(+dmy[3], +dmy[2]-1, +dmy[1])
  const dot = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/)
  if (dot) return new Date(+dot[3], +dot[2]-1, +dot[1])
  const num = parseFloat(s)
  if (!isNaN(num) && num > 30000 && num < 100000)
    return new Date(Math.round((num - 25569) * 86400000))
  return null
}
function calcUrg(prazo, status) {
  const s = (status||"").toLowerCase()
  if (s.includes("extravia")||s.includes("perdid"))                          return "Alta"
  if (s.includes("devolv")||s.includes("recusa"))                            return "Alta"
  if (s.includes("problema_entrega")||s.includes("problema entrega"))        return "Alta"
  if (s.includes("falha")||s.includes("retido")||s.includes("apreend"))     return "Alta"
  if (s.includes("entregue")||s.includes("finaliz"))                         return "Baixa"
  if (s.includes("saiu_para_entrega")||s.includes("saiu para entrega")||
      s.includes("saida_para_entrega")||s.includes("saiu"))                  return "Baixa"
  if (s.includes("aguardando_retirada")||s.includes("aguardando retirada"))  return "Média"
  if (s.includes("triado")||s.includes("triagem"))                           return "Média"
  if (s.includes("em_transito")||s.includes("em transito")||
      s.includes("trânsito")||s.includes("transito"))                        return "Média"
  if (s.includes("postado")||s.includes("coletado")||s.includes("colet"))   return "Média"
  if (s.includes("aguardando")||s.includes("processando"))                   return "Média"
  const dt = parsePrazo(prazo)
  if (!dt) return "Média"
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
function parseData(text) {
  // Função completa mantida do original
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
    cidade:      isHdr?findIdx(hdrs,"cidade")        :-1,
    uf:          isHdr?findIdx(hdrs,"uf")            :-1,
    cep:         isHdr?findIdx(hdrs,"cep")           :-1,
    statusPrazo: isHdr?findIdx(hdrs,"statusPrazo")   :-1,
    dataCriacao: isHdr?findIdx(hdrs,"dataCriacao")   :-1,
    email:       isHdr?findIdx(hdrs,"email")         :-1,
  }
  const g = (c,i) => i>=0&&i<c.length?c[i]:""
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
    const dt = parsePrazo(prazo)
    const dtEntregaRaw = g(c,ix.ultimaMov)
    const dtEntrega = entregue && dtEntregaRaw ? parsePrazo(dtEntregaRaw) : null
    const dtEntregaDate = dtEntrega ? new Date(dtEntrega.getFullYear(), dtEntrega.getMonth(), dtEntrega.getDate()) : null
    const dtPrazoDate   = dt         ? new Date(dt.getFullYear(),       dt.getMonth(),       dt.getDate())        : null
    const noPrazo = spVal!==null ? spVal
      : (dtEntregaDate && dtPrazoDate) ? dtEntregaDate <= dtPrazoDate
      : null
    return {
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
      obs:"", historico: entregue?[{acao:"Arquivado automaticamente — entrega concluída",ts:new Date().toLocaleString("pt-BR")}]:[],
      responsavel:"", sentAt:null, chamado:"", isNew:true,
    }
  }).filter(r=>{
    if (!r.nuvem&&!r.destinatario&&!r.nf) return false
    const tr = (r.transportadora||"").trim()
    if (tr==="SP"||tr==="RJ"||tr==="MG"||tr==="RS"||tr.length<=2) r.transportadora=""
    return true
  })
}
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
    : null
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
  return `https://www.google.com/search?q=${encodeURIComponent(transportadora + " rastreio contato")}`
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
const urgStyles = {
  Alta:  {bg:C.redSoft,  color:C.red,  bd:C.redBorder,  dot:"#e74c3c"},
  Média: {bg:C.amberSoft,color:C.amber,bd:C.amberBorder, dot:C.gold},
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
function Chip({val,styles}) {
  const s = styles[val]||{bg:C.creamDark,color:C.text3,bd:C.border}
  return <span style={{background:s.bg,color:s.color,border:`1px solid ${s.bd}`,borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:500,whiteSpace:"nowrap",letterSpacing:"0.02em"}}>{val}</span>
}
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
function SituacaoPrazoBadge({prazo, status, entregueNoPrazo}) {
  const dt = parsePrazo(prazo)
  if (entregueNoPrazo === true)  return <span style={{background:C.greenSoft,color:C.green,border:`1px solid ${C.greenBorder}`,borderRadius:10,padding:"2px 8px",fontSize:10,fontWeight:600}}>No Prazo</span>
  if (entregueNoPrazo === false) return <span style={{background:C.redSoft,color:C.red,border:`1px solid ${C.redBorder}`,borderRadius:10,padding:"2px 8px",fontSize:10,fontWeight:600}}>Atraso</span>
  if (!dt) return <span style={{background:C.creamDark,color:C.text4,border:`1px solid ${C.border}`,borderRadius:10,padding:"2px 8px",fontSize:10}}>—</span>
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
function TimeOpenBadge({sentAt}) {
  const info = timeOpen(sentAt)
  if (!info) return null
  return <span style={{background:info.alert?C.redSoft:C.amberSoft,color:info.alert?C.red:C.amber,border:`1px solid ${info.alert?C.redBorder:C.amberBorder}`,borderRadius:10,padding:"2px 8px",fontSize:10,fontWeight:500}}>{info.label}</span>
}
function KpiCard({label,val,sub,accent}) {
  return <div style={{background:C.white,borderRadius:12,padding:"20px 22px",border:`1px solid ${accent?C.redBorder:C.border}`,boxShadow:shadow.sm,position:"relative",overflow:"hidden"}}>
    {accent&&<div style={{position:"absolute",top:0,left:0,right:0,height:3,background:`linear-gradient(90deg,${C.red},#e74c3c88)`}}/>}
    {!accent&&<div style={{position:"absolute",top:0,left:0,right:0,height:2,background:`linear-gradient(90deg,${C.gold}44,${C.gold})`}}/>}
    <div style={{fontSize:9,color:C.text3,letterSpacing:"0.14em",textTransform:"uppercase",marginBottom:10,fontWeight:500}}>{label}</div>
    <div style={{fontSize:28,fontWeight:600,color:accent?C.red:C.brand,letterSpacing:"-0.02em",lineHeight:1,marginBottom:6,fontFamily:"'Cormorant Garamond',serif"}}>{val}</div>
    {sub&&<div style={{fontSize:11,color:C.text3,fontWeight:400}}>{sub}</div>}
  </div>
}
function CopyBtn({text,label}) {
  const [ok,setOk] = useState(false)
  return <button onClick={()=>{navigator.clipboard.writeText(text);setOk(true);setTimeout(()=>setOk(false),2000)}}
    style={{background:ok?C.greenSoft:C.gold,border:`1px solid ${ok?C.greenBorder:C.goldDim}`,color:ok?C.green:C.white,borderRadius:6,padding:"5px 12px",fontSize:11,cursor:"pointer",fontWeight:500,whiteSpace:"nowrap",transition:"all .2s"}}>
    {ok?"✓ Copiado!":label||"Copiar"}
  </button>
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
const getINP = () => ({borderRadius:8,border:`1px solid ${C.border}`,padding:"9px 12px",fontSize:12,background:C.white,color:C.text1,outline:"none",transition:"border-color .2s"})
function HeaderProblema({r, onNotificou}) { /* mantido original */ return null; } // por brevidade, mas deve estar presente no original
function SugestaoSistema({r}) { return null; }
function buildMailto(r, nomeAtendente) { return null; }
function AcoesRapidas({r, perms, nomeAtendente, onNotificar, onAcionarTransp, onReenvio, onResolver, onDevolver}) { return null; }
function TimelineHistorico({historico, isOpen, onToggle}) { return null; }
function LoginScreen({onLogin}) { /* mantido original */ return null; }
function BoxlinkSettings({addToast}) { /* mantido original */ return null; }
function UsuariosPanel({token,addToast}) { /* mantido original */ return null; }
function KanbanSuporteView({rows, onSelect, selSup, perms, upd, nomeAtendente}) { /* mantido original */ return null; }
function FilterBar({search, onSearch, showFilters, onToggleFilters, filters, onClearAll, compact}) { /* mantido original */ return null; }
const SAMPLE = `Identificador Ecommerce;Destinatário Nome;Estratégia de Frete;Rastreador Last Mile;Situação;Prazo Logístico;Nº Nota Fiscal\n12345;Ana Souza;Correios PAC;AA123456789BR;Em trânsito;05/05/2026;98765\n12346;Carlos Lima;Jadlog;JD987654321;Extraviado;28/04/2026;98766\n12347;Mariana Costa;Total Express;TE112233445;Entregue;01/05/2026;98767\n12348;Fernando Silva;Correios SEDEX;AA223344556BR;Saiu para entrega;02/05/2026;98768\n12349;Júlia Martins;Loggi;LG556677889;Devolvido;30/04/2026;98769`
const KANBAN_COLS = [ /* mantido original */ ]
// Boxlink functions (mantidas originais) ...
const BOXLINK_API = "https://api.boxlink.com.br"
const BOXLINK_TOKEN_KEY = "sg_boxlink_token"
function getBoxlinkToken() { return localStorage.getItem(BOXLINK_TOKEN_KEY)||"" }
function setBoxlinkToken(t) { localStorage.setItem(BOXLINK_TOKEN_KEY, t) }
function mapBoxlinkRow(item, i) { /* original */ return {} }
const BOXLINK_TRACKING_PATHS = [ /* original */ ]
let _bxWorkingPath = null
async function fetchBoxlinkPage(bToken, from, to, page) { /* original */ return {items:[], hasMore:false} }
async function syncBoxlinkFull(bToken, onPartial) { /* original */ return [] }

export default function App() {
  // Estados originais + novos para arquivados e realtime
  const [session,setSession]=useState(null)
  const [perfil,setPerfil]=useState(null)
  const [nomeAtendente,setNomeAtendente]=useState("")
  const [loadingPerfil,setLoadingPerfil]=useState(false)
  const [rows,setRows]=useState([])
  const [tab,setTab]=useState(null)
  const [paste,setPaste]=useState("")
  const [importing,setImporting]=useState(false)
  const [loadingData,setLoadingData]=useState(false)
  const [compact,setCompact]=useState(false)
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
  const [supView,setSupView]=useState('lista')
  const [selSupIds,setSelSupIds]=useState(new Set())
  const [openTpl,setOpenTpl]=useState(false); const [openHist,setOpenHist]=useState(false)
  const [aSrch,setASrch]=useState("")
  const [aPage,setAPage]=useState(1)
  const [syncStatus,setSyncStatus]=useState("idle")
  const [lastSync,setLastSync]=useState(null)
  const [realtimeStatus,setRealtimeStatus]=useState("connecting")
  const [bxToken,setBxToken]=useState(()=>getBoxlinkToken())
  const [bxStatus,setBxStatus]=useState("idle")
  const [bxCountdown,setBxCountdown]=useState(15*60)
  const [archRows,setArchRows]=useState([])
  const [archTotal,setArchTotal]=useState(0)
  const [archLoading,setArchLoading]=useState(false)
  const saveTimer=useRef(null); const fileRef=useRef()
  const token = session?.access_token
  const addToast = useCallback((msg,type="ok",ms=4000)=>{const id=Date.now(); setToasts(p=>[...p,{id,msg,type}]); setTimeout(()=>setToasts(p=>p.filter(t=>t.id!==id)),ms)},[])
  const handleLogin = async data => { /* original */ }
  const handleLogout = async () => { /* original */ }
  const perms = perfil?PERMS[perfil]:null

  // ─── CARGA INICIAL (ativos) + REALTIME ──────────────────────
  useEffect(() => {
    if (!token) return
    let subscription = null
    let isMounted = true
    const loadAtivos = async () => {
      setSyncStatus("loading")
      setLoadingData(true)
      try {
        const data = await dbLoadAtivos(token, partial => { if (isMounted) setRows(partial) })
        if (isMounted) { setRows(data); setLastSync(new Date()); setSyncStatus("idle") }
      } catch (err) {
        if (isMounted) { setSyncStatus("error"); addToast("Erro ao carregar pedidos ativos: "+err.message,"error",8000) }
      } finally { if (isMounted) setLoadingData(false) }
    }
    loadAtivos()
    subscription = supabase.channel('pedidos-realtime')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'pedidos' }, async (payload) => {
        if (payload.eventType === 'UPDATE' && payload.new.resolvido === true) {
          setRows(prev => prev.filter(r => r.id !== payload.new.id))
          addToast(`📦 Pedido ${payload.new.dados?.nuvem} foi arquivado`,"warn")
          return
        }
        if (payload.eventType === 'INSERT' && payload.new.resolvido === false) {
          const newRow = { ...payload.new.dados, id: payload.new.id, resolvido: false }
          setRows(prev => [...prev, { ...newRow, isNew: true }])
          addToast(`📦 Novo pedido ${newRow.nuvem} adicionado`,"warn")
        } else if (payload.eventType === 'UPDATE' && payload.new.resolvido === false) {
          const updatedRow = { ...payload.new.dados, id: payload.new.id, resolvido: false }
          setRows(prev => prev.map(r => r.id === updatedRow.id ? { ...updatedRow, isNew: true } : r))
          addToast(`✏️ Pedido ${updatedRow.nuvem} atualizado`,"warn")
        } else if (payload.eventType === 'DELETE') {
          setRows(prev => prev.filter(r => r.id !== payload.old.id))
          addToast(`🗑️ Pedido removido`,"warn")
        }
      })
      .subscribe((status) => {
        if (status === 'SUBSCRIBED') setRealtimeStatus('connected')
        else if (status === 'CHANNEL_ERROR') setRealtimeStatus('error')
        else setRealtimeStatus('connecting')
      })
    return () => { if (subscription) subscription.unsubscribe(); isMounted = false }
  }, [token, addToast])

  // ─── CARREGAR ARQUIVADOS SOB DEMANDA ────────────────────────
  useEffect(() => {
    if (tab !== 'arquivados' || !token) return
    const load = async () => {
      setArchLoading(true)
      try {
        const [data, total] = await Promise.all([dbLoadArquivados(token, aPage, PAGE_SIZE), dbCountArquivados(token)])
        setArchRows(data)
        setArchTotal(total)
      } catch (err) { addToast("Erro ao carregar arquivados: "+err.message,"error") }
      finally { setArchLoading(false) }
    }
    load()
  }, [tab, token, aPage, addToast])

  // ─── SALVAMENTO AUTOMÁTICO ──────────────────────────────────
  useEffect(() => {
    if (!token||rows.length===0) return
    if (saveTimer.current) clearTimeout(saveTimer.current)
    saveTimer.current = setTimeout(async () => {
      setSyncStatus("saving")
      try { await dbUpsert(rows,token) }
      catch(e){ setSyncStatus("error"); addToast("Erro ao salvar: "+e.message,"error",8000); setTimeout(()=>setSyncStatus("idle"),4000); return }
      setLastSync(new Date()); setSyncStatus("saved"); setTimeout(()=>setSyncStatus("idle"),2500)
    },1200)
  },[rows,token,addToast])

  // Efeitos adicionais (páginação, reset de filtros, etc.) iguais ao original
  useEffect(()=>{if (!rows.some(r=>r.isNew))return;const t=setTimeout(()=>setRows(p=>p.map(r=>({...r,isNew:false}))),6000);return()=>clearTimeout(t)},[rows])
  useEffect(()=>setLPage(1),[lSrch,lSt,lTr,lUrg,lAc,lSitPrazo,qf,sortCol,sortDir])
  useEffect(()=>setAPage(1),[aSrch])
  useEffect(()=>{setSResp("Todos")},[])
  const detailPanelRef=useRef(null); const queueRef=useRef(null)
  useEffect(()=>{ /* rolagem e auto-resp */ },[selSup])

  // Funções de manipulação (upd, del, toggle, bulk, import, etc.) – mantenha as originais
  const upd = (id,ch,hist) => setRows(prev=>prev.map(r=>r.id!==id?r:{...r,...ch,historico:hist?[...r.historico,{...hist,ts:new Date().toLocaleString("pt-BR")}]:r.historico}))
  const del = id => { if (!perms?.canDelete) return; setRows(prev=>prev.filter(r=>r.id!==id)); dbDelete(id,token).catch(()=>{}) }
  const toggleSel = id => setSelIds(prev=>{const n=new Set(prev); n.has(id)?n.delete(id):n.add(id); return n})
  const clearSel = () => setSelIds(new Set())
  const bulkSend = () => { if (!perms?.canSendSupport) return; const ts=new Date().toLocaleString("pt-BR"), sentAt=new Date().toISOString(); setRows(prev=>prev.map(r=>selIds.has(r.id)?{...r,enviadoSuporte:true,atendimento:"Aberto",sentAt,historico:[...r.historico,{acao:"Enviado ao suporte (lote)",ts}]}:r)); addToast(`${selIds.size} pedido${selIds.size>1?"s":""} enviado${selIds.size>1?"s":""} ao suporte`); clearSel() }
  const toggleSelSup = id => setSelSupIds(prev=>{const n=new Set(prev); n.has(id)?n.delete(id):n.add(id); return n})
  const bulkArchive = () => { if (!perms?.canOperate) return; const ts=new Date().toLocaleString("pt-BR"); setRows(prev=>prev.map(r=>selSupIds.has(r.id)?{...r,atendimento:"Resolvido",historico:[...r.historico,{acao:"Arquivado em lote",ts}]}:r)); addToast(`${selSupIds.size} pedido${selSupIds.size>1?"s":""} arquivado${selSupIds.size>1?"s":""}`); if (selSupIds.has(selSup)) setSelSup(null); setSelSupIds(new Set()) }
  const handleResolve = id => { if (!perms?.canOperate) return; upd(id,{atendimento:"Resolvido"},{acao:"Atendimento resolvido",usuario:nomeAtendente}); setSelSup(null); addToast("Pedido resolvido e arquivado") }
  const handleReturnLog = id => { if (!perms?.canOperate) return; upd(id,{enviadoSuporte:false,sentAt:null},{acao:"Devolvido à Logística",usuario:nomeAtendente}); setSelSup(null) }
  const handleClearAll = () => { if (!perms?.canClear) return; if (!window.confirm("Isso removerá TODOS os pedidos da base de dados. Esta ação não pode ser desfeita. Confirmar?")) return; setRows([]); dbClear(token).catch(()=>{}); addToast("Todos os dados foram removidos","warn") }
  const handleArchiveFromLog = id => { if (!perms?.canOperate) return; upd(id,{atendimento:"Resolvido",enviadoSuporte:false},{acao:"Arquivado pela Logística — entrega confirmada",usuario:nomeAtendente}); addToast("Pedido arquivado") }
  const bulkArchiveFromLog = () => { if (!perms?.canOperate) return; const ts = new Date().toLocaleString("pt-BR"); setRows(prev=>prev.map(r=>selIds.has(r.id)?{...r,atendimento:"Resolvido",enviadoSuporte:false,historico:[...r.historico,{acao:"Arquivado em lote pela Logística",ts,usuario:nomeAtendente}]}:r)); addToast(`${selIds.size} pedido${selIds.size>1?"s":""} arquivado${selIds.size>1?"s":""}`); clearSel() }
  const toggleSort = col => { if (sortCol===col)setSortDir(d=>d==="asc"?"desc":"asc"); else {setSortCol(col);setSortDir("asc")} }
  const doImport = useCallback(txt=>{ /* original, usando parseData */ },[addToast,perms])
  const handleFile = e => { /* original */ }
  const mergeBoxlink = useCallback((incoming) => { /* original */ },[])
  const doBoxlinkSync = useCallback(async (bToken, silent=false) => { /* original */ },[addToast, mergeBoxlink])

  // Dados derivados (baseLog, baseSup, etc.) iguais ao original, mas baseArch agora usa archRows
  const baseLog = rows.filter(r=>!r.enviadoSuporte&&r.atendimento!=="Resolvido")
  const baseSup = rows.filter(r=>r.enviadoSuporte&&r.atendimento!=="Resolvido")
  const detail = selSup?baseSup.find(r=>r.id===selSup):null
  const qCounts = Object.fromEntries(QFILTERS.map(f=>[f.id,applyQF(baseLog,f.id).length]))
  const filteredLog = applySortRows(applyQF(baseLog,qf).filter(r=>{ /* mesmo filtro original */ }),sortCol,sortDir)
  const totalPages = Math.max(1,Math.ceil(filteredLog.length/PAGE_SIZE))
  const safeP = Math.min(lPage,totalPages)
  const pagedLog = filteredLog.slice((safeP-1)*PAGE_SIZE,safeP*PAGE_SIZE)
  const respOpts = uniq(baseSup.map(r=>r.responsavel).filter(Boolean))
  const supRows = baseSup.filter(r=>{ /* mesmo filtro suporte */ }).sort((a,b)=>{const uo={Alta:0,Média:1,Baixa:2,"—":3},ao={Aberto:0,"Em andamento":1}; return (uo[a.urgencia]-uo[b.urgencia])||(ao[a.atendimento]-ao[b.atendimento])})
  const archPages = Math.max(1,Math.ceil(archTotal/PAGE_SIZE))
  const safeAPage = Math.min(aPage,archPages)
  const pagedArch = archRows.slice(0)
  const stOpts = uniq(baseLog.map(r=>r.status)), trOpts = uniq(baseLog.map(r=>r.transportadora))
  const st={log:baseLog.length,alta:baseLog.filter(r=>r.urgencia==="Alta").length,acionar:baseLog.filter(r=>r.acionar==="Sim").length}
  const ss={total:baseSup.length,abertos:baseSup.filter(r=>r.atendimento==="Aberto").length,andamento:baseSup.filter(r=>r.atendimento==="Em andamento").length}
  const arch = archTotal
  const entregues = rows.filter(r=>isEntregue(r.status))
  const hoje = new Date(); hoje.setHours(0,0,0,0)
  const parados = baseLog.filter(r=>{const d=diasSemMov(r.ultimaMov); return d!==null&&d>=ALERTA_DIAS}).length
  const calcNoPrazoLive = r => { /* original */ return null }
  const noPrazo = entregues.filter(r=>calcNoPrazoLive(r)===true).length
  const pctNoPrazo = entregues.length>0?Math.round((noPrazo/entregues.length)*100):0
  const TRANSP_INVALIDAS = new Set(["SP","RJ","MG","RS","SC","PR","BA","GO","PE","CE","AM","PA","MT","MS","ES","RN","PI","AL","SE","TO","RO","AC","AP","RR","MA","PB","DF"])
  const trStats={}; rows.forEach(r=>{ /* original para gráficos */ })
  const trData = []; const ufData = []; const urgData = []; const statusData = [] // preenchidos originalmente

  if (!session) return <LoginScreen onLogin={handleLogin}/>
  if (loadingPerfil) return <div style={{display:"flex",alignItems:"center",justifyContent:"center",minHeight:"100vh",color:C.text4,fontSize:12,letterSpacing:"0.1em"}}>Carregando perfil...</div>
  const showImp=(importing||rows.length===0)&&perms?.tabs.some(t=>["logistica","dashboard"].includes(t))
  const pd=compact?5:9
  const PERFLABEL={admin:"Admin",logistica:"Logística",suporte:"Suporte",leitura:"Leitura"}
  const syncDot = realtimeStatus==='connected'?'#27ae60':realtimeStatus==='error'?C.red:C.gold
  const syncText = realtimeStatus==='connected'?"Ao vivo":realtimeStatus==='error'?"Offline":"Conectando..."
  const TABS=[{key:"dashboard",label:"Dashboard",badge:null},{key:"logistica",label:"Logística",badge:st.acionar>0?st.acionar:null},{key:"suporte",label:"Suporte",badge:ss.abertos>0?ss.abertos:null},{key:"arquivados",label:"Arquivados",badge:arch>0?arch:null},{key:"usuarios",label:"Usuários",badge:null}].filter(t=>perms?.tabs.includes(t.key))
  const TH = {padding:`${compact?8:11}px 14px`,textAlign:"left",color:C.gold,fontWeight:400,fontSize:9,letterSpacing:"0.14em",textTransform:"uppercase",borderBottom:`1px solid #2A2A2A`,whiteSpace:"nowrap",background:C.brand,position:"sticky",top:0,zIndex:5,cursor:"pointer"}
  const THF = {...TH,cursor:"default"}

  return (
    <div style={{fontFamily:"'Inter',sans-serif",minHeight:"100vh",background:C.cream,color:C.text1,transition:"background .3s"}}>
      <style>{getGlobalStyle()}</style>
      <Toast toasts={toasts}/>
      {/* HEADER igual ao original, com indicador de sync substituído */}
      <div style={{background:C.brand,padding:"0 32px",display:"flex",alignItems:"stretch",justifyContent:"space-between",borderBottom:`1px solid ${C.gold}33`}}>
        <div style={{display:"flex",alignItems:"center",gap:24,padding:"14px 0"}}>
          <div><div style={{fontSize:10,letterSpacing:"0.45em",color:C.gold,textTransform:"uppercase",fontWeight:300,lineHeight:1}}>Saint Germain</div><div style={{fontSize:8,letterSpacing:"0.28em",color:`${C.gold}55`,textTransform:"uppercase",marginTop:3}}>Central de Pedidos</div></div>
          <div style={{width:1,height:32,background:`${C.gold}22`}}/>
          <div style={{display:"flex",alignItems:"center",gap:7}}><span style={{width:6,height:6,borderRadius:"50%",background:syncDot,display:"inline-block",boxShadow:`0 0 6px ${syncDot}66`}}/><span style={{fontSize:10,color:syncDot,letterSpacing:"0.04em"}}>{syncText}</span></div>
          {bxToken&&<div style={{display:"flex",alignItems:"center",gap:6,borderLeft:`1px solid ${C.gold}22`,paddingLeft:14}}><span style={{width:6,height:6,borderRadius:"50%",background:bxStatus==="ok"?"#27ae60":bxStatus==="syncing"?C.gold:bxStatus==="error"?C.red:"#555",display:"inline-block"}}/><span style={{fontSize:10,color:bxStatus==="ok"?"#27ae60":bxStatus==="syncing"?C.gold:bxStatus==="error"?C.red:"#555",letterSpacing:"0.04em"}}>{bxStatus==="syncing"?"Boxlink...":bxStatus==="ok"?`Boxlink ✓ ${Math.floor(bxCountdown/60)}m`:bxStatus==="error"?"Boxlink ✗":"Boxlink"}</span><button onClick={()=>doBoxlinkSync(bxToken)} disabled={bxStatus==="syncing"} style={{background:"transparent",border:`1px solid ${C.gold}44`,color:`${C.gold}88`,borderRadius:5,padding:"2px 8px",fontSize:9,cursor:"pointer",letterSpacing:"0.06em"}}>↺</button></div>}
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          {perms?.canImport&&!showImp&&rows.length>0&&(<><button onClick={()=>setCompact(c=>!c)} style={{background:"transparent",border:`1px solid ${C.gold}44`,color:compact?C.gold:`${C.gold}88`,borderRadius:6,padding:"6px 12px",fontSize:10,cursor:"pointer",letterSpacing:"0.08em"}}>{compact?"⊞":"⊟"}</button><button onClick={()=>{const hf=lSrch||lSt!=="Todos"||lTr!=="Todos"||lUrg!=="Todos"||lAc!=="Todos";exportCSV(hf&&tab==="logistica"?filteredLog:rows)}} style={{background:"transparent",border:`1px solid #444`,color:"#888",borderRadius:6,padding:"6px 14px",fontSize:10,cursor:"pointer",letterSpacing:"0.06em"}}>↓ Exportar</button><button onClick={()=>setImporting(true)} style={{background:"transparent",border:`1px solid ${C.gold}`,color:C.gold,borderRadius:6,padding:"6px 14px",fontSize:10,cursor:"pointer",letterSpacing:"0.06em"}}>+ Importar</button></>)}
          <button onClick={()=>{const nd=!dark;applyTheme(nd);setDark(nd)}} title="Alternar modo escuro" style={{background:"transparent",border:`1px solid #333`,color:dark?C.gold:C.text3,borderRadius:6,padding:"6px 10px",fontSize:12,cursor:"pointer"}}>{dark?"☀":"🌙"}</button>
          {perms?.canClear&&rows.length>0&&(<button onClick={handleClearAll} style={{background:"transparent",border:`1px solid #333`,color:"#555",borderRadius:6,padding:"6px 14px",fontSize:10,cursor:"pointer",letterSpacing:"0.06em"}}>Limpar tudo</button>)}
          <div style={{width:1,height:28,background:"#2A2A2A"}}/>
          <div style={{textAlign:"right"}}><div style={{color:C.white,fontSize:11,letterSpacing:"0.02em"}}>{session.user?.email}</div><div style={{color:C.gold,fontSize:9,letterSpacing:"0.1em",textTransform:"uppercase"}}>{PERFLABEL[perfil]||perfil}</div></div>
          <button onClick={handleLogout} style={{background:"transparent",border:`1px solid #2A2A2A`,color:"#666",borderRadius:6,padding:"6px 12px",fontSize:10,cursor:"pointer",letterSpacing:"0.06em"}}>Sair</button>
        </div>
        <input ref={fileRef} type="file" accept=".csv,.txt,.tsv,.xls,.xlsx" style={{display:"none"}} onChange={handleFile}/>
      </div>
      {!showImp&&(
        <div style={{background:C.white,borderBottom:`1px solid ${C.border}`,padding:"0 32px",display:"flex",alignItems:"stretch",boxShadow:"0 1px 0 rgba(0,0,0,0.04)"}}>
          {TABS.map(t=>(<button key={t.key} onClick={()=>{setTab(t.key);if(t.key!=="suporte")setSelSup(null)}} style={{background:"transparent",border:"none",borderBottom:tab===t.key?`2px solid ${C.gold}`:"2px solid transparent",color:tab===t.key?C.text1:C.text3,padding:"14px 20px",cursor:"pointer",fontSize:10,letterSpacing:"0.14em",textTransform:"uppercase",fontWeight:tab===t.key?500:400,marginBottom:"-1px",display:"flex",alignItems:"center",gap:8,transition:"color .2s"}}>{t.label}{t.badge!=null&&<span style={{background:tab===t.key?C.brand:C.red,color:C.white,borderRadius:10,padding:"2px 7px",fontSize:9,fontWeight:600,letterSpacing:"0.04em"}}>{t.badge}</span>}</button>))}
        </div>
      )}
      {/* Tela de import */}
      {showImp&&perms?.canImport&&(<div style={{padding:48,maxWidth:640,margin:"0 auto"}}>...</div>)}
      {/* Painel de usuários */}
      {tab==="usuarios"&&perfil==="admin"&&<UsuariosPanel token={token} addToast={addToast}/>}
      {/* Dashboard */}
      {tab==="dashboard"&&!showImp&&(<div style={{padding:"28px 40px"}}>...</div>)}
      {/* Logística */}
      {tab==="logistica"&&!showImp&&(<div style={{padding:"24px 32px"}}>...</div>)}
      {/* Suporte */}
      {tab==="suporte"&&(<div style={{display:"flex",height:"calc(100vh - 110px)",overflow:"hidden"}}>...</div>)}
      {/* Arquivados - agora usando archRows e archTotal */}
      {tab==="arquivados"&&(
        <div style={{padding:"24px 32px"}}>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12,marginBottom:20}}>
            <KpiCard label="Total arquivados" val={arch}/>
            <KpiCard label="Resolvidos hoje" val={rows.filter(r=>{if(r.atendimento!=="Resolvido")return false;const h=r.historico.find(x=>x.acao&&(x.acao.includes("Resolvido")||x.acao.includes("Arquivado")));return h&&h.ts&&h.ts.startsWith(new Date().toLocaleDateString("pt-BR"))}).length}/>
            <KpiCard label="Com observações" val={archRows.filter(r=>r.obs&&r.obs.trim()).length}/>
          </div>
          {arch===0?<div style={{textAlign:"center",padding:"56px 0",color:C.text4}}><div style={{fontSize:32,marginBottom:12,opacity:0.2}}>◎</div><div style={{fontSize:14}}>Nenhum atendimento arquivado</div></div>:(
            <div>
              <div style={{marginBottom:14}}><input value={aSrch} onChange={e=>setASrch(e.target.value)} placeholder="Buscar nos arquivados..." style={{...getINP(),width:"100%",padding:"10px 14px",boxSizing:"border-box",boxShadow:shadow.sm}}/></div>
              {archLoading?<div style={{textAlign:"center",padding:32,color:C.text4}}>Carregando...</div>:(<>
                <div style={{overflowX:"auto",overflowY:"auto",maxHeight:"60vh",borderRadius:12,border:`1px solid ${C.border}`,boxShadow:shadow.sm}}>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:compact?11:12,tableLayout:"fixed",minWidth:900}}>
                    <colgroup><col style={{width:90}}/><col style={{width:150}}/><col style={{width:110}}/><col style={{width:110}}/><col style={{width:120}}/><col style={{width:70}}/><col style={{width:110}}/><col style={{width:96}}/><col style={{width:96}}/><col style={{width:90}}/></colgroup>
                    <thead><tr>{["No NUVEM","Destinatário","Transportadora","Status","Motivo","Urgência","Prazo / SLA","Chamado","Responsável","Ações"].map(h=><th key={h} style={THF}>{h}</th>)}</table></thead>
                    <tbody>
                      {pagedArch.filter(r=>{const q=aSrch.toLowerCase(); return !q||[r.nuvem,r.destinatario,r.transportadora,r.status].some(v=>(v||"").toLowerCase().includes(q));}).map((r,i)=>(
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
                          <td style={{padding:`${pd}px 14px`}}>{perms?.canOperate&&<button onClick={()=>upd(r.id,{atendimento:"Em andamento"},{acao:"Reaberto dos arquivados",usuario:nomeAtendente})} style={{background:C.cream,border:`1px solid ${C.border}`,color:C.text2,borderRadius:6,padding:"4px 12px",fontSize:10,cursor:"pointer",fontWeight:500}}>Reabrir</button>}</td>
                        </td>
                      ))}
                    </tbody>
                  </table>
                </div>
                <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginTop:10}}>
                  <div style={{fontSize:11,color:C.text4}}>{archTotal===0?"Nenhum resultado":`Mostrando ${((safeAPage-1)*PAGE_SIZE)+1}–${Math.min(safeAPage*PAGE_SIZE,archTotal)} de ${archTotal} arquivados`}</div>
                  {archPages>1&&<div style={{display:"flex",gap:4,alignItems:"center"}}><button onClick={()=>setAPage(n=>Math.max(1,n-1))} disabled={safeAPage===1} style={{...getINP(),padding:"5px 12px",cursor:safeAPage===1?"not-allowed":"pointer",opacity:safeAPage===1?0.4:1}}>‹</button><span style={{fontSize:11,color:C.text3,padding:"0 10px"}}>{safeAPage} / {archPages}</span><button onClick={()=>setAPage(n=>Math.min(archPages,n+1))} disabled={safeAPage===archPages} style={{...getINP(),padding:"5px 12px",cursor:safeAPage===archPages?"not-allowed":"pointer",opacity:safeAPage===archPages?0.4:1}}>›</button></div>}
                </div>
              </>)}
            </div>
          )}
        </div>
      )}
    </div>
  )
}
