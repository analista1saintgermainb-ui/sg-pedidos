import { useState, useEffect, useRef, useCallback } from 'react'
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell } from 'recharts'
import * as XLSX from 'xlsx'

const GOLD = "#C4963A", BRAND = "#000", PAGE_SIZE = 50
const PIE_COLORS = { Alta: "#e74c3c", Média: GOLD, Baixa: "#27ae60", "—": "#aaa" }
const SUPA_URL = "https://jdiuuhfhsiymttxllssr.supabase.co"
const SUPA_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpkaXV1aGZoc2l5bXR0eGxsc3NyIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzc3NTMyNTcsImV4cCI6MjA5MzMyOTI1N30.wNGhwh2bCF0HZSonn09S-15kEVAQGzEP1yWvRx3l_N4"
const SUPA_SERVICE = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpkaXV1aGZoc2l5bXR0eGxsc3NyIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc3Nzc1MzI1NywiZXhwIjoyMDkzMzI5MjU3fQ.yjZ8VKr8YfbMBELdoKevdE1k_dd2OXUlYjUj4n2GeQw"
const SH = { apikey: SUPA_KEY, "Content-Type": "application/json" }
const authSH = (token) => ({ ...SH, Authorization: `Bearer ${token}` })

// ─── Auth ────────────────────────────────────────────────────
async function signIn(email, password) {
  const r = await fetch(`${SUPA_URL}/auth/v1/token?grant_type=password`, {
    method: "POST", headers: SH, body: JSON.stringify({ email, password })
  })
  const data = await r.json()
  if (!r.ok) throw new Error(data.error_description || data.msg || "Erro ao fazer login")
  return data
}
async function signOut(token) {
  await fetch(`${SUPA_URL}/auth/v1/logout`, { method: "POST", headers: authSH(token) })
}
async function createUser(email, password) {
  const r = await fetch(`${SUPA_URL}/auth/v1/admin/users`, {
    method: "POST",
    headers: { apikey: SUPA_SERVICE, Authorization: `Bearer ${SUPA_SERVICE}`, "Content-Type": "application/json" },
    body: JSON.stringify({ email, password, email_confirm: true })
  })
  const data = await r.json()
  if (!r.ok) throw new Error(data.msg || data.message || "Erro ao criar usuário")
  return data
}

// ─── DB ──────────────────────────────────────────────────────
async function dbLoad(token) {
  let all = [], from = 0, step = 1000
  while (true) {
    const r = await fetch(`${SUPA_URL}/rest/v1/pedidos?select=*&order=id&limit=${step}&offset=${from}`, { headers: authSH(token) })
    if (!r.ok) throw new Error(await r.text())
    const data = await r.json()
    all = [...all, ...data.map(row => ({ ...row.dados, id: row.id }))]
    if (data.length < step) break
    from += step
  }
  return all
}
async function dbUpsert(rows, token) {
  if (!rows.length) return
  for (let i = 0; i < rows.length; i += 200) {
    const batch = rows.slice(i, i + 200)
    const payload = batch.map(r => ({ id: Number(r.id), dados: r, updated_at: new Date().toISOString() }))
    const r2 = await fetch(`${SUPA_URL}/rest/v1/pedidos`, {
      method: "POST", headers: { ...authSH(token), Prefer: "resolution=merge-duplicates,return=minimal" },
      body: JSON.stringify(payload)
    })
    if (!r2.ok) throw new Error(`${r2.status}: ${await r2.text()}`)
  }
}
async function dbDelete(id, token) {
  await fetch(`${SUPA_URL}/rest/v1/pedidos?id=eq.${id}`, { method: "DELETE", headers: authSH(token) })
}
async function dbClear(token) {
  await fetch(`${SUPA_URL}/rest/v1/pedidos?id=gte.0`, { method: "DELETE", headers: authSH(token) })
}
async function loadUsuarios(token) {
  const r = await fetch(`${SUPA_URL}/rest/v1/usuarios?select=*&order=created_at`, { headers: authSH(token) })
  return r.ok ? r.json() : []
}
async function saveUsuario(u, token) {
  const r = await fetch(`${SUPA_URL}/rest/v1/usuarios`, {
    method: "POST", headers: { ...authSH(token), Prefer: "resolution=merge-duplicates,return=minimal" },
    body: JSON.stringify(u)
  })
  if (!r.ok) throw new Error(await r.text())
}
async function deleteUsuario(id, token) {
  await fetch(`${SUPA_URL}/rest/v1/usuarios?id=eq.${id}`, { method: "DELETE", headers: authSH(token) })
}

// ─── Permissões ───────────────────────────────────────────────
const PERMS = {
  admin:     { tabs: ["dashboard","logistica","suporte","arquivados","usuarios"], canImport: true, canDelete: true, canClear: true, canSendSupport: true, canOperate: true },
  logistica: { tabs: ["dashboard","logistica"],                                   canImport: true, canDelete: false,canClear: false,canSendSupport: true, canOperate: true },
  suporte:   { tabs: ["suporte","arquivados"],                                    canImport: false,canDelete: false,canClear: false,canSendSupport: false,canOperate: true },
  leitura:   { tabs: ["dashboard","logistica","suporte","arquivados"],             canImport: false,canDelete: false,canClear: false,canSendSupport: false,canOperate: false },
}

// ─── Helpers idênticos ao original ───────────────────────────
const HEADER_MAP = {
  nuvem: ["identificador ecommerce","id ecommerce","no nuvem","nuvem","pedido"],
  destinatario: ["destinatário nome","destinatario nome","destinatário","destinatario","nome do pedido","nome do cliente","nome","cliente"],
  transportadora: ["estratégia de frete","estrategia de frete","transportadora","frete"],
  rastreio: ["rastreador last mile","código de rastreio","codigo de rastreio","rastreio","last mile"],
  status: ["situação","situacao","situac","status"],
  prazo: ["prazo logístico","prazo logistico","prazo"],
  nf: ["nº nota fiscal","no nota fiscal","nota fiscal","no nf","nf"],
}
const norm = s => (s||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().trim()
const findIdx = (hdrs, key) => hdrs.findIndex(h => (HEADER_MAP[key]||[]).some(v => norm(h).includes(norm(v))))
const urgMap   = { Alta:{bg:"#f8d7da",color:"#721c24",bd:"#f5c6cb"}, Média:{bg:"#fff3cd",color:"#856404",bd:"#ffeeba"}, Baixa:{bg:"#d4edda",color:"#155724",bd:"#c3e6cb"}, "—":{bg:"#ebebeb",color:"#555",bd:"#ddd"} }
const acionMap = { Sim:{bg:"#f8d7da",color:"#721c24",bd:"#f5c6cb"}, Avaliar:{bg:"#fff3cd",color:"#856404",bd:"#ffeeba"}, Não:{bg:"#d4edda",color:"#155724",bd:"#c3e6cb"} }
const atendMap = { Aberto:{bg:"#f8d7da",color:"#721c24",bd:"#f5c6cb"}, "Em andamento":{bg:"#fff3cd",color:"#856404",bd:"#ffeeba"}, Resolvido:{bg:"#d4edda",color:"#155724",bd:"#c3e6cb"} }
const SEL = { borderRadius:3, border:"1px solid #d5d5d5", padding:"6px 9px", fontSize:12, background:"#fff", color:"#000" }
const QFILTERS = [{id:"todos",label:"Todos"},{id:"urgente",label:"Urgente"},{id:"extraviados",label:"Extraviados"},{id:"devolvidos",label:"Devolvidos"},{id:"vence_hoje",label:"Vence hoje"},{id:"vencidos",label:"Vencidos"}]
const uniq = arr => ["Todos", ...Array.from(new Set(arr.filter(Boolean).sort()))]

function calcMotivo(s) {
  const v = (s||"").toLowerCase()
  if (v.includes("extravia")||v.includes("perdid")) return "Objeto extraviado"
  if (v.includes("devolv")||v.includes("recusa"))   return "Devolução / Recusa"
  if (v.includes("atras"))                           return "Atraso na entrega"
  if (v.includes("entregue")||v.includes("finaliz")) return "Entrega concluída"
  if (v.includes("saiu"))                            return "Saiu para entrega"
  if (v.includes("trânsito")||v.includes("transito")) return "Em trânsito"
  if (v.includes("postado")||v.includes("coletado")) return "Aguardando movimentação"
  if (v.includes("pendente")||v.includes("aguardando")||v.includes("processando")) return "Aguardando coleta"
  return "—"
}
function parsePrazo(v) {
  if (!v) return null
  const c = v.replace(/[^\d\/\-\.]/g,"")
  for (const t of [c, c.split("/").reverse().join("-"), c.split(".").reverse().join("-")]) {
    const d = new Date(t); if (!isNaN(d.getTime())) return d
  }
  return null
}
function calcUrg(prazo, status) {
  const s = (status||"").toLowerCase()
  if (s.includes("extravia")||s.includes("devolv")) return "Alta"
  if (s.includes("entregue")||s.includes("finaliz")) return "Baixa"
  const dt = parsePrazo(prazo); if (!dt) return "—"
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
  if (diff<0)   return {label:`${Math.abs(diff)}d vencido`,color:"#721c24",bg:"#f8d7da"}
  if (diff===0)  return {label:"Vence hoje",              color:"#7d3c00",bg:"#fde8cc"}
  if (diff<=3)   return {label:`${diff}d restantes`,      color:"#856404",bg:"#fff3cd"}
  return              {label:`${diff}d restantes`,         color:"#155724",bg:"#d4edda"}
}
function timeOpen(sentAt) {
  if (!sentAt) return null
  const ms = Date.now()-new Date(sentAt).getTime()
  const h = Math.floor(ms/3600000), d = Math.floor(h/24)
  if (d>0) return {label:`${d}d na fila`,alert:d>=2}
  if (h>0) return {label:`${h}h na fila`,alert:false}
  return {label:"< 1h",alert:false}
}
function rowBg(urg, compact, i) {
  const lo=compact?0.04:0.05, hi=compact?0.07:0.09
  if (urg==="Alta")  return `rgba(231,76,60,${i%2===0?lo:hi})`
  if (urg==="Média") return `rgba(230,162,0,${i%2===0?lo:hi})`
  return i%2===0?"#fff":"#fafafa"
}
function isEntregue(status) {
  const s = (status||"").toLowerCase()
  return s.includes("entregue") || s.includes("finaliz") || s.includes("entrega realizada")
}
function calcEntregueNoPrazo(prazo, dataEntrega) {
  const dt = parsePrazo(prazo)
  const de = parsePrazo(dataEntrega)
  if (!dt) return null
  const ref = de || new Date()
  return ref <= dt
}
function parseData(text) {
  const sep = text.includes("\t")?"\t":text.includes(";")?";":","
  const lines = text.trim().split("\n").filter(l=>l.trim())
  if (!lines.length) return []
  const first = lines[0].split(sep).map(h=>h.trim().replace(/^["']|["']$/g,""))
  const isHdr = first.some(h=>["nuvem","destinat","identificador","ecommerce","rastreio","situac","status","frete"].some(k=>norm(h).includes(k)))
  const hdrs = isHdr?first:[]
  const data = isHdr?lines.slice(1):lines
  const ix = {
    nuvem:   isHdr?findIdx(hdrs,"nuvem"):0,
    dest:    isHdr?findIdx(hdrs,"destinatario"):1,
    transp:  isHdr?findIdx(hdrs,"transportadora"):2,
    rastreio:isHdr?findIdx(hdrs,"rastreio"):3,
    status:  isHdr?findIdx(hdrs,"status"):4,
    prazo:   isHdr?findIdx(hdrs,"prazo"):5,
    nf:      isHdr?findIdx(hdrs,"nf"):6,
  }
  const g = (c,i) => i>=0&&i<c.length?c[i]:""
  return data.map((line,i) => {
    const c = line.split(sep).map(v=>v.trim().replace(/^["']|["']$/g,""))
    const status=g(c,ix.status), prazo=g(c,ix.prazo), urg=calcUrg(prazo,status)
    const entregue = isEntregue(status)
    const noPrazo = entregue ? calcEntregueNoPrazo(prazo, null) : null
    return {
      id: Date.now()+i,
      nuvem: g(c,ix.nuvem), destinatario: g(c,ix.dest), transportadora: g(c,ix.transp),
      rastreio: g(c,ix.rastreio), status, prazo, nf: g(c,ix.nf),
      motivo: calcMotivo(status), urgencia: urg, acionar: calcAcionar(urg,status),
      // Auto-arquiva entregues
      enviadoSuporte: false,
      atendimento: entregue ? "Resolvido" : "Aberto",
      entregueNoPrazo: noPrazo,
      obs:"", historico: entregue ? [{acao:"Arquivado automaticamente — entrega concluída", ts: new Date().toLocaleString("pt-BR")}] : [],
      responsavel:"", sentAt:null, chamado:"", isNew:true
    }
  }).filter(r=>r.nuvem||r.destinatario||r.nf)
}
function applyQF(rows, qf) {
  if (qf==="todos") return rows
  if (qf==="urgente") return rows.filter(r=>r.urgencia==="Alta")
  if (qf==="extraviados") return rows.filter(r=>(r.status||"").toLowerCase().includes("extravia"))
  if (qf==="devolvidos") return rows.filter(r=>(r.status||"").toLowerCase().includes("devolv"))
  if (qf==="vence_hoje") return rows.filter(r=>{const d=parsePrazo(r.prazo);if(!d)return false;const h=new Date();h.setHours(0,0,0,0);return Math.ceil((d-h)/86400000)<=0})
  if (qf==="vencidos") return rows.filter(r=>{const d=parsePrazo(r.prazo);if(!d)return false;const h=new Date();h.setHours(0,0,0,0);return d<h})
  return rows
}
function applySortRows(rows, col, dir) {
  if (!col) return rows
  return [...rows].sort((a,b) => {
    let va=a[col]||"", vb=b[col]||""
    if (col==="prazo") { va=parsePrazo(va)||new Date(0); vb=parsePrazo(vb)||new Date(0) }
    if (col==="urgencia") { const o={Alta:0,Média:1,Baixa:2,"—":3}; va=o[va]??9; vb=o[vb]??9 }
    const cmp = typeof va==="object"?va-vb:String(va).localeCompare(String(vb),"pt-BR")
    return dir==="asc"?cmp:-cmp
  })
}
function exportCSV(rows) {
  const h = ["No NUVEM","Destinatário","Transportadora","Cód. Rastreio","Status","Prazo","No NF","Motivo","Urgência","Acionar?","Suporte","Atendimento","Chamado","Responsável","Observações"]
  const e = v => `"${String(v||"").replace(/"/g,'""')}"`
  const csv = [h.map(e).join(";"), ...rows.map(r=>[r.nuvem,r.destinatario,r.transportadora,r.rastreio,r.status,r.prazo,r.nf,r.motivo,r.urgencia,r.acionar,r.enviadoSuporte?"Sim":"Não",r.atendimento,r.chamado,r.responsavel,r.obs].map(e).join(";"))].join("\n")
  const a = document.createElement("a")
  a.href = URL.createObjectURL(new Blob(["\uFEFF"+csv],{type:"text/csv;charset=utf-8;"}))
  a.download = "saint_germain_pedidos.csv"; a.click()
}
function getTemplate(r, ch) {
  const nome = (r.destinatario||"").split(" ")[0]||"Cliente"
  const m = (r.motivo||"").toLowerCase()
  const extrav=m.includes("extravia"), devolv=m.includes("devolu")||m.includes("recusa"), atraso=m.includes("atraso")
  const det = `• Pedido: #${r.nuvem}\n• NF: ${r.nf}\n• Transportadora: ${r.transportadora}\n• Rastreio: ${r.rastreio||"—"}\n• Prazo: ${r.prazo||"—"}`
  if (ch==="wpp") {
    if (extrav) return `Olá, ${nome}! 😊\n\nAqui é a equipe *Saint Germain*. Sua encomenda está com status de *objeto extraviado* junto à *${r.transportadora}*. Já acionamos nossa equipe.\n\nRetornaremos em até *2 dias úteis*. Pedimos desculpas! 🙏`
    if (devolv) return `Olá, ${nome}! 😊\n\nAqui é a equipe *Saint Germain*. Sua encomenda foi *devolvida* ao nosso CD. Poderia confirmar o endereço para novo envio sem custo? 📦`
    if (atraso) return `Olá, ${nome}! 😊\n\nAqui é a equipe *Saint Germain*. Identificamos um atraso no pedido *#${r.nuvem}* pela *${r.transportadora}*. Estamos acompanhando! 🙏`
    return `Olá, ${nome}! 😊\n\nAqui é a equipe *Saint Germain*. Atualizando sobre o pedido *#${r.nuvem}*.\n\nStatus: *${r.status}*${r.prazo?`\nPrazo: *${r.prazo}*`:""}\n\nQualquer dúvida é só chamar! ✨`
  } else {
    if (extrav) return `Assunto: Pedido #${r.nuvem} — Objeto Extraviado\n\nOlá, ${r.destinatario},\n\n${det}\n\nNossa equipe está apurando com a transportadora. Retornamos em até 2 dias úteis.\n\nAtenciosamente,\nEquipe Saint Germain`
    if (devolv) return `Assunto: Pedido #${r.nuvem} — Devolução\n\nOlá, ${r.destinatario},\n\nSua encomenda retornou ao CD.\n\n${det}\n\nConfirme seu endereço para novo envio sem custo.\n\nAtenciosamente,\nEquipe Saint Germain`
    if (atraso) return `Assunto: Pedido #${r.nuvem} — Atraso\n\nOlá, ${r.destinatario},\n\n${det}\n\nEstamos acompanhando junto à transportadora.\n\nAtenciosamente,\nEquipe Saint Germain`
    return `Assunto: Pedido #${r.nuvem} — Atualização\n\nOlá, ${r.destinatario},\n\n${det}\nStatus: ${r.status}\n\nQualquer dúvida, responda este chamado.\n\nAtenciosamente,\nEquipe Saint Germain`
  }
}

// ─── Componentes UI ───────────────────────────────────────────
function Chip({ val, map }) {
  const s = map[val]||{bg:"#ebebeb",color:"#555",bd:"#ddd"}
  return <span style={{background:s.bg,color:s.color,border:`1px solid ${s.bd}`,borderRadius:3,padding:"2px 8px",fontSize:11,fontWeight:500,whiteSpace:"nowrap"}}>{val}</span>
}
function StatusBadge({ val }) {
  const s=(val||"").toLowerCase(); let bg="#ebebeb",color="#555",bd="#ddd"
  if(s.includes("entregue")||s.includes("finaliz")){bg="#d4edda";color="#155724";bd="#c3e6cb"}
  else if(s.includes("trânsito")||s.includes("transito")){bg="#d1ecf1";color="#0c5460";bd="#bee5eb"}
  else if(s.includes("saiu")){bg="#cce5ff";color="#004085";bd="#b8daff"}
  else if(s.includes("extravia")){bg="#f8d7da";color="#721c24";bd="#f5c6cb"}
  else if(s.includes("devolv")){bg="#fff3cd";color="#856404";bd="#ffeeba"}
  return <span style={{background:bg,color,border:`1px solid ${bd}`,borderRadius:3,padding:"2px 8px",fontSize:11,fontWeight:500,whiteSpace:"nowrap"}}>{val||"—"}</span>
}
function SlaCell({ prazo }) {
  const sla = slaInfo(prazo)
  return <div style={{lineHeight:1.5}}><div style={{fontSize:11,color:"#888"}}>{prazo||"—"}</div>{sla&&<span style={{background:sla.bg,color:sla.color,borderRadius:2,padding:"1px 5px",fontSize:10,fontWeight:500}}>{sla.label}</span>}</div>
}
function TimeOpenBadge({ sentAt }) {
  const info = timeOpen(sentAt); if (!info) return null
  return <span style={{background:info.alert?"#f8d7da":"#fff3cd",color:info.alert?"#721c24":"#856404",borderRadius:3,padding:"2px 7px",fontSize:10,fontWeight:500}}>{info.label}</span>
}
function StatCard({ label, val, accent, sub }) {
  return <div style={{background:"#fff",borderRadius:6,padding:"14px 16px",border:accent?"1px solid #f5c6cb":"1px solid #e8e8e8"}}><div style={{fontSize:10,color:"#999",marginBottom:4,letterSpacing:"0.05em",textTransform:"uppercase"}}>{label}</div><div style={{fontSize:26,fontWeight:500,color:accent?"#c0392b":GOLD}}>{val}</div>{sub&&<div style={{fontSize:11,color:"#aaa",marginTop:2}}>{sub}</div>}</div>
}
function CopyBtn({ text, label }) {
  const [ok, setOk] = useState(false)
  return <button onClick={()=>{navigator.clipboard.writeText(text);setOk(true);setTimeout(()=>setOk(false),2000)}} style={{background:ok?"#d4edda":GOLD,border:"none",color:ok?"#155724":"#111",borderRadius:3,padding:"4px 12px",fontSize:11,cursor:"pointer",fontWeight:500,whiteSpace:"nowrap"}}>{ok?"✓ Copiado!":label||"Copiar"}</button>
}
function SortIcon({ col, sortCol, sortDir }) {
  if (sortCol!==col) return <span style={{color:"#555",fontSize:9,marginLeft:3}}>⇅</span>
  return <span style={{color:GOLD,fontSize:9,marginLeft:3}}>{sortDir==="asc"?"↑":"↓"}</span>
}
function Toast({ toasts }) {
  return <div style={{position:"fixed",bottom:20,right:20,display:"flex",flexDirection:"column",gap:8,zIndex:9999,pointerEvents:"none"}}>{toasts.map(t=><div key={t.id} style={{background:t.type==="error"?"#c0392b":t.type==="warn"?GOLD:"#27ae60",color:"#fff",borderRadius:4,padding:"10px 16px",fontSize:12,fontWeight:500,boxShadow:"0 4px 12px rgba(0,0,0,0.15)",maxWidth:300,lineHeight:1.4}}>{t.msg}</div>)}</div>
}

// ─── Tela de Login ────────────────────────────────────────────
function LoginScreen({ onLogin }) {
  const [email, setEmail]       = useState("")
  const [password, setPassword] = useState("")
  const [loading, setLoading]   = useState(false)
  const [error, setError]       = useState("")

  const handleLogin = async e => {
    e.preventDefault(); setLoading(true); setError("")
    try {
      const data = await signIn(email, password)
      onLogin(data)
    } catch(err) {
      setError(err.message)
    } finally { setLoading(false) }
  }

  return (
    <div style={{minHeight:"100vh",background:"#fff",display:"flex",flexDirection:"column"}}>
      <div style={{background:BRAND,padding:"18px 32px"}}>
        <div style={{color:"#fff",fontSize:15,fontWeight:400,letterSpacing:"0.38em",textTransform:"uppercase"}}>Saint Germain</div>
        <div style={{color:"#555",fontSize:9,letterSpacing:"0.28em",textTransform:"uppercase",marginTop:2}}>Central de Pedidos</div>
      </div>
      <div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center"}}>
        <div style={{width:"100%",maxWidth:380,padding:"0 24px"}}>
          <div style={{textAlign:"center",marginBottom:32}}>
            <div style={{color:GOLD,fontSize:11,letterSpacing:"0.2em",textTransform:"uppercase",marginBottom:8}}>Acesso restrito</div>
            <div style={{fontSize:22,fontWeight:400,color:"#000"}}>Entrar na plataforma</div>
          </div>
          <form onSubmit={handleLogin}>
            <div style={{marginBottom:16}}>
              <div style={{fontSize:10,color:"#999",textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:6}}>E-mail</div>
              <input value={email} onChange={e=>setEmail(e.target.value)} type="email" placeholder="seu@email.com" required
                style={{width:"100%",borderRadius:3,border:"1px solid #d5d5d5",padding:"10px 12px",fontSize:14,boxSizing:"border-box"}}/>
            </div>
            <div style={{marginBottom:24}}>
              <div style={{fontSize:10,color:"#999",textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:6}}>Senha</div>
              <input value={password} onChange={e=>setPassword(e.target.value)} type="password" placeholder="••••••••" required
                style={{width:"100%",borderRadius:3,border:"1px solid #d5d5d5",padding:"10px 12px",fontSize:14,boxSizing:"border-box"}}/>
            </div>
            {error && <div style={{background:"#f8d7da",color:"#721c24",borderRadius:3,padding:"10px 12px",fontSize:12,marginBottom:16}}>{error}</div>}
            <button type="submit" disabled={loading} style={{width:"100%",background:BRAND,border:"none",color:"#fff",borderRadius:3,padding:"12px 0",fontSize:13,fontWeight:500,cursor:loading?"not-allowed":"pointer",opacity:loading?0.6:1,letterSpacing:"0.08em",textTransform:"uppercase"}}>
              {loading?"Entrando...":"Entrar"}
            </button>
          </form>
        </div>
      </div>
    </div>
  )
}

// ─── Painel de Usuários (Admin) ───────────────────────────────
function UsuariosPanel({ token, addToast }) {
  const [usuarios, setUsuarios] = useState([])
  const [loading, setLoading]   = useState(true)
  const [form, setForm]         = useState({ email:"", nome:"", perfil:"logistica", senha:"" })
  const [saving, setSaving]     = useState(false)

  useEffect(() => {
    loadUsuarios(token).then(data => { setUsuarios(data); setLoading(false) })
  }, [])

  const handleCreate = async e => {
    e.preventDefault(); setSaving(true)
    try {
      const authData = await createUser(form.email, form.senha)
      await saveUsuario({ id: authData.id, email: form.email, nome: form.nome, perfil: form.perfil, ativo: true }, token)
      addToast(`Usuário ${form.email} criado com sucesso!`)
      setForm({ email:"", nome:"", perfil:"logistica", senha:"" })
      const updated = await loadUsuarios(token)
      setUsuarios(updated)
    } catch(err) { addToast("Erro: "+err.message, "error") }
    finally { setSaving(false) }
  }

  const handleDelete = async id => {
    if (!confirm("Remover este usuário?")) return
    await deleteUsuario(id, token)
    setUsuarios(u => u.filter(x => x.id!==id))
    addToast("Usuário removido", "warn")
  }

  const handlePerfil = async (id, perfil) => {
    const u = usuarios.find(x => x.id===id)
    await saveUsuario({ ...u, perfil }, token)
    setUsuarios(prev => prev.map(x => x.id===id?{...x,perfil}:x))
    addToast("Perfil atualizado")
  }

  const LABEL = { admin:"Admin", logistica:"Logística", suporte:"Suporte", leitura:"Somente leitura" }

  return (
    <div style={{padding:"20px 24px",maxWidth:800}}>
      <div style={{fontSize:13,fontWeight:500,color:"#000",marginBottom:20,letterSpacing:"0.04em",textTransform:"uppercase"}}>Gestão de Usuários</div>

      {/* Formulário novo usuário */}
      <div style={{background:"#f8f8f8",borderRadius:6,border:"1px solid #e8e8e8",padding:20,marginBottom:24}}>
        <div style={{fontSize:10,color:"#999",textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:14}}>Adicionar usuário</div>
        <form onSubmit={handleCreate}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:12}}>
            <div>
              <div style={{fontSize:10,color:"#999",textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:5}}>Nome</div>
              <input value={form.nome} onChange={e=>setForm(f=>({...f,nome:e.target.value}))} placeholder="Nome completo" required style={{...SEL,width:"100%",boxSizing:"border-box",padding:"8px 10px"}}/>
            </div>
            <div>
              <div style={{fontSize:10,color:"#999",textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:5}}>E-mail</div>
              <input value={form.email} onChange={e=>setForm(f=>({...f,email:e.target.value}))} type="email" placeholder="email@exemplo.com" required style={{...SEL,width:"100%",boxSizing:"border-box",padding:"8px 10px"}}/>
            </div>
            <div>
              <div style={{fontSize:10,color:"#999",textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:5}}>Senha inicial</div>
              <input value={form.senha} onChange={e=>setForm(f=>({...f,senha:e.target.value}))} type="password" placeholder="Mínimo 6 caracteres" required minLength={6} style={{...SEL,width:"100%",boxSizing:"border-box",padding:"8px 10px"}}/>
            </div>
            <div>
              <div style={{fontSize:10,color:"#999",textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:5}}>Perfil de acesso</div>
              <select value={form.perfil} onChange={e=>setForm(f=>({...f,perfil:e.target.value}))} style={{...SEL,width:"100%",boxSizing:"border-box",padding:"8px 10px"}}>
                {Object.entries(LABEL).map(([k,v])=><option key={k} value={k}>{v}</option>)}
              </select>
            </div>
          </div>
          <button type="submit" disabled={saving} style={{background:BRAND,border:"none",color:"#fff",borderRadius:3,padding:"9px 24px",fontSize:12,fontWeight:500,cursor:saving?"not-allowed":"pointer",opacity:saving?0.6:1,letterSpacing:"0.06em",textTransform:"uppercase"}}>
            {saving?"Criando...":"+ Criar usuário"}
          </button>
        </form>
      </div>

      {/* Lista de usuários */}
      {loading ? <div style={{color:"#bbb",fontSize:13}}>Carregando...</div> : (
        <div style={{borderRadius:6,border:"1px solid #e8e8e8",overflow:"hidden"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead>
              <tr style={{background:BRAND}}>
                {["Nome","E-mail","Perfil","Ações"].map(h=><th key={h} style={{padding:"10px 14px",textAlign:"left",color:GOLD,fontWeight:500,fontSize:10,letterSpacing:"0.08em",textTransform:"uppercase"}}>{h}</th>)}
              </tr>
            </thead>
            <tbody>
              {usuarios.length===0?<tr><td colSpan={4} style={{padding:24,textAlign:"center",color:"#bbb"}}>Nenhum usuário cadastrado</td></tr>
              :usuarios.map((u,i)=>(
                <tr key={u.id} style={{background:i%2===0?"#fff":"#fafafa",borderBottom:"1px solid #f2f2f2"}}>
                  <td style={{padding:"10px 14px",color:"#000"}}>{u.nome||"—"}</td>
                  <td style={{padding:"10px 14px",color:"#666"}}>{u.email}</td>
                  <td style={{padding:"10px 14px"}}>
                    <select value={u.perfil} onChange={e=>handlePerfil(u.id,e.target.value)} style={{...SEL,padding:"4px 8px",fontSize:11}}>
                      {Object.entries(LABEL).map(([k,v])=><option key={k} value={k}>{v}</option>)}
                    </select>
                  </td>
                  <td style={{padding:"10px 14px"}}>
                    <button onClick={()=>handleDelete(u.id)} style={{background:"transparent",border:"1px solid #f5c6cb",color:"#c0392b",borderRadius:3,padding:"3px 10px",fontSize:11,cursor:"pointer"}}>Remover</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  )
}

// ─── App Principal ────────────────────────────────────────────
const SAMPLE = `Identificador Ecommerce;Destinatário Nome;Estratégia de Frete;Rastreador Last Mile;Situação;Prazo Logístico;Nº Nota Fiscal
12345;Ana Souza;Correios PAC;AA123456789BR;Em trânsito;05/05/2026;98765
12346;Carlos Lima;Jadlog;JD987654321;Extraviado;28/04/2026;98766
12347;Mariana Costa;Total Express;TE112233445;Entregue;01/05/2026;98767`

export default function App() {
  const [session, setSession]   = useState(null)  // { access_token, user }
  const [perfil, setPerfil]     = useState(null)   // perfil do usuário
  const [loadingPerfil, setLoadingPerfil] = useState(false)

  const [rows, setRows]           = useState([])
  const [tab, setTab]             = useState(null)
  const [paste, setPaste]         = useState("")
  const [importing, setImporting] = useState(false)
  const [compact, setCompact]     = useState(false)
  const [toasts, setToasts]       = useState([])
  const [lSrch, setLSrch] = useState(""); const [lSt, setLSt]   = useState("Todos")
  const [lTr, setLTr]     = useState("Todos"); const [lUrg, setLUrg] = useState("Todos")
  const [lAc, setLAc]     = useState("Todos"); const [qf, setQf]   = useState("todos")
  const [lPage, setLPage] = useState(1)
  const [selIds, setSelIds]       = useState(new Set())
  const [sortCol, setSortCol]     = useState(null); const [sortDir, setSortDir] = useState("asc")
  const [sSrch, setSSrch]   = useState(""); const [sAtend, setSAtend] = useState("Todos")
  const [sUrg, setSUrg]     = useState("Todos")
  const [selSup, setSelSup]       = useState(null)
  const [selSupIds, setSelSupIds] = useState(new Set())
  const [openTpl, setOpenTpl]     = useState(false); const [openHist, setOpenHist] = useState(false)
  const [aSrch, setASrch]         = useState("")
  const [syncStatus, setSyncStatus] = useState("idle")
  const [lastSync, setLastSync]   = useState(null)
  const [countdown, setCountdown] = useState(10)
  const [confirmClear, setConfirmClear] = useState(false)
  const saveTimer = useRef(null)
  const fileRef   = useRef()

  const addToast = useCallback((msg, type="ok", ms=4000) => {
    const id = Date.now()
    setToasts(p => [...p, {id, msg, type}])
    setTimeout(() => setToasts(p => p.filter(t => t.id!==id)), ms)
  }, [])

  const token = session?.access_token

  // Login
  const handleLogin = async data => {
    setSession(data); setLoadingPerfil(true)
    try {
      const r = await fetch(`${SUPA_URL}/rest/v1/usuarios?id=eq.${data.user.id}&select=*`, { headers: authSH(data.access_token) })
      const arr = await r.json()
      const p = arr[0]?.perfil || "leitura"
      setPerfil(p)
      const perms = PERMS[p]
      setTab(perms.tabs[0])
    } catch(e) { setPerfil("leitura"); setTab("dashboard") }
    setLoadingPerfil(false)
  }
  const handleLogout = async () => {
    await signOut(token)
    setSession(null); setPerfil(null); setRows([]); setTab(null)
  }

  const perms = perfil ? PERMS[perfil] : null

  // Load data
  useEffect(() => {
    if (!token) return
    setSyncStatus("loading")
    dbLoad(token).then(data => {
      if (data.length>0) {
        // Corrige automaticamente entregues que ainda não foram arquivados
        // Exceção: pedidos em Suporte ficam no Suporte para finalizar tratativa
        const fixed = data.map(r => ({
          ...r, isNew: false,
          atendimento: isEntregue(r.status) && !r.enviadoSuporte ? "Resolvido" : r.atendimento,
          enviadoSuporte: isEntregue(r.status) && !r.enviadoSuporte ? false : r.enviadoSuporte,
        }))
        setRows(fixed)
        setLastSync(new Date())
      }
      setSyncStatus("idle")
    }).catch(e => { setSyncStatus("error"); addToast("Erro ao carregar: "+e.message,"error",8000) })
  }, [token])

  // Poll
  useEffect(() => {
    if (!token) return
    const poll = async () => {
      setCountdown(10)
      try {
        const remote = await dbLoad(token)
        if (remote.length>0) {
          let nc = 0
          setRows(prev => {
            const rm = new Map(remote.map(r=>[r.id,r]))
            const lm = new Map(prev.map(r=>[r.id,r]))
            const merged = [...rm.values()].map(r => { const loc=lm.get(r.id); if(!loc){nc++;return{...r,isNew:true}} return loc.historico.length>=r.historico.length?loc:{...r,isNew:false} })
            prev.forEach(r => { if(!rm.has(r.id)) merged.push(r) })
            return merged
          })
          if (nc>0) addToast(`${nc} pedido${nc>1?"s":""} atualizado${nc>1?"s":""} por outro usuário`,"warn")
          setLastSync(new Date())
        }
      } catch(e) {}
    }
    const interval   = setInterval(poll, 10000)
    const cdInterval = setInterval(() => setCountdown(p => p>0?p-1:10), 1000)
    return () => { clearInterval(interval); clearInterval(cdInterval) }
  }, [token, addToast])

  // Save
  useEffect(() => {
    if (!token||rows.length===0) return
    if (saveTimer.current) clearTimeout(saveTimer.current)
    saveTimer.current = setTimeout(async () => {
      setSyncStatus("saving")
      try {
        await dbUpsert(rows, token)
        setLastSync(new Date()); setSyncStatus("saved")
        setTimeout(() => setSyncStatus("idle"), 2500)
      } catch(e) {
        setSyncStatus("error"); addToast("Erro ao salvar: "+e.message,"error",8000)
        setTimeout(() => setSyncStatus("idle"), 4000)
      }
    }, 1200)
  }, [rows, token, addToast])

  useEffect(() => { if(!rows.some(r=>r.isNew)) return; const t=setTimeout(()=>setRows(p=>p.map(r=>({...r,isNew:false}))),6000); return()=>clearTimeout(t) }, [rows])
  useEffect(() => setLPage(1), [lSrch,lSt,lTr,lUrg,lAc,qf,sortCol,sortDir])
  useEffect(() => { setOpenTpl(false); setOpenHist(false) }, [selSup])

  const doImport = useCallback(txt => {
    if (!perms?.canImport) return
    const parsed = parseData(txt)
    if (!parsed.length) { addToast("Nenhum dado reconhecido.","error"); return }
    let added=0, updated=0, skipped=0
    setRows(prev => {
      const byNuvem = new Map(prev.map(r=>[r.nuvem, r]))
      const result = [...prev]
      for (const novo of parsed) {
        const existing = byNuvem.get(novo.nuvem)
        if (!existing) {
          // Novo pedido — adiciona
          result.push(novo); added++
        } else if (norm(existing.status) === norm(novo.status)) {
          // Mesmo status — se for entregue, não está em suporte e não arquivado, arquiva
          if (isEntregue(novo.status) && !existing.enviadoSuporte && existing.atendimento !== "Resolvido") {
            const idx = result.findIndex(r=>r.nuvem===novo.nuvem)
            if (idx>=0) {
              result[idx] = { ...existing, atendimento:"Resolvido", enviadoSuporte:false,
                historico:[...existing.historico,{acao:"Arquivado automaticamente — entrega concluída",ts:new Date().toLocaleString("pt-BR")}] }
              updated++
            }
          } else { skipped++ }
        } else {
          // Status diferente — atualiza mantendo obs/historico/responsavel
          const idx = result.findIndex(r=>r.nuvem===novo.nuvem)
          if (idx>=0) {
            const statusAnterior = existing.status
            // Se está em Suporte e status mudou → marca alerta para o atendente
            const alertaStatus = existing.enviadoSuporte && norm(statusAnterior) !== norm(novo.status)
            result[idx] = {
              ...novo,
              id: existing.id,
              obs: existing.obs,
              responsavel: existing.responsavel,
              chamado: existing.chamado,
              enviadoSuporte: existing.enviadoSuporte,
              atendimento: existing.enviadoSuporte ? existing.atendimento : novo.atendimento,
              alertaStatus: alertaStatus ? `Status atualizado: ${statusAnterior} → ${novo.status}` : existing.alertaStatus,
              historico: [...existing.historico, {acao:`Status atualizado: ${statusAnterior} → ${novo.status}`, ts:new Date().toLocaleString("pt-BR")}],
              isNew: true
            }
            updated++
          }
        }
      }
      return result
    })
    setTimeout(() => {
      const parts = []
      if (added>0)   parts.push(`${added} novo${added>1?"s":""}`)
      if (updated>0) parts.push(`${updated} atualizado${updated>1?"s":""}`)
      if (skipped>0) parts.push(`${skipped} ignorado${skipped>1?"s":""} (mesmo status)`)
      addToast(parts.join(" · ") || "Nenhuma alteração")
    }, 100)
    setPaste(""); setImporting(false)
  }, [addToast, perms])

  const handleFile = e => {
    if (!perms?.canImport) return
    const f = e.target.files[0]; if(!f) return
    if (/\.(xlsx?)$/i.test(f.name)) {
      const rd=new FileReader(); rd.onload=ev=>{const wb=XLSX.read(new Uint8Array(ev.target.result),{type:"array"});doImport(XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]],{FS:";",blankrows:false}))}; rd.readAsArrayBuffer(f)
    } else {
      const rd=new FileReader(); rd.onload=ev=>doImport(ev.target.result); rd.readAsText(f,/\.csv$/i.test(f.name)?"windows-1252":"UTF-8")
    }
    e.target.value=""
  }

  const upd = (id,ch,hist) => setRows(prev=>prev.map(r=>{ if(r.id!==id) return r; const historico=hist?[...r.historico,{...hist,ts:new Date().toLocaleString("pt-BR")}]:r.historico; return{...r,...ch,historico} }))
  const del = id => { if(!perms?.canDelete) return; setRows(prev=>prev.filter(r=>r.id!==id)); dbDelete(id,token).catch(()=>{}) }
  const toggleSel = id => setSelIds(prev=>{const n=new Set(prev);n.has(id)?n.delete(id):n.add(id);return n})
  const clearSel = () => setSelIds(new Set())
  const bulkSend = () => {
    if (!perms?.canSendSupport) return
    const ts=new Date().toLocaleString("pt-BR"),sentAt=new Date().toISOString()
    setRows(prev=>prev.map(r=>selIds.has(r.id)?{...r,enviadoSuporte:true,atendimento:"Aberto",sentAt,historico:[...r.historico,{acao:"Enviado ao suporte (lote)",ts}]}:r))
    addToast(`${selIds.size} pedido${selIds.size>1?"s":""} enviado${selIds.size>1?"s":""} ao suporte`)
    clearSel()
  }
  const toggleSelSup = id => setSelSupIds(prev=>{const n=new Set(prev);n.has(id)?n.delete(id):n.add(id);return n})
  const bulkArchive = () => {
    if (!perms?.canOperate) return
    const ts=new Date().toLocaleString("pt-BR")
    setRows(prev=>prev.map(r=>selSupIds.has(r.id)?{...r,atendimento:"Resolvido",historico:[...r.historico,{acao:"Arquivado em lote",ts}]}:r))
    addToast(`${selSupIds.size} pedido${selSupIds.size>1?"s":""} arquivado${selSupIds.size>1?"s":""}`)
    if(selSupIds.has(selSup)) setSelSup(null)
    setSelSupIds(new Set())
  }
  const handleInitiate = id => { if(!perms?.canOperate) return; upd(id,{atendimento:"Em andamento"},{acao:"Atendimento iniciado"}); setSelSup(id) }
  const handleResolve  = id => { if(!perms?.canOperate) return; upd(id,{atendimento:"Resolvido"},{acao:"Status → Resolvido"}); setSelSup(null); addToast("Pedido resolvido e arquivado") }
  const handleReturnLog = id => { if(!perms?.canOperate) return; upd(id,{enviadoSuporte:false,sentAt:null},{acao:"Devolvido à Logística"}); setSelSup(null) }
  const handleClearAll = () => {
    if (!perms?.canClear) return
    if (!confirmClear) { setConfirmClear(true); setTimeout(()=>setConfirmClear(false),4000); return }
    setRows([]); setConfirmClear(false); dbClear(token).catch(()=>{})
    addToast("Todos os dados foram removidos","warn")
  }
  const toggleSort = col => { if(sortCol===col) setSortDir(d=>d==="asc"?"desc":"asc"); else{setSortCol(col);setSortDir("asc")} }

  if (!session) return <LoginScreen onLogin={handleLogin}/>
  if (loadingPerfil) return <div style={{display:"flex",alignItems:"center",justifyContent:"center",minHeight:"100vh",color:"#999",fontSize:13}}>Carregando perfil...</div>

  const baseLog  = rows.filter(r=>!r.enviadoSuporte&&r.atendimento!=="Resolvido")
  const baseSup  = rows.filter(r=>r.enviadoSuporte&&r.atendimento!=="Resolvido")
  const baseArch = rows.filter(r=>r.atendimento==="Resolvido")
  const detail   = selSup?baseSup.find(r=>r.id===selSup):null

  const qCounts = Object.fromEntries(QFILTERS.map(f=>[f.id,applyQF(baseLog,f.id).length]))
  const filteredLog = applySortRows(
    applyQF(baseLog,qf).filter(r => {
      const q=lSrch.toLowerCase()
      return (!q||[r.nuvem,r.destinatario,r.transportadora,r.rastreio,r.status,r.nf,r.motivo].some(v=>v.toLowerCase().includes(q)))
        &&(lSt==="Todos"||r.status===lSt)&&(lTr==="Todos"||r.transportadora===lTr)
        &&(lUrg==="Todos"||r.urgencia===lUrg)&&(lAc==="Todos"||r.acionar===lAc)
    }), sortCol, sortDir
  )
  const totalPages = Math.max(1,Math.ceil(filteredLog.length/PAGE_SIZE))
  const safeP = Math.min(lPage,totalPages)
  const pagedLog = filteredLog.slice((safeP-1)*PAGE_SIZE,safeP*PAGE_SIZE)
  const supRows = baseSup.filter(r=>{const q=sSrch.toLowerCase();return(!q||[r.nuvem,r.destinatario,r.rastreio,r.nf,r.status].some(v=>v.toLowerCase().includes(q)))&&(sAtend==="Todos"||r.atendimento===sAtend)&&(sUrg==="Todos"||r.urgencia===sUrg)}).sort((a,b)=>{const uo={Alta:0,Média:1,Baixa:2,"—":3},ao={Aberto:0,"Em andamento":1};return(uo[a.urgencia]-uo[b.urgencia])||(ao[a.atendimento]-ao[b.atendimento])})
  const archRows = baseArch.filter(r=>{const q=aSrch.toLowerCase();return!q||[r.nuvem,r.destinatario,r.transportadora,r.status,r.nf].some(v=>v.toLowerCase().includes(q))}).sort((a,b)=>{const ta=(a.historico.find(h=>h.acao&&(h.acao.includes("Resolvido")||h.acao.includes("Arquivado")))||{}).ts||"";const tb=(b.historico.find(h=>h.acao&&(h.acao.includes("Resolvido")||h.acao.includes("Arquivado")))||{}).ts||"";return tb.localeCompare(ta)})

  const stOpts=uniq(baseLog.map(r=>r.status)), trOpts=uniq(baseLog.map(r=>r.transportadora))
  const st={log:baseLog.length,alta:baseLog.filter(r=>r.urgencia==="Alta").length,acionar:baseLog.filter(r=>r.acionar==="Sim").length}
  const ss={total:baseSup.length,abertos:baseSup.filter(r=>r.atendimento==="Aberto").length,andamento:baseSup.filter(r=>r.atendimento==="Em andamento").length}
  const arch=baseArch.length

  const urgData=["Alta","Média","Baixa","—"].map(u=>({name:u,value:rows.filter(r=>r.urgencia===u).length})).filter(d=>d.value>0)
  const statusMap={};rows.forEach(r=>{if(r.status)statusMap[r.status]=(statusMap[r.status]||0)+1})
  const statusData=Object.entries(statusMap).sort((a,b)=>b[1]-a[1]).slice(0,8).map(([name,value])=>({name,value}))
  const carrierMap={};rows.forEach(r=>{if(r.transportadora)carrierMap[r.transportadora]=(carrierMap[r.transportadora]||0)+1})
  const carrierData=Object.entries(carrierMap).sort((a,b)=>b[1]-a[1]).slice(0,8).map(([name,value])=>({name,value}))
  const hoje=new Date();hoje.setHours(0,0,0,0)
  const slaData=[{name:"No prazo",value:rows.filter(r=>{const d=parsePrazo(r.prazo);return d&&d>=hoje}).length,fill:"#27ae60"},{name:"Entregues",value:rows.filter(r=>(r.status||"").toLowerCase().includes("entregue")).length,fill:GOLD},{name:"Vencidos",value:rows.filter(r=>{const d=parsePrazo(r.prazo);return d&&d<hoje&&!(r.status||"").toLowerCase().includes("entregue")}).length,fill:"#e74c3c"}].filter(d=>d.value>0)

  const showImp = (importing||rows.length===0) && perms?.tabs.some(t=>["logistica","dashboard"].includes(t))
  const pd=compact?4:7, ct="#999"
  const thS={padding:"8px 10px",textAlign:"left",color:"#fff",fontWeight:500,fontSize:10,letterSpacing:"0.08em",textTransform:"uppercase",borderBottom:"1px solid #222",whiteSpace:"nowrap",background:BRAND,position:"sticky",top:0,zIndex:5,cursor:"pointer"}
  const thF={...thS,cursor:"default"}

  const TABS_CONFIG = [
    {key:"dashboard", label:"Dashboard", badge:null},
    {key:"logistica", label:"Logística", badge:st.acionar>0?st.acionar:null},
    {key:"suporte",   label:"Suporte",   badge:ss.abertos>0?ss.abertos:null},
    {key:"arquivados",label:"Arquivados",badge:arch>0?arch:null},
    {key:"usuarios",  label:"Usuários",  badge:null},
  ].filter(t => perms?.tabs.includes(t.key))

  const PERFLABEL = { admin:"Admin", logistica:"Logística", suporte:"Suporte", leitura:"Leitura" }
  const syncColor = syncStatus==="error"?"#e74c3c":syncStatus==="saving"?GOLD:syncStatus==="saved"?"#27ae60":syncStatus==="loading"?"#888":"#555"

  return (
    <div style={{fontFamily:"'Helvetica Neue',Arial,sans-serif",minHeight:"100vh",background:"#fff"}}>
      <Toast toasts={toasts}/>

      {/* HEADER */}
      <div style={{background:BRAND,padding:"13px 24px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <div>
          <div style={{color:"#fff",fontSize:15,fontWeight:400,letterSpacing:"0.38em",textTransform:"uppercase"}}>Saint Germain</div>
          <div style={{color:"#555",fontSize:9,letterSpacing:"0.28em",textTransform:"uppercase",marginTop:2}}>Central de Pedidos</div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          <div style={{display:"flex",alignItems:"center",gap:6,fontSize:10,color:syncColor}}>
            <span style={{width:7,height:7,borderRadius:"50%",background:syncColor,display:"inline-block"}}/>
            {syncStatus==="loading"&&"Carregando..."}
            {syncStatus==="saving"&&"Salvando..."}
            {syncStatus==="saved"&&"Sincronizado ✓"}
            {syncStatus==="error"&&"Erro"}
            {syncStatus==="idle"&&lastSync&&`Sync em ${countdown}s`}
          </div>
          {perms?.canImport && !showImp && rows.length>0 && <button onClick={()=>setCompact(c=>!c)} style={{background:"transparent",border:"1px solid #333",color:compact?GOLD:"#555",borderRadius:2,padding:"4px 10px",fontSize:10,cursor:"pointer"}}>{compact?"⊞ Normal":"⊟ Compacto"}</button>}
          {perms?.canImport && !showImp && rows.length>0 && <button onClick={()=>{const hf=lSrch||lSt!=="Todos"||lTr!=="Todos"||lUrg!=="Todos"||lAc!=="Todos";exportCSV(hf&&tab==="logistica"?filteredLog:rows)}} style={{background:"transparent",border:"1px solid #444",color:"#aaa",borderRadius:2,padding:"5px 12px",fontSize:11,cursor:"pointer"}}>↓ Exportar</button>}
          {perms?.canImport && !showImp && <button onClick={()=>setImporting(true)} style={{background:"transparent",border:`1px solid ${GOLD}`,color:GOLD,borderRadius:2,padding:"5px 12px",fontSize:11,cursor:"pointer"}}>+ Importar</button>}
          {perms?.canClear && rows.length>0 && <button onClick={handleClearAll} style={{background:confirmClear?"#c0392b":"transparent",border:`1px solid ${confirmClear?"#c0392b":"#444"}`,color:confirmClear?"#fff":"#666",borderRadius:2,padding:"5px 12px",fontSize:11,cursor:"pointer"}}>{confirmClear?"⚠ Confirmar":"Limpar"}</button>}
          <div style={{display:"flex",alignItems:"center",gap:8,borderLeft:"1px solid #333",paddingLeft:12}}>
            <div style={{textAlign:"right"}}>
              <div style={{color:"#fff",fontSize:11}}>{session.user?.email}</div>
              <div style={{color:GOLD,fontSize:10,letterSpacing:"0.05em"}}>{PERFLABEL[perfil]||perfil}</div>
            </div>
            <button onClick={handleLogout} style={{background:"transparent",border:"1px solid #444",color:"#aaa",borderRadius:2,padding:"5px 10px",fontSize:11,cursor:"pointer"}}>Sair</button>
          </div>
        </div>
        <input ref={fileRef} type="file" accept=".csv,.txt,.tsv,.xls,.xlsx" style={{display:"none"}} onChange={handleFile}/>
      </div>

      {/* TABS */}
      <div style={{background:"#fff",borderBottom:"1px solid #e8e8e8",padding:"0 20px",display:"flex",alignItems:"stretch"}}>
        {TABS_CONFIG.map(t=>(
          <button key={t.key} onClick={()=>{setTab(t.key);if(t.key!=="suporte")setSelSup(null)}} style={{background:"transparent",border:"none",borderBottom:tab===t.key?"2px solid #000":"2px solid transparent",color:tab===t.key?"#000":"#999",padding:"11px 16px",cursor:"pointer",fontSize:11,letterSpacing:"0.1em",textTransform:"uppercase",fontWeight:tab===t.key?500:400,marginBottom:"-1px",display:"flex",alignItems:"center",gap:6}}>
            {t.label}
            {t.badge!=null&&<span style={{background:tab===t.key?"#000":"#e74c3c",color:"#fff",borderRadius:10,padding:"1px 6px",fontSize:10,fontWeight:500}}>{t.badge}</span>}
          </button>
        ))}
      </div>

      {/* IMPORT */}
      {showImp && perms?.canImport && (
        <div style={{padding:32,maxWidth:620,margin:"0 auto"}}>
          <div style={{textAlign:"center",marginBottom:28}}>
            <div style={{color:GOLD,fontSize:11,letterSpacing:"0.22em",textTransform:"uppercase",marginBottom:10}}>{importing?"Adicionar pedidos":"Central de Pedidos"}</div>
            <div style={{color:"#000",fontSize:22,fontWeight:400,marginBottom:8}}>Importe seus dados</div>
            <div style={{color:"#888",fontSize:13}}>Aceita .csv (cp1252), .xls, .xlsx ou colagem direta do Excel.</div>
          </div>
          <textarea value={paste} onChange={e=>setPaste(e.target.value)} placeholder="Cole aqui os dados copiados do sistema..."
            style={{width:"100%",minHeight:130,borderRadius:3,border:"1px solid #d5d5d5",padding:12,fontSize:13,resize:"vertical",fontFamily:"monospace",boxSizing:"border-box"}}/>
          <div style={{display:"flex",gap:10,marginTop:10}}>
            <button onClick={()=>doImport(paste)} disabled={!paste.trim()} style={{flex:1,background:BRAND,border:"none",color:"#fff",borderRadius:3,padding:"11px 0",fontSize:12,fontWeight:500,cursor:paste.trim()?"pointer":"not-allowed",opacity:paste.trim()?1:0.4,letterSpacing:"0.08em",textTransform:"uppercase"}}>Importar colados</button>
            <button onClick={()=>fileRef.current.click()} style={{flex:1,background:"transparent",border:`1px solid ${GOLD}`,color:GOLD,borderRadius:3,padding:"11px 0",fontSize:12,cursor:"pointer"}}>Importar arquivo</button>
          </div>
          <div style={{textAlign:"center",marginTop:14,display:"flex",justifyContent:"center",gap:20}}>
            <button onClick={()=>doImport(SAMPLE)} style={{background:"transparent",border:"none",color:"#bbb",fontSize:12,cursor:"pointer",textDecoration:"underline"}}>Carregar exemplo</button>
            {importing&&<button onClick={()=>setImporting(false)} style={{background:"transparent",border:"none",color:"#bbb",fontSize:12,cursor:"pointer"}}>Cancelar</button>}
          </div>
        </div>
      )}

      {/* USUÁRIOS */}
      {tab==="usuarios" && perfil==="admin" && <UsuariosPanel token={token} addToast={addToast}/>}

      {/* DASHBOARD */}
      {tab==="dashboard"&&!showImp&&(
        <div style={{padding:"20px 24px"}}>
          {/* KPIs principais */}
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:16}}>
            <StatCard label="Em logística"   val={st.log}   sub={`${st.acionar} precisam de suporte`}/>
            <StatCard label="No suporte"     val={ss.total} accent={ss.abertos>0} sub={`${ss.abertos} abertos`}/>
            <StatCard label="Urgência alta"  val={st.alta}  accent={st.alta>0} sub="na logística"/>
            <StatCard label="Arquivados"     val={arch}     sub="atendimentos concluídos"/>
          </div>

          {(() => {
            const entregues     = rows.filter(r=>isEntregue(r.status))
            const total         = rows.length
            const noPrazo       = entregues.filter(r=>r.entregueNoPrazo===true).length
            const foraPrazo     = entregues.filter(r=>r.entregueNoPrazo===false).length
            const semInfo       = entregues.filter(r=>r.entregueNoPrazo===null).length
            const pctNoPrazo    = entregues.length>0 ? Math.round((noPrazo/entregues.length)*100) : 0
            const vencidos      = rows.filter(r=>{const d=parsePrazo(r.prazo);if(!d)return false;const h=new Date();h.setHours(0,0,0,0);return d<h&&!isEntregue(r.status)}).length
            const emRisco       = rows.filter(r=>{const d=parsePrazo(r.prazo);if(!d)return false;const h=new Date();h.setHours(0,0,0,0);const diff=Math.ceil((d-h)/86400000);return diff>=0&&diff<=3&&!isEntregue(r.status)}).length

            // Por transportadora
            const trStats = {}
            rows.forEach(r=>{
              if(!r.transportadora) return
              if(!trStats[r.transportadora]) trStats[r.transportadora]={total:0,entregues:0,noPrazo:0,foraPrazo:0,vencidos:0,emRisco:0}
              const s = trStats[r.transportadora]
              s.total++
              if(isEntregue(r.status)){
                s.entregues++
                if(r.entregueNoPrazo===true)  s.noPrazo++
                if(r.entregueNoPrazo===false) s.foraPrazo++
              } else {
                const d=parsePrazo(r.prazo)
                if(d){
                  const h=new Date();h.setHours(0,0,0,0)
                  const diff=Math.ceil((d-h)/86400000)
                  if(diff<0) s.vencidos++
                  else if(diff<=3) s.emRisco++
                }
              }
            })
            const trData = Object.entries(trStats)
              .map(([name,s])=>({name, total:s.total, entregues:s.entregues, noPrazo:s.noPrazo, foraPrazo:s.foraPrazo, vencidos:s.vencidos, pct:s.entregues>0?Math.round((s.noPrazo/s.entregues)*100):0}))
              .sort((a,b)=>b.total-a.total).slice(0,8)

            const trBarData = trData.map(t=>({name:t.name, "No prazo":t.noPrazo, "Fora do prazo":t.foraPrazo, "Vencidos":t.vencidos}))

            const statusDistData = [
              {name:"Entregues no prazo",  value:noPrazo,   fill:"#27ae60"},
              {name:"Entregues fora prazo",value:foraPrazo, fill:"#e74c3c"},
              {name:"Em trânsito",         value:rows.filter(r=>(r.status||"").toLowerCase().includes("transito")||r.status?.toLowerCase().includes("trânsito")).length, fill:"#3498db"},
              {name:"Vencidos",            value:vencidos,  fill:"#c0392b"},
              {name:"Em risco (≤3d)",      value:emRisco,   fill:GOLD},
              {name:"Outros",              value:Math.max(0,total-noPrazo-foraPrazo-vencidos-emRisco-rows.filter(r=>(r.status||"").toLowerCase().includes("transito")||r.status?.toLowerCase().includes("trânsito")).length), fill:"#bbb"},
            ].filter(d=>d.value>0)

            return (
              <div>
                {/* KPIs de desempenho */}
                <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:10,marginBottom:16}}>
                  {[
                    {label:"Total de pedidos",    val:total,      sub:"na base",             accent:false},
                    {label:"Entregues",           val:entregues.length, sub:`${Math.round((entregues.length/Math.max(total,1))*100)}% do total`, accent:false},
                    {label:"Entregues no prazo",  val:`${pctNoPrazo}%`, sub:`${noPrazo} pedidos`,  accent:false},
                    {label:"Vencidos em aberto",  val:vencidos,   sub:"prazo expirado",       accent:vencidos>0},
                    {label:"Em risco (≤3 dias)",  val:emRisco,    sub:"vencem em breve",      accent:emRisco>0},
                  ].map(k=><div key={k.label} style={{background:"#fff",borderRadius:6,padding:"12px 14px",border:k.accent?"1px solid #f5c6cb":"1px solid #e8e8e8"}}>
                    <div style={{fontSize:10,color:"#999",marginBottom:4,letterSpacing:"0.05em",textTransform:"uppercase"}}>{k.label}</div>
                    <div style={{fontSize:24,fontWeight:500,color:k.accent?"#c0392b":GOLD}}>{k.val}</div>
                    <div style={{fontSize:11,color:"#aaa",marginTop:2}}>{k.sub}</div>
                  </div>)}
                </div>

                {/* Gráficos linha 1 */}
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
                  <div style={{background:"#fff",borderRadius:6,border:"1px solid #e8e8e8",padding:16}}>
                    <div style={{fontSize:10,fontWeight:500,color:"#999",marginBottom:4,letterSpacing:"0.08em",textTransform:"uppercase"}}>Distribuição geral de pedidos</div>
                    <ResponsiveContainer width="100%" height={200}>
                      <PieChart><Pie data={statusDistData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={80} innerRadius={45} paddingAngle={2}>{statusDistData.map((e,i)=><Cell key={i} fill={e.fill}/>)}</Pie><Tooltip contentStyle={{fontSize:12}}/></PieChart>
                    </ResponsiveContainer>
                    <div style={{display:"flex",flexWrap:"wrap",gap:"4px 12px",justifyContent:"center"}}>
                      {statusDistData.map(e=><span key={e.name} style={{fontSize:10,color:ct,display:"flex",alignItems:"center",gap:4}}><span style={{width:7,height:7,borderRadius:"50%",background:e.fill,display:"inline-block"}}/>{e.name} ({e.value})</span>)}
                    </div>
                  </div>
                  <div style={{background:"#fff",borderRadius:6,border:"1px solid #e8e8e8",padding:16}}>
                    <div style={{fontSize:10,fontWeight:500,color:"#999",marginBottom:4,letterSpacing:"0.08em",textTransform:"uppercase"}}>Desempenho por transportadora</div>
                    <ResponsiveContainer width="100%" height={220}>
                      <BarChart data={trBarData} layout="vertical" margin={{left:0,right:10,top:4,bottom:0}}>
                        <XAxis type="number" tick={{fontSize:10,fill:ct}} axisLine={false} tickLine={false}/>
                        <YAxis type="category" dataKey="name" width={100} tick={{fontSize:10,fill:ct}} axisLine={false} tickLine={false}/>
                        <Tooltip contentStyle={{fontSize:12}}/>
                        <Bar dataKey="No prazo"      stackId="a" fill="#27ae60" radius={[0,0,0,0]} name="No prazo"/>
                        <Bar dataKey="Fora do prazo" stackId="a" fill="#e74c3c" radius={[0,0,0,0]} name="Fora do prazo"/>
                        <Bar dataKey="Vencidos"      stackId="a" fill="#c0392b" radius={[0,3,3,0]} name="Vencidos"/>
                      </BarChart>
                    </ResponsiveContainer>
                    <div style={{display:"flex",gap:12,justifyContent:"center"}}>
                      {[["No prazo","#27ae60"],["Fora do prazo","#e74c3c"],["Vencidos","#c0392b"]].map(([l,c])=><span key={l} style={{fontSize:10,color:ct,display:"flex",alignItems:"center",gap:4}}><span style={{width:7,height:7,borderRadius:1,background:c,display:"inline-block"}}/>{l}</span>)}
                    </div>
                  </div>
                </div>

                {/* Ranking transportadoras */}
                <div style={{background:"#fff",borderRadius:6,border:"1px solid #e8e8e8",padding:16,marginBottom:16}}>
                  <div style={{fontSize:10,fontWeight:500,color:"#999",marginBottom:14,letterSpacing:"0.08em",textTransform:"uppercase"}}>Ranking de transportadoras — taxa de entrega no prazo</div>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                    <thead>
                      <tr style={{borderBottom:"1px solid #f0f0f0"}}>
                        {["Transportadora","Total","Entregues","No prazo","Fora prazo","Vencidos","% No prazo"].map(h=><th key={h} style={{padding:"6px 10px",textAlign:"left",fontSize:10,color:"#999",fontWeight:500,letterSpacing:"0.06em",textTransform:"uppercase"}}>{h}</th>)}
                      </tr>
                    </thead>
                    <tbody>
                      {trData.map((t,i)=>(
                        <tr key={t.name} style={{borderBottom:"1px solid #fafafa",background:i%2===0?"#fff":"#fafafa"}}>
                          <td style={{padding:"8px 10px",fontWeight:500,color:"#000"}}>{t.name}</td>
                          <td style={{padding:"8px 10px",color:"#666"}}>{t.total}</td>
                          <td style={{padding:"8px 10px",color:"#666"}}>{t.entregues}</td>
                          <td style={{padding:"8px 10px",color:"#27ae60",fontWeight:500}}>{t.noPrazo}</td>
                          <td style={{padding:"8px 10px",color:"#e74c3c",fontWeight:500}}>{t.foraPrazo}</td>
                          <td style={{padding:"8px 10px",color:"#c0392b"}}>{t.vencidos}</td>
                          <td style={{padding:"8px 10px"}}>
                            <div style={{display:"flex",alignItems:"center",gap:8}}>
                              <div style={{flex:1,height:6,background:"#f0f0f0",borderRadius:3,overflow:"hidden"}}>
                                <div style={{height:"100%",width:`${t.pct}%`,background:t.pct>=80?"#27ae60":t.pct>=60?GOLD:"#e74c3c",borderRadius:3}}/>
                              </div>
                              <span style={{fontSize:11,fontWeight:500,color:t.pct>=80?"#27ae60":t.pct>=60?"#856404":"#e74c3c",minWidth:32}}>{t.pct}%</span>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {/* Urgência + Status dos pedidos em aberto */}
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
                  <div style={{background:"#fff",borderRadius:6,border:"1px solid #e8e8e8",padding:16}}>
                    <div style={{fontSize:10,fontWeight:500,color:"#999",marginBottom:4,letterSpacing:"0.08em",textTransform:"uppercase"}}>Urgência — pedidos em aberto</div>
                    <ResponsiveContainer width="100%" height={190}><PieChart><Pie data={urgData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={74} innerRadius={40} paddingAngle={3}>{urgData.map((e,i)=><Cell key={i} fill={PIE_COLORS[e.name]||GOLD}/>)}</Pie><Tooltip contentStyle={{fontSize:12}}/></PieChart></ResponsiveContainer>
                    <div style={{display:"flex",justifyContent:"center",gap:10,flexWrap:"wrap"}}>{urgData.map(e=><span key={e.name} style={{fontSize:10,color:ct,display:"flex",alignItems:"center",gap:4}}><span style={{width:7,height:7,borderRadius:"50%",background:PIE_COLORS[e.name]||GOLD,display:"inline-block"}}/>{e.name} ({e.value})</span>)}</div>
                  </div>
                  <div style={{background:"#fff",borderRadius:6,border:"1px solid #e8e8e8",padding:16}}>
                    <div style={{fontSize:10,fontWeight:500,color:"#999",marginBottom:4,letterSpacing:"0.08em",textTransform:"uppercase"}}>Status dos pedidos em aberto</div>
                    <ResponsiveContainer width="100%" height={220}><BarChart data={statusData} layout="vertical" margin={{left:0,right:20,top:4,bottom:0}}><XAxis type="number" tick={{fontSize:10,fill:ct}} axisLine={false} tickLine={false}/><YAxis type="category" dataKey="name" width={120} tick={{fontSize:10,fill:ct}} axisLine={false} tickLine={false}/><Tooltip contentStyle={{fontSize:12}}/><Bar dataKey="value" fill={GOLD} radius={[0,3,3,0]} name="Pedidos"/></BarChart></ResponsiveContainer>
                  </div>
                </div>
              </div>
            )
          })()}
        </div>
      )}

      {/* LOGÍSTICA */}
      {tab==="logistica"&&!showImp&&(
        <div style={{padding:"16px 24px"}}>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12,marginBottom:14}}>
            <StatCard label="Em logística" val={st.log}/><StatCard label="Urgência alta" val={st.alta} accent={st.alta>0}/><StatCard label="Acionar suporte" val={st.acionar} accent={st.acionar>0}/>
          </div>
          <div style={{display:"flex",gap:6,marginBottom:10,flexWrap:"wrap"}}>
            {QFILTERS.map(f=><button key={f.id} onClick={()=>{setQf(f.id);clearSel()}} style={{background:qf===f.id?"#000":"transparent",border:`1px solid ${qf===f.id?"#000":"#d5d5d5"}`,color:qf===f.id?"#fff":"#777",borderRadius:2,padding:"4px 12px",fontSize:11,cursor:"pointer",fontWeight:qf===f.id?500:400}}>{f.label}{f.id!=="todos"?` (${qCounts[f.id]||0})`:""}</button>)}
          </div>
          <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:10,alignItems:"center"}}>
            <input value={lSrch} onChange={e=>setLSrch(e.target.value)} placeholder="Buscar pedido, destinatário, rastreio..." style={{...SEL,flex:1,minWidth:160}}/>
            <select value={lSt} onChange={e=>setLSt(e.target.value)} style={SEL}>{stOpts.map(o=><option key={o}>{o}</option>)}</select>
            <select value={lTr} onChange={e=>setLTr(e.target.value)} style={SEL}>{trOpts.map(o=><option key={o}>{o}</option>)}</select>
            <select value={lUrg} onChange={e=>setLUrg(e.target.value)} style={SEL}>{["Todos","Alta","Média","Baixa","—"].map(o=><option key={o}>{o}</option>)}</select>
            <select value={lAc} onChange={e=>setLAc(e.target.value)} style={SEL}>{["Todos","Sim","Avaliar","Não"].map(o=><option key={o}>{o}</option>)}</select>
            {(lSrch||lSt!=="Todos"||lTr!=="Todos"||lUrg!=="Todos"||lAc!=="Todos")&&<button onClick={()=>{setLSrch("");setLSt("Todos");setLTr("Todos");setLUrg("Todos");setLAc("Todos")}} style={{...SEL,cursor:"pointer"}}>Limpar</button>}
          </div>
          {perms?.canSendSupport && selIds.size>0&&(
            <div style={{background:"#000",borderRadius:3,padding:"10px 16px",marginBottom:10,display:"flex",alignItems:"center",gap:12}}>
              <span style={{color:"#aaa",fontSize:13,flex:1}}>{selIds.size} selecionado{selIds.size>1?"s":""}</span>
              <button onClick={bulkSend} style={{background:GOLD,border:"none",color:"#111",borderRadius:2,padding:"6px 16px",fontSize:12,cursor:"pointer",fontWeight:500}}>Enviar ao Suporte ({selIds.size})</button>
              <button onClick={clearSel} style={{background:"transparent",border:"1px solid #555",color:"#aaa",borderRadius:2,padding:"6px 12px",fontSize:12,cursor:"pointer"}}>Cancelar</button>
            </div>
          )}
          <div style={{overflowX:"auto",overflowY:"auto",maxHeight:"55vh",borderRadius:4,border:"1px solid #e0e0e0"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:compact?11:12,tableLayout:"fixed",minWidth:1020}}>
              <colgroup><col style={{width:34}}/><col style={{width:76}}/><col style={{width:118}}/><col style={{width:95}}/><col style={{width:108}}/><col style={{width:96}}/><col style={{width:104}}/><col style={{width:66}}/><col style={{width:114}}/><col style={{width:64}}/><col style={{width:74}}/><col style={{width:97}}/><col style={{width:34}}/></colgroup>
              <thead>
                <tr>
                  <th style={thF}>{perms?.canSendSupport&&<input type="checkbox" onChange={e=>e.target.checked?setSelIds(new Set(pagedLog.map(r=>r.id))):clearSel()} checked={selIds.size>0&&pagedLog.every(r=>selIds.has(r.id))} style={{cursor:"pointer"}}/>}</th>
                  {[["nuvem","No NUVEM"],["destinatario","Destinatário"],["transportadora","Transportadora"],["rastreio","Cód. Rastreio"],["status","Status"],["prazo","Prazo / SLA"],["nf","No NF"],["motivo","Motivo (auto)"],["urgencia","Urgência"],["acionar","Acionar?"]].map(([col,label])=>(
                    <th key={col} onClick={()=>toggleSort(col)} style={thS}>{label} <SortIcon col={col} sortCol={sortCol} sortDir={sortDir}/></th>
                  ))}
                  <th style={thF}>Ação</th><th style={thF}></th>
                </tr>
              </thead>
              <tbody>
                {pagedLog.length===0?<tr><td colSpan={13} style={{textAlign:"center",padding:28,color:"#bbb"}}>Nenhum pedido encontrado</td></tr>
                :pagedLog.map((r,i)=>(
                  <tr key={r.id} style={{background:r.isNew?"rgba(196,150,58,0.12)":rowBg(r.urgencia,compact,i),borderBottom:"1px solid #f2f2f2"}}>
                    <td style={{padding:`${pd}px 6px`,textAlign:"center"}}>{perms?.canSendSupport&&<input type="checkbox" checked={selIds.has(r.id)} onChange={()=>toggleSel(r.id)} style={{cursor:"pointer"}}/>}</td>
                    <td style={{padding:`${pd}px 10px`,fontWeight:500,color:"#000"}}>{r.nuvem}</td>
                    <td style={{padding:`${pd}px 10px`,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"#000"}} title={r.destinatario}>{r.destinatario}</td>
                    <td style={{padding:`${pd}px 10px`,color:"#666",overflow:"hidden",textOverflow:"ellipsis"}}>{r.transportadora}</td>
                    <td style={{padding:`${pd}px 10px`,color:"#666",fontFamily:"monospace",fontSize:11,overflow:"hidden",textOverflow:"ellipsis"}}>{r.rastreio}</td>
                    <td style={{padding:`${pd}px 10px`}}><StatusBadge val={r.status}/></td>
                    <td style={{padding:`${pd}px 10px`}}><SlaCell prazo={r.prazo}/></td>
                    <td style={{padding:`${pd}px 10px`,color:"#666"}}>{r.nf}</td>
                    <td style={{padding:`${pd}px 10px`,color:"#666",fontSize:11,overflow:"hidden",textOverflow:"ellipsis"}} title={r.motivo}>{r.motivo}</td>
                    <td style={{padding:`${pd}px 10px`}}><Chip val={r.urgencia} map={urgMap}/></td>
                    <td style={{padding:`${pd}px 10px`}}><Chip val={r.acionar} map={acionMap}/></td>
                    <td style={{padding:`${pd}px 10px`}}>{perms?.canSendSupport&&<button onClick={()=>upd(r.id,{enviadoSuporte:true,atendimento:"Aberto",sentAt:new Date().toISOString()},{acao:"Enviado ao suporte"})} style={{background:"transparent",border:"1px solid #ccc",color:"#333",borderRadius:2,padding:"3px 10px",fontSize:11,cursor:"pointer",width:"100%"}}>Enviar →</button>}</td>
                    <td style={{padding:`${pd}px 6px`,textAlign:"center"}}>{perms?.canDelete&&<button onClick={()=>del(r.id)} style={{background:"transparent",border:"none",color:"#ccc",cursor:"pointer",fontSize:15}}>×</button>}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginTop:10}}>
            <div style={{fontSize:11,color:"#bbb"}}>{filteredLog.length===0?"Nenhum resultado":`Mostrando ${((safeP-1)*PAGE_SIZE)+1}–${Math.min(safeP*PAGE_SIZE,filteredLog.length)} de ${filteredLog.length} pedidos`}</div>
            {totalPages>1&&<div style={{display:"flex",gap:4,alignItems:"center"}}><button onClick={()=>setLPage(n=>Math.max(1,n-1))} disabled={safeP===1} style={{...SEL,padding:"4px 10px",cursor:safeP===1?"not-allowed":"pointer",opacity:safeP===1?0.4:1}}>‹</button><span style={{fontSize:12,color:"#bbb",padding:"0 8px"}}>{safeP} / {totalPages}</span><button onClick={()=>setLPage(n=>Math.min(totalPages,n+1))} disabled={safeP===totalPages} style={{...SEL,padding:"4px 10px",cursor:safeP===totalPages?"not-allowed":"pointer",opacity:safeP===totalPages?0.4:1}}>›</button></div>}
          </div>
        </div>
      )}

      {/* SUPORTE */}
      {tab==="suporte"&&(
        <div style={{display:"flex",minHeight:520}}>
          <div style={{width:detail?"37%":"100%",borderRight:detail?"1px solid #ebebeb":"none",display:"flex",flexDirection:"column",background:"#fff"}}>
            <div style={{padding:"14px 16px 12px",borderBottom:"1px solid #f0f0f0"}}>
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8,marginBottom:12}}>
                {[{label:"Na fila",v:ss.total},{label:"Abertos",v:ss.abertos,red:true},{label:"Em andamento",v:ss.andamento}].map(s=>(
                  <div key={s.label} style={{textAlign:"center",padding:"8px 0",borderRadius:4,background:"#fafafa",border:"1px solid #f0f0f0"}}>
                    <div style={{fontSize:20,fontWeight:500,color:s.red&&s.v>0?"#c0392b":GOLD}}>{s.v}</div>
                    <div style={{fontSize:9,color:"#aaa",marginTop:2,letterSpacing:"0.06em",textTransform:"uppercase"}}>{s.label}</div>
                  </div>
                ))}
              </div>
              <div style={{display:"flex",gap:6,marginBottom:selSupIds.size>0?8:0}}>
                <input value={sSrch} onChange={e=>setSSrch(e.target.value)} placeholder="Buscar..." style={{...SEL,flex:1,fontSize:11}}/>
                <select value={sAtend} onChange={e=>setSAtend(e.target.value)} style={{...SEL,fontSize:11}}>{["Todos","Aberto","Em andamento"].map(o=><option key={o}>{o}</option>)}</select>
                <select value={sUrg} onChange={e=>setSUrg(e.target.value)} style={{...SEL,fontSize:11}}>{["Todos","Alta","Média","Baixa"].map(o=><option key={o}>{o}</option>)}</select>
              </div>
              {perms?.canOperate && selSupIds.size>0&&(
                <div style={{background:"#000",borderRadius:3,padding:"8px 12px",display:"flex",alignItems:"center",gap:8}}>
                  <span style={{color:"#aaa",fontSize:12,flex:1}}>{selSupIds.size} selecionado{selSupIds.size>1?"s":""}</span>
                  <button onClick={bulkArchive} style={{background:GOLD,border:"none",color:"#111",borderRadius:2,padding:"4px 12px",fontSize:11,cursor:"pointer",fontWeight:500}}>Arquivar ({selSupIds.size})</button>
                  <button onClick={()=>setSelSupIds(new Set())} style={{background:"transparent",border:"none",color:"#666",fontSize:11,cursor:"pointer"}}>✕</button>
                </div>
              )}
            </div>
            <div style={{overflowY:"auto",flex:1,maxHeight:"64vh"}}>
              {ss.total===0?<div style={{textAlign:"center",padding:"52px 20px",color:"#ccc"}}><div style={{fontSize:26,marginBottom:8,opacity:0.4}}>◎</div><div style={{fontSize:14,marginBottom:4}}>Fila vazia</div></div>
              :supRows.length===0?<div style={{textAlign:"center",padding:24,color:"#ccc",fontSize:13}}>Nenhum resultado</div>
              :supRows.map(r=>{
                const isSel=selSup===r.id, acColor=r.urgencia==="Alta"?"#e74c3c":r.urgencia==="Média"?GOLD:"#27ae60"
                return(
                  <div key={r.id} onClick={()=>setSelSup(isSel?null:r.id)} style={{padding:"11px 16px 11px 12px",cursor:"pointer",borderBottom:"1px solid #f5f5f5",borderLeft:`3px solid ${acColor}`,background:isSel?"#fdf8f0":r.alertaStatus?"#fffbf0":"#fff",display:"flex",alignItems:"flex-start",gap:8}}>
                    {perms?.canOperate&&<input type="checkbox" checked={selSupIds.has(r.id)} onClick={e=>e.stopPropagation()} onChange={()=>toggleSelSup(r.id)} style={{marginTop:3,cursor:"pointer",flexShrink:0}}/>}
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:3}}><span style={{fontWeight:500,fontSize:12,color:"#000"}}>{r.nuvem}</span><TimeOpenBadge sentAt={r.sentAt}/></div>
                      <div style={{fontSize:12,color:"#555",marginBottom:5,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.destinatario}</div>
                      {r.alertaStatus&&(
                        <div style={{background:"#fff3cd",border:"1px solid #ffeeba",borderRadius:3,padding:"4px 8px",fontSize:11,color:"#856404",marginBottom:5,display:"flex",alignItems:"center",gap:6}}>
                          <span>⚠</span>
                          <span><b>Status alterado:</b> {r.alertaStatus} — notifique o cliente!</span>
                        </div>
                      )}
                      <div style={{display:"flex",gap:4,flexWrap:"wrap"}}><StatusBadge val={r.status}/><Chip val={r.urgencia} map={urgMap}/><Chip val={r.atendimento} map={atendMap}/></div>
                      {r.responsavel&&<div style={{fontSize:10,color:"#aaa",marginTop:5}}>Resp: {r.responsavel}</div>}
                    </div>
                  </div>
                )
              })}
            </div>
          </div>
          {detail?(
            <div style={{flex:1,display:"flex",flexDirection:"column",overflowY:"auto",maxHeight:"73vh",background:"#fafafa"}}>
              <div style={{background:"#fff",padding:"18px 24px 14px",borderBottom:"1px solid #e8e8e8",position:"sticky",top:0,zIndex:5}}>
                <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:10}}>
                  <div><div style={{fontSize:10,color:"#aaa",letterSpacing:"0.1em",textTransform:"uppercase",marginBottom:4}}>Pedido em atendimento</div><div style={{fontSize:18,fontWeight:400,color:"#000",marginBottom:3}}>{detail.destinatario}</div><div style={{fontSize:12,color:"#888"}}>#{detail.nuvem} · NF {detail.nf} · {detail.transportadora}</div></div>
                  <button onClick={()=>setSelSup(null)} style={{background:"transparent",border:"none",color:"#ccc",cursor:"pointer",fontSize:22,padding:"0 4px",lineHeight:1}}>×</button>
                </div>
                <div style={{display:"flex",gap:5,flexWrap:"wrap",marginBottom:14}}><StatusBadge val={detail.status}/><Chip val={detail.urgencia} map={urgMap}/><Chip val={detail.atendimento} map={atendMap}/><TimeOpenBadge sentAt={detail.sentAt}/></div>
                {detail.alertaStatus&&(
                  <div style={{background:"#fff3cd",border:"1px solid #ffeeba",borderRadius:4,padding:"10px 14px",marginBottom:14,display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:12}}>
                    <div>
                      <div style={{fontSize:12,fontWeight:500,color:"#856404",marginBottom:4}}>⚠ Status do pedido foi atualizado — notifique o cliente!</div>
                      <div style={{fontSize:11,color:"#856404"}}>{detail.alertaStatus}</div>
                      <div style={{fontSize:11,color:"#999",marginTop:4}}>Use os textos prontos abaixo para contatar via WhatsApp ou Zendesk</div>
                    </div>
                    <button onClick={()=>upd(detail.id,{alertaStatus:null},{acao:"Alerta de status dispensado — cliente notificado"})}
                      style={{background:"#856404",border:"none",color:"#fff",borderRadius:3,padding:"5px 12px",fontSize:11,cursor:"pointer",whiteSpace:"nowrap",flexShrink:0}}>
                      ✓ Notifiquei
                    </button>
                  </div>
                )}
                {perms?.canOperate&&(
                  <div style={{display:"flex",gap:8}}>
                    {detail.atendimento==="Aberto"&&<button onClick={()=>handleInitiate(detail.id)} style={{flex:1,background:"#000",border:"none",color:"#fff",borderRadius:3,padding:"10px 0",fontSize:12,cursor:"pointer",fontWeight:500,letterSpacing:"0.08em",textTransform:"uppercase"}}>Iniciar atendimento</button>}
                    {detail.atendimento==="Em andamento"&&<button onClick={()=>handleResolve(detail.id)} style={{flex:1,background:GOLD,border:"none",color:"#111",borderRadius:3,padding:"10px 0",fontSize:12,cursor:"pointer",fontWeight:500}}>Marcar como resolvido →</button>}
                    <button onClick={()=>handleReturnLog(detail.id)} style={{background:"transparent",border:"1px solid #d5d5d5",color:"#666",borderRadius:3,padding:"10px 14px",fontSize:12,cursor:"pointer",whiteSpace:"nowrap"}}>← Devolver</button>
                  </div>
                )}
              </div>
              <div style={{padding:"20px 24px",flex:1}}>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:18,background:"#fff",borderRadius:4,border:"1px solid #e8e8e8",padding:16}}>
                  {[["Rastreio",detail.rastreio||"—"],["Motivo",detail.motivo],["Transportadora",detail.transportadora]].map(([lbl,val])=>(
                    <div key={lbl}><div style={{fontSize:9,color:"#aaa",textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:4}}>{lbl}</div><div style={{fontSize:12,color:"#000",fontFamily:lbl==="Rastreio"?"monospace":"inherit"}}>{val}</div></div>
                  ))}
                  <div><div style={{fontSize:9,color:"#aaa",textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:4}}>Prazo / SLA</div><SlaCell prazo={detail.prazo}/></div>
                </div>
                {perms?.canOperate&&(
                  <>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:14}}>
                      <div><div style={{fontSize:9,color:"#aaa",textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:6}}>Responsável</div><input value={detail.responsavel||""} onChange={e=>upd(detail.id,{responsavel:e.target.value})} placeholder="Nome do responsável..." style={{...SEL,width:"100%",boxSizing:"border-box",padding:"8px 10px",fontSize:12}}/></div>
                      <div><div style={{fontSize:9,color:"#aaa",textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:6}}>Nº Chamado Zendesk</div><input value={detail.chamado||""} onChange={e=>upd(detail.id,{chamado:e.target.value})} placeholder="Ex: #45821" style={{...SEL,width:"100%",boxSizing:"border-box",padding:"8px 10px",fontSize:12}}/></div>
                    </div>
                    <div style={{marginBottom:20}}><div style={{fontSize:9,color:"#aaa",textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:6}}>Observações</div><textarea value={detail.obs||""} onChange={e=>upd(detail.id,{obs:e.target.value})} placeholder="Anotações do atendimento..." rows={3} style={{width:"100%",borderRadius:3,border:"1px solid #d5d5d5",padding:"8px 10px",fontSize:12,resize:"vertical",fontFamily:"inherit",background:"#fff",color:"#000",boxSizing:"border-box"}}/></div>
                  </>
                )}
                <div style={{height:1,background:"#ebebeb",marginBottom:16}}/>
                <div style={{marginBottom:10,border:"1px solid #e8e8e8",borderRadius:4,overflow:"hidden",background:"#fff"}}>
                  <button onClick={()=>setOpenTpl(v=>!v)} style={{width:"100%",padding:"12px 16px",background:"transparent",border:"none",cursor:"pointer",display:"flex",justifyContent:"space-between",alignItems:"center",fontSize:12,fontWeight:500,color:"#000"}}>
                    <span>✉ Textos prontos para atendimento</span><span style={{fontSize:18,color:"#ccc",fontWeight:300}}>{openTpl?"−":"+"}</span>
                  </button>
                  {openTpl&&(
                    <div style={{padding:14,borderTop:"1px solid #f0f0f0",background:"#fafafa"}}>
                      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                        {[{ch:"wpp",label:"WhatsApp"},{ch:"zendesk",label:"Zendesk"}].map(({ch,label})=>(
                          <div key={ch} style={{background:"#fff",borderRadius:3,border:"1px solid #e0e0e0",overflow:"hidden"}}>
                            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"8px 12px",borderBottom:"1px solid #e0e0e0",background:"#000"}}><span style={{fontSize:11,color:GOLD,fontWeight:500,letterSpacing:"0.06em"}}>{label}</span><CopyBtn text={getTemplate(detail,ch)} label="Copiar"/></div>
                            <pre style={{margin:0,padding:"10px 12px",fontSize:11,color:"#555",whiteSpace:"pre-wrap",fontFamily:"inherit",lineHeight:1.6,maxHeight:160,overflowY:"auto"}}>{getTemplate(detail,ch)}</pre>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
                <div style={{border:"1px solid #e8e8e8",borderRadius:4,overflow:"hidden",background:"#fff"}}>
                  <button onClick={()=>setOpenHist(v=>!v)} style={{width:"100%",padding:"12px 16px",background:"transparent",border:"none",cursor:"pointer",display:"flex",justifyContent:"space-between",alignItems:"center",fontSize:12,fontWeight:500,color:"#000"}}>
                    <span>Histórico{detail.historico.length>0?` (${detail.historico.length})`:""}</span><span style={{fontSize:18,color:"#ccc",fontWeight:300}}>{openHist?"−":"+"}</span>
                  </button>
                  {openHist&&<div style={{padding:14,borderTop:"1px solid #f0f0f0",background:"#fafafa"}}>{detail.historico.length===0?<div style={{fontSize:12,color:"#bbb"}}>Nenhuma ação registrada.</div>:<div style={{display:"flex",flexDirection:"column",gap:8}}>{detail.historico.map((h,i)=><div key={i} style={{display:"flex",gap:10,alignItems:"center",fontSize:12}}><span style={{color:"#bbb",whiteSpace:"nowrap",fontSize:11,minWidth:130}}>{h.ts}</span><span style={{width:5,height:5,borderRadius:"50%",background:GOLD,flexShrink:0}}/><span style={{color:"#333"}}>{h.acao}</span></div>)}</div>}</div>}
                </div>
              </div>
            </div>
          ):ss.total>0?<div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center",background:"#fafafa"}}><div style={{textAlign:"center"}}><div style={{fontSize:11,letterSpacing:"0.15em",textTransform:"uppercase",color:"#ccc"}}>Selecione um pedido para atender</div></div></div>:null}
        </div>
      )}

      {/* ARQUIVADOS */}
      {tab==="arquivados"&&(
        <div style={{padding:"16px 24px"}}>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12,marginBottom:16}}>
            <StatCard label="Total arquivados" val={arch}/>
            <StatCard label="Resolvidos hoje" val={rows.filter(r=>{if(r.atendimento!=="Resolvido")return false;const h=r.historico.find(x=>x.acao&&(x.acao.includes("Resolvido")||x.acao.includes("Arquivado")));return h&&h.ts&&h.ts.startsWith(new Date().toLocaleDateString("pt-BR"))}).length}/>
            <StatCard label="Com observações" val={baseArch.filter(r=>r.obs&&r.obs.trim()).length}/>
          </div>
          {arch===0?<div style={{textAlign:"center",padding:"44px 0",color:"#ccc"}}><div style={{fontSize:28,marginBottom:10,opacity:0.3}}>◎</div><div style={{fontSize:14,marginBottom:6}}>Nenhum atendimento arquivado</div></div>:(
            <div>
              <div style={{marginBottom:12}}><input value={aSrch} onChange={e=>setASrch(e.target.value)} placeholder="Buscar nos arquivados..." style={{...SEL,width:"100%",padding:"7px 10px",boxSizing:"border-box"}}/></div>
              <div style={{overflowX:"auto",overflowY:"auto",maxHeight:"60vh",borderRadius:4,border:"1px solid #e0e0e0"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:compact?11:12,tableLayout:"fixed",minWidth:900}}>
                  <colgroup><col style={{width:78}}/><col style={{width:135}}/><col style={{width:95}}/><col style={{width:95}}/><col style={{width:105}}/><col style={{width:62}}/><col style={{width:98}}/><col style={{width:90}}/><col style={{width:80}}/><col style={{width:88}}/></colgroup>
                  <thead><tr>{["No NUVEM","Destinatário","Transportadora","Status","Motivo","Urgência","Prazo / SLA","Chamado","Responsável","Ações"].map(h=><th key={h} style={thF}>{h}</th>)}</tr></thead>
                  <tbody>
                    {archRows.length===0?<tr><td colSpan={10} style={{textAlign:"center",padding:28,color:"#ccc"}}>Nenhum resultado</td></tr>
                    :archRows.map((r,i)=>(
                      <tr key={r.id} style={{background:i%2===0?"#fff":"#fafafa",borderBottom:"1px solid #f2f2f2"}}>
                        <td style={{padding:`${pd}px 10px`,fontWeight:500,color:"#aaa"}}>{r.nuvem}</td>
                        <td style={{padding:`${pd}px 10px`,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"#000"}} title={r.destinatario}>{r.destinatario}</td>
                        <td style={{padding:`${pd}px 10px`,color:"#666",overflow:"hidden",textOverflow:"ellipsis"}}>{r.transportadora}</td>
                        <td style={{padding:`${pd}px 10px`}}><StatusBadge val={r.status}/></td>
                        <td style={{padding:`${pd}px 10px`,color:"#666",fontSize:11,overflow:"hidden",textOverflow:"ellipsis"}} title={r.motivo}>{r.motivo}</td>
                        <td style={{padding:`${pd}px 10px`}}><Chip val={r.urgencia} map={urgMap}/></td>
                        <td style={{padding:`${pd}px 10px`}}><SlaCell prazo={r.prazo}/></td>
                        <td style={{padding:`${pd}px 10px`,color:"#888",fontSize:11}}>{r.chamado||"—"}</td>
                        <td style={{padding:`${pd}px 10px`,color:"#888",fontSize:11,overflow:"hidden",textOverflow:"ellipsis"}}>{r.responsavel||"—"}</td>
                        <td style={{padding:`${pd}px 10px`}}>{perms?.canOperate&&<button onClick={()=>upd(r.id,{atendimento:"Em andamento"},{acao:"Reaberto dos arquivados"})} style={{background:"transparent",border:"1px solid #ccc",color:"#333",borderRadius:2,padding:"3px 10px",fontSize:11,cursor:"pointer"}}>Reabrir</button>}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div style={{marginTop:8,fontSize:11,color:"#bbb"}}>{archRows.length} de {arch} arquivados</div>
            </div>
          )}
        </div>
      )}
    </div>
  )
}
