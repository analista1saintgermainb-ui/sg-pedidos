import { useState, useEffect, useRef, useCallback } from 'react'
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, LineChart, Line } from 'recharts'
import * as XLSX from 'xlsx'

// ─── Design System ────────────────────────────────────────────
const C = {
  brand:    "#0C0C0C",
  brandSoft:"#1A1A1A",
  gold:     "#B8974A",
  goldLight:"#D4AF6A",
  goldDim:  "#8C7038",
  cream:    "#F8F5EF",
  creamDark:"#F0EDE5",
  white:    "#FFFFFF",
  border:   "#E8E3D8",
  borderDark:"#D4CFC4",
  text1:    "#1A1A1A",
  text2:    "#5C5750",
  text3:    "#9C9690",
  text4:    "#C4C0B8",
  red:      "#C0392B",
  redSoft:  "#F9ECEB",
  redBorder:"#EBCBC8",
  green:    "#2E7D50",
  greenSoft:"#EAF4EE",
  greenBorder:"#C0DCCB",
  amber:    "#8C6D1F",
  amberSoft:"#FDF6E3",
  amberBorder:"#E8D5A3",
  blue:     "#1A5276",
  blueSoft: "#EAF2FB",
  blueBorder:"#AACDE6",
}

const shadow = {
  sm:  "0 1px 3px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04)",
  md:  "0 4px 12px rgba(0,0,0,0.08), 0 2px 4px rgba(0,0,0,0.04)",
  lg:  "0 8px 24px rgba(0,0,0,0.10), 0 4px 8px rgba(0,0,0,0.06)",
}

// Global styles
const globalStyle = `
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

async function signIn(email, password) {
  const r = await fetch(`${SUPA_URL}/auth/v1/token?grant_type=password`, { method:"POST", headers:SH, body:JSON.stringify({email,password}) })
  const d = await r.json(); if(!r.ok) throw new Error(d.error_description||d.msg||"Erro ao fazer login"); return d
}
async function signOut(token) { await fetch(`${SUPA_URL}/auth/v1/logout`,{method:"POST",headers:aSH(token)}) }
async function createUser(email, password) {
  const r = await fetch(`${SUPA_URL}/auth/v1/admin/users`,{method:"POST",headers:{apikey:SUPA_SERVICE,Authorization:`Bearer ${SUPA_SERVICE}`,"Content-Type":"application/json"},body:JSON.stringify({email,password,email_confirm:true})})
  const d = await r.json(); if(!r.ok) throw new Error(d.msg||d.message||"Erro ao criar usuário"); return d
}
async function dbLoadFast(token, onPartial) {
  let all = [], from = 0, step = 1000
  while (true) {
    const r = await fetch(`${SUPA_URL}/rest/v1/pedidos?select=*&order=id&limit=${step}&offset=${from}`, { headers: aSH(token) })
    if (!r.ok) throw new Error(await r.text())
    const data = await r.json()
    const chunk = data.map(row => ({ ...row.dados, id: row.id }))
    all = [...all, ...chunk]
    // Mostra primeiros 1000 imediatamente sem esperar o resto
    if (from === 0 && chunk.length > 0) onPartial(all)
    if (data.length < step) break
    from += step
  }
  return all
}
async function dbUpsert(rows,token) {
  for(let i=0;i<rows.length;i+=200){
    const batch=rows.slice(i,i+200)
    const payload=batch.map(r=>({id:Number(r.id),dados:r,updated_at:new Date().toISOString()}))
    const r2=await fetch(`${SUPA_URL}/rest/v1/pedidos`,{method:"POST",headers:{...aSH(token),Prefer:"resolution=merge-duplicates,return=minimal"},body:JSON.stringify(payload)})
    if(!r2.ok) throw new Error(`${r2.status}: ${await r2.text()}`)
  }
}
async function dbDelete(id,token){ await fetch(`${SUPA_URL}/rest/v1/pedidos?id=eq.${id}`,{method:"DELETE",headers:aSH(token)}) }
async function dbClear(token){ await fetch(`${SUPA_URL}/rest/v1/pedidos?id=gte.0`,{method:"DELETE",headers:aSH(token)}) }
async function loadUsuarios(token){ const r=await fetch(`${SUPA_URL}/rest/v1/usuarios?select=*&order=created_at`,{headers:aSH(token)}); return r.ok?r.json():[] }
async function saveUsuario(u,token){ await fetch(`${SUPA_URL}/rest/v1/usuarios`,{method:"POST",headers:{...aSH(token),Prefer:"resolution=merge-duplicates,return=minimal"},body:JSON.stringify(u)}) }
async function deleteUsuario(id,token){ await fetch(`${SUPA_URL}/rest/v1/usuarios?id=eq.${id}`,{method:"DELETE",headers:aSH(token)}) }

// ─── Permissões ───────────────────────────────────────────────
const PERMS = {
  admin:    {tabs:["dashboard","logistica","suporte","arquivados","usuarios"],canImport:true,canDelete:true,canClear:true,canSendSupport:true,canOperate:true},
  logistica:{tabs:["dashboard","logistica"],canImport:true,canDelete:false,canClear:false,canSendSupport:true,canOperate:true},
  suporte:  {tabs:["suporte","arquivados"],canImport:false,canDelete:false,canClear:false,canSendSupport:false,canOperate:true},
  leitura:  {tabs:["dashboard","logistica","suporte","arquivados"],canImport:false,canDelete:false,canClear:false,canSendSupport:false,canOperate:false},
}

// ─── Helpers ──────────────────────────────────────────────────
const HEADER_MAP = {
  nuvem:["identificador ecommerce","id ecommerce","no nuvem","nuvem","pedido"],
  destinatario:["destinatário nome","destinatario nome","destinatário","destinatario","nome do pedido","nome do cliente","nome","cliente"],
  transportadora:["estratégia de frete","estrategia de frete","transportadora","frete"],
  rastreio:["rastreador last mile","código de rastreio","codigo de rastreio","rastreio","last mile"],
  status:["situação","situacao","situac","status"],
  prazo:["prazo logístico","prazo logistico","prazo"],
  nf:["nº nota fiscal","no nota fiscal","nota fiscal","no nf","nf"],
  ultimaMov:["última movimentação","ultima movimentacao","ultima movimentação","última movimentacao","ultima mov","última mov","movimentacao","movimentação"],
}
const norm = s => (s||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().trim()
const findIdx = (hdrs,key) => hdrs.findIndex(h=>(HEADER_MAP[key]||[]).some(v=>norm(h).includes(norm(v))))
const uniq = arr => ["Todos",...Array.from(new Set(arr.filter(Boolean).sort()))]
const QFILTERS = [
  {id:"todos",label:"Todos"},
  {id:"urgente",label:"Urgente"},
  {id:"extraviados",label:"Extraviados"},
  {id:"devolvidos",label:"Devolvidos"},
  {id:"vence_hoje",label:"Vence hoje"},
  {id:"vencidos",label:"Vencidos"},
  {id:"parados",label:`Parados +${ALERTA_DIAS}d`},
]
const PAGE_SIZE = 50

const ALERTA_DIAS = 7

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
  if (dias >= ALERTA_DIAS) return { dias, label: `${dias}d sem atualização`, alerta: true }
  if (dias >= 3)            return { dias, label: `${dias}d sem atualização`, alerta: false }
  return null
} const s=(status||"").toLowerCase(); return s.includes("entregue")||s.includes("finaliz")||s.includes("entrega realizada") }
function calcMotivo(s){
  const v=(s||"").toLowerCase()
  if(v.includes("extravia")||v.includes("perdid"))return"Objeto extraviado"
  if(v.includes("devolv")||v.includes("recusa"))return"Devolução / Recusa"
  if(v.includes("atras"))return"Atraso na entrega"
  if(v.includes("entregue")||v.includes("finaliz"))return"Entrega concluída"
  if(v.includes("saiu"))return"Saiu para entrega"
  if(v.includes("trânsito")||v.includes("transito"))return"Em trânsito"
  if(v.includes("postado")||v.includes("coletado"))return"Aguardando movimentação"
  if(v.includes("pendente")||v.includes("aguardando")||v.includes("processando"))return"Aguardando coleta"
  return"—"
}
function parsePrazo(v){
  if(!v)return null
  const c=v.replace(/[^\d\/\-\.]/g,"")
  for(const t of[c,c.split("/").reverse().join("-"),c.split(".").reverse().join("-")]){
    const d=new Date(t);if(!isNaN(d.getTime()))return d
  }
  return null
}
function calcUrg(prazo,status){
  const s=(status||"").toLowerCase()
  if(s.includes("extravia")||s.includes("devolv"))return"Alta"
  if(s.includes("entregue")||s.includes("finaliz"))return"Baixa"
  const dt=parsePrazo(prazo);if(!dt)return"—"
  const h=new Date();h.setHours(0,0,0,0)
  const d=Math.ceil((dt-h)/86400000)
  if(d<=1)return"Alta";if(d<=3)return"Média";return"Baixa"
}
function calcAcionar(urg,status){
  const s=(status||"").toLowerCase()
  if(urg==="Alta"||s.includes("extravia")||s.includes("devolv"))return"Sim"
  if(urg==="Média")return"Avaliar"
  return"Não"
}
function slaInfo(prazo){
  const dt=parsePrazo(prazo);if(!dt)return null
  const h=new Date();h.setHours(0,0,0,0)
  const diff=Math.ceil((dt-h)/86400000)
  if(diff<0)  return{label:`${Math.abs(diff)}d vencido`,color:C.red,bg:C.redSoft,bd:C.redBorder}
  if(diff===0)return{label:"Vence hoje",              color:C.amber,bg:C.amberSoft,bd:C.amberBorder}
  if(diff<=3) return{label:`${diff}d restantes`,      color:C.amber,bg:C.amberSoft,bd:C.amberBorder}
  return          {label:`${diff}d restantes`,         color:C.green,bg:C.greenSoft,bd:C.greenBorder}
}
function timeOpen(sentAt){
  if(!sentAt)return null
  const ms=Date.now()-new Date(sentAt).getTime()
  const h=Math.floor(ms/3600000),d=Math.floor(h/24)
  if(d>0)return{label:`${d}d na fila`,alert:d>=2}
  if(h>0)return{label:`${h}h na fila`,alert:false}
  return{label:"< 1h",alert:false}
}
function parseData(text){
  const sep=text.includes("\t")?"\t":text.includes(";")?";":","
  const lines=text.trim().split("\n").filter(l=>l.trim())
  if(!lines.length)return[]
  const first=lines[0].split(sep).map(h=>h.trim().replace(/^["']|["']$/g,""))
  const isHdr=first.some(h=>["nuvem","destinat","identificador","ecommerce","rastreio","situac","status","frete"].some(k=>norm(h).includes(k)))
  const hdrs=isHdr?first:[]
  const data=isHdr?lines.slice(1):lines
  const ix={nuvem:isHdr?findIdx(hdrs,"nuvem"):0,dest:isHdr?findIdx(hdrs,"destinatario"):1,transp:isHdr?findIdx(hdrs,"transportadora"):2,rastreio:isHdr?findIdx(hdrs,"rastreio"):3,status:isHdr?findIdx(hdrs,"status"):4,prazo:isHdr?findIdx(hdrs,"prazo"):5,nf:isHdr?findIdx(hdrs,"nf"):6}
  const g=(c,i)=>i>=0&&i<c.length?c[i]:""
  return data.map((line,i)=>{
    const c=line.split(sep).map(v=>v.trim().replace(/^["']|["']$/g,""))
    const status=g(c,ix.status),prazo=g(c,ix.prazo),urg=calcUrg(prazo,status)
    const entregue=isEntregue(status)
    const dt=parsePrazo(prazo)
    const noPrazo=entregue&&dt?new Date()<=dt:entregue&&dt?false:null
    return{id:Date.now()+i,nuvem:g(c,ix.nuvem),destinatario:g(c,ix.dest),transportadora:g(c,ix.transp),rastreio:g(c,ix.rastreio),status,prazo,nf:g(c,ix.nf),ultimaMov:g(c,ix.ultimaMov),cidade:g(c,ix.cidade),uf:g(c,ix.uf),motivo:calcMotivo(status),urgencia:urg,acionar:calcAcionar(urg,status),enviadoSuporte:false,atendimento:entregue?"Resolvido":"Aberto",entregueNoPrazo:noPrazo,alertaStatus:null,obs:"",historico:entregue?[{acao:"Arquivado automaticamente — entrega concluída",ts:new Date().toLocaleString("pt-BR")}]:[],responsavel:"",sentAt:null,chamado:"",isNew:true}
  }).filter(r=>r.nuvem||r.destinatario||r.nf)
}
function applyQF(rows,qf){
  if(qf==="todos")return rows
  if(qf==="urgente")return rows.filter(r=>r.urgencia==="Alta")
  if(qf==="extraviados")return rows.filter(r=>(r.status||"").toLowerCase().includes("extravia"))
  if(qf==="devolvidos")return rows.filter(r=>(r.status||"").toLowerCase().includes("devolv"))
  if(qf==="vence_hoje")return rows.filter(r=>{const d=parsePrazo(r.prazo);if(!d)return false;const h=new Date();h.setHours(0,0,0,0);return Math.ceil((d-h)/86400000)<=0})
  if(qf==="vencidos")return rows.filter(r=>{const d=parsePrazo(r.prazo);if(!d)return false;const h=new Date();h.setHours(0,0,0,0);return d<h})
  return rows
}
function applySortRows(rows,col,dir){
  if(!col)return rows
  return[...rows].sort((a,b)=>{
    let va=a[col]||"",vb=b[col]||""
    if(col==="prazo"){va=parsePrazo(va)||new Date(0);vb=parsePrazo(vb)||new Date(0)}
    if(col==="urgencia"){const o={Alta:0,Média:1,Baixa:2,"—":3};va=o[va]??9;vb=o[vb]??9}
    const cmp=typeof va==="object"?va-vb:String(va).localeCompare(String(vb),"pt-BR")
    return dir==="asc"?cmp:-cmp
  })
}
function exportCSV(rows){
  const h=["No NUVEM","Destinatário","Transportadora","Cód. Rastreio","Status","Prazo","No NF","Motivo","Urgência","Acionar?","Suporte","Atendimento","Chamado","Responsável","Observações"]
  const e=v=>`"${String(v||"").replace(/"/g,'""')}"`
  const csv=[h.map(e).join(";"),...rows.map(r=>[r.nuvem,r.destinatario,r.transportadora,r.rastreio,r.status,r.prazo,r.nf,r.motivo,r.urgencia,r.acionar,r.enviadoSuporte?"Sim":"Não",r.atendimento,r.chamado,r.responsavel,r.obs].map(e).join(";"))].join("\n")
  const a=document.createElement("a")
  a.href=URL.createObjectURL(new Blob(["\uFEFF"+csv],{type:"text/csv;charset=utf-8;"}))
  a.download="saint_germain_pedidos.csv";a.click()
}
function getTemplate(r, ch, nomeAtendente) {
  const nome  = (r.destinatario||"").split(" ")[0] || "Cliente"
  const atend = (nomeAtendente||"").split(" ")[0]  || "Time SG"
  const m     = (r.motivo||"").toLowerCase()
  const extrav = m.includes("extravia")
  const devolv = m.includes("devolu") || m.includes("recusa")
  const atraso = m.includes("atraso")
  const parado = diasSemMov(r.ultimaMov) !== null && diasSemMov(r.ultimaMov) >= ALERTA_DIAS
  const linkRastreio = r.rastreio
    ? `https://saintgermain.rastreio.estoca.com.br/tracking?code=${r.rastreio}`
    : null

  if (ch === "wpp") {
    if (extrav) return `Olá, ${nome}! Tudo bem? Me chamo ${atend} e faço parte do time de encantamento da SG 🤍\n\nEstou entrando em contato pois identificamos uma ocorrência no seu pedido *#${r.nuvem}*.\n\nSua encomenda está com o status de *objeto extraviado* junto à transportadora *${r.transportadora}*. Já acionamos nossa equipe para apurar o caso com urgência.\n\nRetornaremos com uma atualização em até *2 dias úteis*. Pedimos sinceras desculpas pelo transtorno! 🙏\n\nQualquer dúvida estamos à disposição 🤍\n${atend} — Time de Encantamento SG`

    if (devolv) return `Olá, ${nome}! Tudo bem? Me chamo ${atend} e faço parte do time de encantamento da SG 🤍\n\nEstou entrando em contato sobre seu pedido *#${r.nuvem}*.\n\nIdentificamos que sua encomenda foi *devolvida* ao nosso centro de distribuição após tentativas de entrega sem sucesso. 😔\n\nPara realizarmos um novo envio sem nenhum custo adicional, poderia confirmar seu endereço de entrega completo respondendo esta mensagem?\n\nEstamos aqui para resolver isso da melhor forma para você 🤍\n${atend} — Time de Encantamento SG`

    if (atraso) return `Olá, ${nome}! Tudo bem? Me chamo ${atend} e faço parte do time de encantamento da SG 🤍\n\nEstou entrando em contato sobre seu pedido *#${r.nuvem}*.\n\nIdentificamos um atraso na entrega pela transportadora *${r.transportadora}*. O prazo previsto era *${r.prazo||"—"}* e já estamos acompanhando de perto junto à transportadora.\n\nAssim que tivermos uma atualização, te avisamos imediatamente! Pedimos desculpas pelo inconveniente 🙏\n\nQualquer dúvida estamos à disposição 🤍\n${atend} — Time de Encantamento SG`

    if (parado) return `Olá, ${nome}! Tudo bem? Me chamo ${atend} e faço parte do time de encantamento da SG 🤍\n\nEstou entrando em contato sobre seu pedido *#${r.nuvem}*.\n\nPercebemos que seu pedido está *em trânsito* com a transportadora *${r.transportadora}*, mas sem novas atualizações de rastreio nos últimos dias. Já estamos apurando a situação com a transportadora.\n\nRetornaremos em breve com uma atualização! Pedimos desculpas pela espera 🙏\n\nQualquer dúvida estamos à disposição 🤍\n${atend} — Time de Encantamento SG`

    return `Olá, ${nome}! Tudo bem? Me chamo ${atend} e faço parte do time de encantamento da SG 🤍\n\nVocê pode rastrear seu pedido *#${r.nuvem}* diretamente neste link:\n${linkRastreio||r.rastreio||"—"}\n\nStatus atual: *${r.status}*${r.prazo ? `\nPrazo previsto: *${r.prazo}*` : ""}\n\nSe tiver qualquer dúvida ou encontrar algum desafio, estamos à disposição para ajudar você! 🤍\n${atend} — Time de Encantamento SG`

  } else {
    const assinatura = `Atenciosamente,\n${atend}\nTime de Encantamento — Saint Germain`
    const det = `• Pedido: #${r.nuvem}\n• NF: ${r.nf}\n• Transportadora: ${r.transportadora}\n• Rastreio: ${r.rastreio||"—"}\n• Prazo previsto: ${r.prazo||"—"}`

    if (extrav) return `Assunto: Pedido #${r.nuvem} — Objeto Extraviado\n\nOlá, ${r.destinatario}! Tudo bem?\n\nMe chamo ${atend} e faço parte do time de encantamento da SG 🤍\n\nEntramos em contato para informar sobre uma ocorrência identificada no seu pedido.\n\n${det}\nStatus: Objeto extraviado\n\nJá acionamos nossa equipe de logística para apurar o caso com urgência junto à transportadora. Retornaremos com uma solução em até 2 dias úteis.\n\nPedimos sinceras desculpas pelo transtorno e agradecemos a sua compreensão.\n\n${assinatura}`

    if (devolv) return `Assunto: Pedido #${r.nuvem} — Devolução de Encomenda\n\nOlá, ${r.destinatario}! Tudo bem?\n\nMe chamo ${atend} e faço parte do time de encantamento da SG 🤍\n\nEntramos em contato para informar que sua encomenda retornou ao nosso centro de distribuição após tentativas de entrega sem sucesso.\n\n${det}\n\nPara realizarmos um novo envio sem nenhum custo adicional, pedimos que confirme seu endereço de entrega completo respondendo a este chamado.\n\nEstamos à disposição para resolver da melhor forma para você 🤍\n\n${assinatura}`

    if (atraso) return `Assunto: Pedido #${r.nuvem} — Atraso na Entrega\n\nOlá, ${r.destinatario}! Tudo bem?\n\nMe chamo ${atend} e faço parte do time de encantamento da SG 🤍\n\nEntramos em contato para informar que identificamos um atraso na entrega do seu pedido.\n\n${det}\n\nJá estamos acompanhando o caso junto à transportadora e te manteremos informado(a) sobre qualquer atualização.\n\nPedimos desculpas pelo inconveniente e agradecemos a sua paciência 🙏\n\n${assinatura}`

    if (parado) return `Assunto: Pedido #${r.nuvem} — Atualização de Rastreio\n\nOlá, ${r.destinatario}! Tudo bem?\n\nMe chamo ${atend} e faço parte do time de encantamento da SG 🤍\n\nIdentificamos que seu pedido está em trânsito, porém sem novas atualizações de rastreio nos últimos dias. Já estamos apurando a situação com a transportadora.\n\n${det}\n\nRetornaremos em breve com uma atualização. Pedimos desculpas pela espera 🙏\n\n${assinatura}`

    return `Assunto: Pedido #${r.nuvem} — Rastreio\n\nOlá, ${r.destinatario}! Tudo bem?\n\nMe chamo ${atend} e faço parte do time de encantamento da SG 🤍\n\nVocê pode rastrear seu pedido diretamente neste link:\n${linkRastreio||r.rastreio||"—"}\n\n${det}\nStatus atual: ${r.status}\n\nSe tiver qualquer dúvida ou encontrar algum desafio, estamos à disposição para ajudar você! 🤍\n\n${assinatura}`
  }
}

// ─── Design Tokens para badges ───────────────────────────────
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
  Aberto:        {bg:C.redSoft,  color:C.red,  bd:C.redBorder},
  "Em andamento":{bg:C.amberSoft,color:C.amber,bd:C.amberBorder},
  Resolvido:     {bg:C.greenSoft,color:C.green,bd:C.greenBorder},
}

// ─── Componentes ──────────────────────────────────────────────
function Chip({val,styles}){
  const s=styles[val]||{bg:C.creamDark,color:C.text3,bd:C.border}
  return<span style={{background:s.bg,color:s.color,border:`1px solid ${s.bd}`,borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:500,whiteSpace:"nowrap",letterSpacing:"0.02em"}}>{val}</span>
}

function StatusBadge({val}){
  const s=(val||"").toLowerCase()
  let bg=C.creamDark,color=C.text3,bd=C.border
  if(s.includes("entregue")||s.includes("finaliz")){bg=C.greenSoft;color=C.green;bd=C.greenBorder}
  else if(s.includes("trânsito")||s.includes("transito")){bg=C.blueSoft;color=C.blue;bd=C.blueBorder}
  else if(s.includes("saiu")){bg:"#EBF5FB";color=C.blue;bd=C.blueBorder}
  else if(s.includes("extravia")){bg=C.redSoft;color=C.red;bd=C.redBorder}
  else if(s.includes("devolv")){bg=C.amberSoft;color=C.amber;bd=C.amberBorder}
  return<span style={{background:bg,color,border:`1px solid ${bd}`,borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:500,whiteSpace:"nowrap",letterSpacing:"0.02em"}}>{val||"—"}</span>
}

function SlaCell({prazo}){
  const sla=slaInfo(prazo)
  return<div style={{lineHeight:1.6}}>
    <div style={{fontSize:11,color:C.text2,fontWeight:400}}>{prazo||"—"}</div>
    {sla&&<span style={{background:sla.bg,color:sla.color,border:`1px solid ${sla.bd}`,borderRadius:10,padding:"1px 7px",fontSize:10,fontWeight:500}}>{sla.label}</span>}
  </div>
}

function SemMovBadge({ultimaMov}){
  const info = semMovInfo(ultimaMov)
  if (!info) return <span style={{fontSize:11,color:C.text4}}>{ultimaMov||"—"}</span>
  return <div>
    <div style={{fontSize:10,color:C.text3,marginBottom:2}}>{ultimaMov}</div>
    <span style={{background:info.alerta?C.redSoft:C.amberSoft,color:info.alerta?C.red:C.amber,border:`1px solid ${info.alerta?C.redBorder:C.amberBorder}`,borderRadius:10,padding:"1px 7px",fontSize:10,fontWeight:600,whiteSpace:"nowrap"}}>
      {info.alerta?"⚠ ":""}{info.label}
    </span>
  </div>
}
  const info=timeOpen(sentAt);if(!info)return null
  return<span style={{background:info.alert?C.redSoft:C.amberSoft,color:info.alert?C.red:C.amber,border:`1px solid ${info.alert?C.redBorder:C.amberBorder}`,borderRadius:10,padding:"2px 8px",fontSize:10,fontWeight:500}}>{info.label}</span>
}

function KpiCard({label,val,sub,accent,trend}){
  return<div style={{background:C.white,borderRadius:12,padding:"20px 22px",border:`1px solid ${accent?C.redBorder:C.border}`,boxShadow:shadow.sm,position:"relative",overflow:"hidden"}}>
    {accent&&<div style={{position:"absolute",top:0,left:0,right:0,height:3,background:`linear-gradient(90deg,${C.red},#e74c3c88)`}}/>}
    {!accent&&<div style={{position:"absolute",top:0,left:0,right:0,height:2,background:`linear-gradient(90deg,${C.gold}44,${C.gold})`}}/>}
    <div style={{fontSize:9,color:C.text3,letterSpacing:"0.14em",textTransform:"uppercase",marginBottom:10,fontWeight:500}}>{label}</div>
    <div style={{fontSize:28,fontWeight:600,color:accent?C.red:C.brand,letterSpacing:"-0.02em",lineHeight:1,marginBottom:6,fontFamily:"'Cormorant Garamond',serif"}}>{val}</div>
    {sub&&<div style={{fontSize:11,color:C.text3,fontWeight:400}}>{sub}</div>}
  </div>
}

function CopyBtn({text,label}){
  const[ok,setOk]=useState(false)
  return<button onClick={()=>{navigator.clipboard.writeText(text);setOk(true);setTimeout(()=>setOk(false),2000)}}
    style={{background:ok?C.greenSoft:C.gold,border:`1px solid ${ok?C.greenBorder:C.goldDim}`,color:ok?C.green:C.white,borderRadius:6,padding:"5px 12px",fontSize:11,cursor:"pointer",fontWeight:500,whiteSpace:"nowrap",transition:"all .2s"}}>
    {ok?"✓ Copiado!":label||"Copiar"}
  </button>
}

function SortIcon({col,sortCol,sortDir}){
  if(sortCol!==col)return<span style={{color:C.text4,fontSize:9,marginLeft:4}}>⇅</span>
  return<span style={{color:C.gold,fontSize:9,marginLeft:4}}>{sortDir==="asc"?"↑":"↓"}</span>
}

function Toast({toasts}){
  return<div style={{position:"fixed",bottom:24,right:24,display:"flex",flexDirection:"column",gap:8,zIndex:9999,pointerEvents:"none"}}>
    {toasts.map(t=><div key={t.id} style={{background:t.type==="error"?C.red:t.type==="warn"?C.amber:C.green,color:C.white,borderRadius:10,padding:"12px 18px",fontSize:12,fontWeight:500,boxShadow:shadow.lg,maxWidth:320,lineHeight:1.5,letterSpacing:"0.01em"}}>{t.msg}</div>)}
  </div>
}

const INP = {borderRadius:8,border:`1px solid ${C.border}`,padding:"9px 12px",fontSize:12,background:C.white,color:C.text1,outline:"none",transition:"border-color .2s"}

// ─── Login ────────────────────────────────────────────────────
function LoginScreen({onLogin}){
  const[email,setEmail]=useState("")
  const[password,setPassword]=useState("")
  const[loading,setLoading]=useState(false)
  const[error,setError]=useState("")
  const handle=async e=>{
    e.preventDefault();setLoading(true);setError("")
    try{onLogin(await signIn(email,password))}
    catch(err){setError(err.message)}
    finally{setLoading(false)}
  }
  return<div style={{minHeight:"100vh",background:C.brand,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",position:"relative",overflow:"hidden"}}>
    {/* Decorative gold lines */}
    <div style={{position:"absolute",top:0,left:0,right:0,height:"1px",background:`linear-gradient(90deg,transparent,${C.gold},transparent)`}}/>
    <div style={{position:"absolute",bottom:0,left:0,right:0,height:"1px",background:`linear-gradient(90deg,transparent,${C.gold},transparent)`}}/>
    <div style={{position:"absolute",top:"50%",left:"50%",transform:"translate(-50%,-50%)",width:600,height:600,borderRadius:"50%",border:`1px solid ${C.gold}18`,pointerEvents:"none"}}/>
    <div style={{position:"absolute",top:"50%",left:"50%",transform:"translate(-50%,-50%)",width:400,height:400,borderRadius:"50%",border:`1px solid ${C.gold}12`,pointerEvents:"none"}}/>

    <div style={{textAlign:"center",marginBottom:48,position:"relative"}}>
      <div style={{fontSize:36,letterSpacing:"0.28em",color:C.white,textTransform:"uppercase",fontFamily:"'Cormorant Garamond',serif",fontWeight:500,lineHeight:1}}>Saint Germain</div>
      <div style={{width:48,height:"1px",background:C.gold,margin:"16px auto"}}/>
      <div style={{fontSize:9,letterSpacing:"0.35em",color:`${C.gold}99`,textTransform:"uppercase"}}>Central de Pedidos</div>
    </div>

    <div style={{background:C.white,borderRadius:16,padding:"40px 44px",width:"100%",maxWidth:400,boxShadow:shadow.lg,position:"relative"}}>
      <div style={{position:"absolute",top:0,left:"50%",transform:"translateX(-50%)",width:48,height:2,background:C.gold,borderRadius:1}}/>
      <form onSubmit={handle}>
        <div style={{marginBottom:20}}>
          <label style={{fontSize:9,letterSpacing:"0.14em",textTransform:"uppercase",color:C.text3,fontWeight:500,display:"block",marginBottom:7}}>E-mail</label>
          <input value={email} onChange={e=>setEmail(e.target.value)} type="email" placeholder="seu@email.com" required style={{...INP,width:"100%",boxSizing:"border-box",fontSize:13}}/>
        </div>
        <div style={{marginBottom:28}}>
          <label style={{fontSize:9,letterSpacing:"0.14em",textTransform:"uppercase",color:C.text3,fontWeight:500,display:"block",marginBottom:7}}>Senha</label>
          <input value={password} onChange={e=>setPassword(e.target.value)} type="password" placeholder="••••••••" required style={{...INP,width:"100%",boxSizing:"border-box",fontSize:13}}/>
        </div>
        {error&&<div style={{background:C.redSoft,color:C.red,border:`1px solid ${C.redBorder}`,borderRadius:8,padding:"10px 14px",fontSize:12,marginBottom:20,lineHeight:1.5}}>{error}</div>}
        <button type="submit" disabled={loading} style={{width:"100%",background:loading?C.text4:C.brand,border:"none",color:C.white,borderRadius:8,padding:"13px 0",fontSize:11,fontWeight:500,cursor:loading?"not-allowed":"pointer",letterSpacing:"0.18em",textTransform:"uppercase",transition:"background .2s"}}>
          {loading?"Autenticando...":"Entrar"}
        </button>
      </form>
    </div>

    <div style={{marginTop:32,fontSize:9,letterSpacing:"0.2em",color:`${C.white}33`,textTransform:"uppercase"}}>Acesso restrito · Uso interno</div>
  </div>
}

// ─── Painel de Usuários ───────────────────────────────────────
function UsuariosPanel({token,addToast}){
  const[usuarios,setUsuarios]=useState([])
  const[loading,setLoading]=useState(true)
  const[form,setForm]=useState({email:"",nome:"",perfil:"logistica",senha:""})
  const[saving,setSaving]=useState(false)
  const LABEL={admin:"Admin",logistica:"Logística",suporte:"Suporte",leitura:"Somente leitura"}
  useEffect(()=>{loadUsuarios(token).then(d=>{setUsuarios(d);setLoading(false)});},[])
  const handleCreate=async e=>{
    e.preventDefault();setSaving(true)
    try{
      const auth=await createUser(form.email,form.senha)
      await saveUsuario({id:auth.id,email:form.email,nome:form.nome,perfil:form.perfil,ativo:true},token)
      addToast(`Usuário ${form.email} criado!`)
      setForm({email:"",nome:"",perfil:"logistica",senha:""})
      setUsuarios(await loadUsuarios(token))
    }catch(err){addToast("Erro: "+err.message,"error")}
    finally{setSaving(false)}
  }
  const handleDelete=async id=>{
    if(!confirm("Remover este usuário?"))return
    await deleteUsuario(id,token);setUsuarios(u=>u.filter(x=>x.id!==id));addToast("Usuário removido","warn")
  }
  const handlePerfil=async(id,perfil)=>{
    const u=usuarios.find(x=>x.id===id)
    await saveUsuario({...u,perfil},token)
    setUsuarios(prev=>prev.map(x=>x.id===id?{...x,perfil}:x));addToast("Perfil atualizado")
  }
  return<div style={{padding:"32px 40px",maxWidth:860}}>
    <div style={{marginBottom:28}}>
      <div style={{fontSize:9,letterSpacing:"0.18em",textTransform:"uppercase",color:C.gold,marginBottom:6}}>Administração</div>
      <div style={{fontSize:22,fontWeight:500,color:C.text1,fontFamily:"'Cormorant Garamond',serif",letterSpacing:"0.03em"}}>Gestão de Usuários</div>
    </div>
    <div style={{background:C.white,borderRadius:14,border:`1px solid ${C.border}`,padding:28,marginBottom:24,boxShadow:shadow.sm}}>
      <div style={{fontSize:9,letterSpacing:"0.14em",textTransform:"uppercase",color:C.text3,marginBottom:20,fontWeight:500}}>Adicionar usuário</div>
      <form onSubmit={handleCreate}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:14}}>
          {[["Nome","nome","text","Nome completo"],["E-mail","email","email","email@exemplo.com"],["Senha inicial","senha","password","Mínimo 6 caracteres"]].map(([lbl,key,type,ph])=>(
            <div key={key}>
              <label style={{fontSize:9,letterSpacing:"0.12em",textTransform:"uppercase",color:C.text3,fontWeight:500,display:"block",marginBottom:6}}>{lbl}</label>
              <input value={form[key]} onChange={e=>setForm(f=>({...f,[key]:e.target.value}))} type={type} placeholder={ph} required={key!=="nome"} minLength={key==="senha"?6:undefined} style={{...INP,width:"100%",boxSizing:"border-box"}}/>
            </div>
          ))}
          <div>
            <label style={{fontSize:9,letterSpacing:"0.12em",textTransform:"uppercase",color:C.text3,fontWeight:500,display:"block",marginBottom:6}}>Perfil de acesso</label>
            <select value={form.perfil} onChange={e=>setForm(f=>({...f,perfil:e.target.value}))} style={{...INP,width:"100%",boxSizing:"border-box"}}>
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
              {["Nome","E-mail","Perfil","Ações"].map(h=><th key={h} style={{padding:"13px 18px",textAlign:"left",color:C.gold,fontWeight:400,fontSize:9,letterSpacing:"0.14em",textTransform:"uppercase"}}>{h}</th>)}
            </tr>
          </thead>
          <tbody>
            {usuarios.length===0?<tr><td colSpan={4} style={{padding:32,textAlign:"center",color:C.text4}}>Nenhum usuário cadastrado</td></tr>
            :usuarios.map((u,i)=>(
              <tr key={u.id} style={{background:i%2===0?C.white:C.cream,borderBottom:`1px solid ${C.border}`}}>
                <td style={{padding:"12px 18px",color:C.text1,fontWeight:500}}>{u.nome||"—"}</td>
                <td style={{padding:"12px 18px",color:C.text2}}>{u.email}</td>
                <td style={{padding:"12px 18px"}}>
                  <select value={u.perfil} onChange={e=>handlePerfil(u.id,e.target.value)} style={{...INP,padding:"5px 10px",fontSize:11}}>
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
}

const SAMPLE=`Identificador Ecommerce;Destinatário Nome;Estratégia de Frete;Rastreador Last Mile;Situação;Prazo Logístico;Nº Nota Fiscal
12345;Ana Souza;Correios PAC;AA123456789BR;Em trânsito;05/05/2026;98765
12346;Carlos Lima;Jadlog;JD987654321;Extraviado;28/04/2026;98766
12347;Mariana Costa;Total Express;TE112233445;Entregue;01/05/2026;98767
12348;Fernando Silva;Correios SEDEX;AA223344556BR;Saiu para entrega;02/05/2026;98768
12349;Júlia Martins;Loggi;LG556677889;Devolvido;30/04/2026;98769`

// ─── App ──────────────────────────────────────────────────────
export default function App(){
  const[session,setSession]=useState(null)
  const[perfil,setPerfil]=useState(null)
  const[nomeAtendente,setNomeAtendente]=useState("")
  const[loadingPerfil,setLoadingPerfil]=useState(false)
  const[rows,setRows]=useState([])
  const[tab,setTab]=useState(null)
  const[paste,setPaste]=useState("")
  const[importing,setImporting]=useState(false)
  const[loadingData,setLoadingData]=useState(false)
  const[compact,setCompact]=useState(false)
  const[toasts,setToasts]=useState([])
  const[lSrch,setLSrch]=useState("");const[lSt,setLSt]=useState("Todos")
  const[lTr,setLTr]=useState("Todos");const[lUrg,setLUrg]=useState("Todos")
  const[lAc,setLAc]=useState("Todos");const[qf,setQf]=useState("todos")
  const[lPage,setLPage]=useState(1)
  const[selIds,setSelIds]=useState(new Set())
  const[sortCol,setSortCol]=useState(null);const[sortDir,setSortDir]=useState("asc")
  const[sSrch,setSSrch]=useState("");const[sAtend,setSAtend]=useState("Todos")
  const[sUrg,setSUrg]=useState("Todos")
  const[selSup,setSelSup]=useState(null)
  const[selSupIds,setSelSupIds]=useState(new Set())
  const[openTpl,setOpenTpl]=useState(false);const[openHist,setOpenHist]=useState(false)
  const[aSrch,setASrch]=useState("")
  const[syncStatus,setSyncStatus]=useState("idle")
  const[lastSync,setLastSync]=useState(null)
  const[countdown,setCountdown]=useState(10)
  const[confirmClear,setConfirmClear]=useState(false)
  const saveTimer=useRef(null);const fileRef=useRef()
  const token=session?.access_token

  const addToast=useCallback((msg,type="ok",ms=4000)=>{
    const id=Date.now();setToasts(p=>[...p,{id,msg,type}]);setTimeout(()=>setToasts(p=>p.filter(t=>t.id!==id)),ms)
  },[])

  const handleLogin=async data=>{
    setSession(data);setLoadingPerfil(true)
    try{
      const r=await fetch(`${SUPA_URL}/rest/v1/usuarios?id=eq.${data.user.id}&select=*`,{headers:aSH(data.access_token)})
      const arr=await r.json();const p=arr[0]?.perfil||"leitura"
      setNomeAtendente(arr[0]?.nome||data.user?.email||"")
      setPerfil(p);setTab(PERMS[p].tabs[0])
    }catch(e){setPerfil("leitura");setTab("dashboard")}
    setLoadingPerfil(false)
  }
  const handleLogout=async()=>{await signOut(token);setSession(null);setPerfil(null);setRows([]);setTab(null)}
  const perms=perfil?PERMS[perfil]:null

  useEffect(()=>{
    if(!token)return
    setSyncStatus("loading")
    setLoadingData(true)

    const fixRows = data => data.map(r=>({
      ...r, isNew:false,
      atendimento: isEntregue(r.status)&&!r.enviadoSuporte ? "Resolvido" : r.atendimento,
      enviadoSuporte: isEntregue(r.status)&&!r.enviadoSuporte ? false : r.enviadoSuporte,
    }))

    dbLoadFast(token, (partial) => {
      // Mostra primeiros 1000 registros imediatamente
      setRows(fixRows(partial))
      setLastSync(new Date())
      setLoadingData(false) // Some com loading assim que tiver dados
    }).then(data => {
      if(data.length > 0) { setRows(fixRows(data)); setLastSync(new Date()) }
      setSyncStatus("idle")
      setLoadingData(false)
    }).catch(e => {
      setSyncStatus("error")
      addToast("Erro ao carregar: "+e.message,"error",8000)
      setLoadingData(false)
    })
  },[token])
  },[token])

  useEffect(()=>{
    if(!token)return
    const poll=async()=>{
      setCountdown(10)
      try{
        const remote=await dbLoad(token)
        if(remote.length>0){
          let nc=0
          setRows(prev=>{
            const rm=new Map(remote.map(r=>[r.id,r]));const lm=new Map(prev.map(r=>[r.id,r]))
            const merged=[...rm.values()].map(r=>{const loc=lm.get(r.id);if(!loc){nc++;return{...r,isNew:true}}return loc.historico.length>=r.historico.length?loc:{...r,isNew:false}})
            prev.forEach(r=>{if(!rm.has(r.id))merged.push(r)});return merged
          })
          if(nc>0)addToast(`${nc} pedido${nc>1?"s":""} atualizado${nc>1?"s":""} por outro usuário`,"warn")
          setLastSync(new Date())
        }
      }catch(e){}
    }
    const interval=setInterval(poll,10000);const cd=setInterval(()=>setCountdown(p=>p>0?p-1:10),1000)
    return()=>{clearInterval(interval);clearInterval(cd)}
  },[token,addToast])

  useEffect(()=>{
    if(!token||rows.length===0)return
    if(saveTimer.current)clearTimeout(saveTimer.current)
    saveTimer.current=setTimeout(async()=>{
      setSyncStatus("saving")
      try { await dbUpsert(rows,token) }
      catch(e){setSyncStatus("error");addToast("Erro ao salvar: "+e.message,"error",8000);setTimeout(()=>setSyncStatus("idle"),4000);return}
      setLastSync(new Date());setSyncStatus("saved");setTimeout(()=>setSyncStatus("idle"),2500)
    },1200)
  },[rows,token,addToast])

  useEffect(()=>{if(!rows.some(r=>r.isNew))return;const t=setTimeout(()=>setRows(p=>p.map(r=>({...r,isNew:false}))),6000);return()=>clearTimeout(t)},[rows])
  useEffect(()=>setLPage(1),[lSrch,lSt,lTr,lUrg,lAc,qf,sortCol,sortDir])
  useEffect(()=>{setOpenTpl(false);setOpenHist(false)},[selSup])

  const doImport=useCallback(txt=>{
    if(!perms?.canImport)return
    const parsed=parseData(txt)
    if(!parsed.length){addToast("Nenhum dado reconhecido.","error");return}
    let added=0,updated=0,skipped=0
    setRows(prev=>{
      const byNuvem=new Map(prev.map(r=>[r.nuvem,r]));const result=[...prev]
      for(const novo of parsed){
        const existing=byNuvem.get(novo.nuvem)
        if(!existing){result.push(novo);added++}
        else if(norm(existing.status)===norm(novo.status)){
          if(isEntregue(novo.status)&&!existing.enviadoSuporte&&existing.atendimento!=="Resolvido"){
            const idx=result.findIndex(r=>r.nuvem===novo.nuvem)
            if(idx>=0){result[idx]={...existing,atendimento:"Resolvido",enviadoSuporte:false,historico:[...existing.historico,{acao:"Arquivado automaticamente — entrega concluída",ts:new Date().toLocaleString("pt-BR")}]};updated++}
          }else{skipped++}
        }else{
          const idx=result.findIndex(r=>r.nuvem===novo.nuvem)
          if(idx>=0){
            const alertaStatus=existing.enviadoSuporte&&norm(existing.status)!==norm(novo.status)?`Status atualizado: ${existing.status} → ${novo.status}`:existing.alertaStatus
            result[idx]={...novo,id:existing.id,obs:existing.obs,responsavel:existing.responsavel,chamado:existing.chamado,enviadoSuporte:existing.enviadoSuporte,atendimento:existing.enviadoSuporte?existing.atendimento:novo.atendimento,alertaStatus,historico:[...existing.historico,{acao:`Status atualizado: ${existing.status} → ${novo.status}`,ts:new Date().toLocaleString("pt-BR")}],isNew:true};updated++
          }
        }
      }
      return result
    })
    setTimeout(()=>{
      const parts=[]
      if(added>0)parts.push(`${added} novo${added>1?"s":""}`)
      if(updated>0)parts.push(`${updated} atualizado${updated>1?"s":""}`)
      if(skipped>0)parts.push(`${skipped} ignorado${skipped>1?"s":""} (mesmo status)`)
      addToast(parts.join(" · ")||"Nenhuma alteração")
    },100)
    setPaste("");setImporting(false)
  },[addToast,perms])

  const handleFile=e=>{
    if(!perms?.canImport)return
    const f=e.target.files[0];if(!f)return
    if(/\.(xlsx?)$/i.test(f.name)){const rd=new FileReader();rd.onload=ev=>{const wb=XLSX.read(new Uint8Array(ev.target.result),{type:"array"});doImport(XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]],{FS:";",blankrows:false}))};rd.readAsArrayBuffer(f)}
    else{const rd=new FileReader();rd.onload=ev=>doImport(ev.target.result);rd.readAsText(f,/\.csv$/i.test(f.name)?"windows-1252":"UTF-8")}
    e.target.value=""
  }

  const upd=(id,ch,hist)=>setRows(prev=>prev.map(r=>{if(r.id!==id)return r;const historico=hist?[...r.historico,{...hist,ts:new Date().toLocaleString("pt-BR")}]:r.historico;return{...r,...ch,historico}}))
  const del=id=>{if(!perms?.canDelete)return;setRows(prev=>prev.filter(r=>r.id!==id));dbDelete(id,token).catch(()=>{})}
  const toggleSel=id=>setSelIds(prev=>{const n=new Set(prev);n.has(id)?n.delete(id):n.add(id);return n})
  const clearSel=()=>setSelIds(new Set())
  const bulkSend=()=>{
    if(!perms?.canSendSupport)return
    const ts=new Date().toLocaleString("pt-BR"),sentAt=new Date().toISOString()
    setRows(prev=>prev.map(r=>selIds.has(r.id)?{...r,enviadoSuporte:true,atendimento:"Aberto",sentAt,historico:[...r.historico,{acao:"Enviado ao suporte (lote)",ts}]}:r))
    addToast(`${selIds.size} pedido${selIds.size>1?"s":""} enviado${selIds.size>1?"s":""} ao suporte`);clearSel()
  }
  const toggleSelSup=id=>setSelSupIds(prev=>{const n=new Set(prev);n.has(id)?n.delete(id):n.add(id);return n})
  const bulkArchive=()=>{
    if(!perms?.canOperate)return
    const ts=new Date().toLocaleString("pt-BR")
    setRows(prev=>prev.map(r=>selSupIds.has(r.id)?{...r,atendimento:"Resolvido",historico:[...r.historico,{acao:"Arquivado em lote",ts}]}:r))
    addToast(`${selSupIds.size} pedido${selSupIds.size>1?"s":""} arquivado${selSupIds.size>1?"s":""}`)
    if(selSupIds.has(selSup))setSelSup(null);setSelSupIds(new Set())
  }
  const handleInitiate=id=>{if(!perms?.canOperate)return;upd(id,{atendimento:"Em andamento"},{acao:"Atendimento iniciado"});setSelSup(id)}
  const handleResolve=id=>{if(!perms?.canOperate)return;upd(id,{atendimento:"Resolvido"},{acao:"Status → Resolvido"});setSelSup(null);addToast("Pedido resolvido e arquivado")}
  const handleReturnLog=id=>{if(!perms?.canOperate)return;upd(id,{enviadoSuporte:false,sentAt:null},{acao:"Devolvido à Logística"});setSelSup(null)}
  const handleClearAll=()=>{
    if(!perms?.canClear)return
    if(!confirmClear){setConfirmClear(true);setTimeout(()=>setConfirmClear(false),4000);return}
    setRows([]);setConfirmClear(false);dbClear(token).catch(()=>{});addToast("Todos os dados foram removidos","warn")
  }
  const toggleSort=col=>{if(sortCol===col)setSortDir(d=>d==="asc"?"desc":"asc");else{setSortCol(col);setSortDir("asc")}}

  const baseLog =rows.filter(r=>!r.enviadoSuporte&&r.atendimento!=="Resolvido")
  const baseSup =rows.filter(r=>r.enviadoSuporte&&r.atendimento!=="Resolvido")
  const baseArch=rows.filter(r=>r.atendimento==="Resolvido")
  const detail  =selSup?baseSup.find(r=>r.id===selSup):null
  const qCounts =Object.fromEntries(QFILTERS.map(f=>[f.id,applyQF(baseLog,f.id).length]))
  const filteredLog=applySortRows(applyQF(baseLog,qf).filter(r=>{
    const q=lSrch.toLowerCase()
    return(!q||[r.nuvem,r.destinatario,r.transportadora,r.rastreio,r.status,r.nf,r.motivo].some(v=>v.toLowerCase().includes(q)))
      &&(lSt==="Todos"||r.status===lSt)&&(lTr==="Todos"||r.transportadora===lTr)
      &&(lUrg==="Todos"||r.urgencia===lUrg)&&(lAc==="Todos"||r.acionar===lAc)
  }),sortCol,sortDir)
  const totalPages=Math.max(1,Math.ceil(filteredLog.length/PAGE_SIZE))
  const safeP=Math.min(lPage,totalPages)
  const pagedLog=filteredLog.slice((safeP-1)*PAGE_SIZE,safeP*PAGE_SIZE)
  const supRows=baseSup.filter(r=>{const q=sSrch.toLowerCase();return(!q||[r.nuvem,r.destinatario,r.rastreio,r.nf,r.status].some(v=>v.toLowerCase().includes(q)))&&(sAtend==="Todos"||r.atendimento===sAtend)&&(sUrg==="Todos"||r.urgencia===sUrg)}).sort((a,b)=>{const uo={Alta:0,Média:1,Baixa:2,"—":3},ao={Aberto:0,"Em andamento":1};return(uo[a.urgencia]-uo[b.urgencia])||(ao[a.atendimento]-ao[b.atendimento])})
  const archRows=baseArch.filter(r=>{const q=aSrch.toLowerCase();return!q||[r.nuvem,r.destinatario,r.transportadora,r.status,r.nf].some(v=>v.toLowerCase().includes(q))}).sort((a,b)=>{const ta=(a.historico.find(h=>h.acao&&(h.acao.includes("Resolvido")||h.acao.includes("Arquivado")))||{}).ts||"";const tb=(b.historico.find(h=>h.acao&&(h.acao.includes("Resolvido")||h.acao.includes("Arquivado")))||{}).ts||"";return tb.localeCompare(ta)})

  const stOpts=uniq(baseLog.map(r=>r.status)),trOpts=uniq(baseLog.map(r=>r.transportadora))
  const st={log:baseLog.length,alta:baseLog.filter(r=>r.urgencia==="Alta").length,acionar:baseLog.filter(r=>r.acionar==="Sim").length}
  const ss={total:baseSup.length,abertos:baseSup.filter(r=>r.atendimento==="Aberto").length,andamento:baseSup.filter(r=>r.atendimento==="Em andamento").length}
  const arch=baseArch.length

  const entregues=rows.filter(r=>isEntregue(r.status))
  const noPrazo=entregues.filter(r=>r.entregueNoPrazo===true).length
  const foraPrazo=entregues.filter(r=>r.entregueNoPrazo===false).length
  const pctNoPrazo=entregues.length>0?Math.round((noPrazo/entregues.length)*100):0
  const hoje=new Date();hoje.setHours(0,0,0,0)
  const vencidos=rows.filter(r=>{const d=parsePrazo(r.prazo);if(!d)return false;return d<hoje&&!isEntregue(r.status)}).length
  const emRisco=rows.filter(r=>{const d=parsePrazo(r.prazo);if(!d)return false;const diff=Math.ceil((d-hoje)/86400000);return diff>=0&&diff<=3&&!isEntregue(r.status)}).length

  const parados = baseLog.filter(r=>{ const d=diasSemMov(r.ultimaMov); return d!==null&&d>=ALERTA_DIAS }).length
  rows.forEach(r=>{
    if(!r.transportadora)return
    if(!trStats[r.transportadora])trStats[r.transportadora]={total:0,entregues:0,noPrazo:0,foraPrazo:0,vencidos:0}
    const s=trStats[r.transportadora];s.total++
    if(isEntregue(r.status)){s.entregues++;if(r.entregueNoPrazo===true)s.noPrazo++;if(r.entregueNoPrazo===false)s.foraPrazo++}
    else{const d=parsePrazo(r.prazo);if(d&&d<hoje)s.vencidos++}
  })
  const trData=Object.entries(trStats).map(([name,s])=>({name,total:s.total,entregues:s.entregues,noPrazo:s.noPrazo,foraPrazo:s.foraPrazo,vencidos:s.vencidos,pct:s.entregues>0?Math.round((s.noPrazo/s.entregues)*100):0})).sort((a,b)=>b.total-a.total).slice(0,8)
  const trBarData=trData.map(t=>({name:t.name,"No prazo":t.noPrazo,"Fora prazo":t.foraPrazo,"Vencidos":t.vencidos}))

  // Stats por UF
  const ufStats={}
  rows.forEach(r=>{
    const uf=(r.uf||"").toUpperCase().trim(); if(!uf||uf.length>3)return
    if(!ufStats[uf])ufStats[uf]={total:0,entregues:0,noPrazo:0,tempoTotal:0,tempoCount:0}
    const s=ufStats[uf]; s.total++
    if(isEntregue(r.status)){s.entregues++; if(r.entregueNoPrazo===true)s.noPrazo++}
  })
  const ufData=Object.entries(ufStats).map(([uf,s])=>({uf,total:s.total,entregues:s.entregues,noPrazo:s.noPrazo,pct:s.entregues>0?Math.round((s.noPrazo/s.entregues)*100):0})).sort((a,b)=>b.total-a.total).slice(0,15)
  const urgData=["Alta","Média","Baixa"].map(u=>({name:u,value:baseLog.filter(r=>r.urgencia===u).length,fill:urgStyles[u].dot})).filter(d=>d.value>0)
  const statusMap={};rows.filter(r=>!isEntregue(r.status)).forEach(r=>{if(r.status)statusMap[r.status]=(statusMap[r.status]||0)+1})
  const statusData=Object.entries(statusMap).sort((a,b)=>b[1]-a[1]).slice(0,8).map(([name,value])=>({name,value}))

  if(!session)return<LoginScreen onLogin={handleLogin}/>
  if(loadingPerfil)return<div style={{display:"flex",alignItems:"center",justifyContent:"center",minHeight:"100vh",color:C.text4,fontSize:12,letterSpacing:"0.1em"}}>Carregando perfil...</div>

  const showImp=(importing||rows.length===0)&&perms?.tabs.some(t=>["logistica","dashboard"].includes(t))
  const pd=compact?5:9
  const PERFLABEL={admin:"Admin",logistica:"Logística",suporte:"Suporte",leitura:"Leitura"}
  const syncDot=syncStatus==="error"?C.red:syncStatus==="saving"?C.gold:syncStatus==="saved"?"#27ae60":"#555"
  const syncText=syncStatus==="loading"?"Carregando...":syncStatus==="saving"?"Salvando...":syncStatus==="saved"?"Sincronizado ✓":syncStatus==="error"?"Erro":lastSync?`Sync em ${countdown}s`:""

  const TABS=[{key:"dashboard",label:"Dashboard",badge:null},{key:"logistica",label:"Logística",badge:st.acionar>0?st.acionar:null},{key:"suporte",label:"Suporte",badge:ss.abertos>0?ss.abertos:null},{key:"arquivados",label:"Arquivados",badge:arch>0?arch:null},{key:"usuarios",label:"Usuários",badge:null}].filter(t=>perms?.tabs.includes(t.key))

  const TH={padding:`${compact?8:11}px 14px`,textAlign:"left",color:C.gold,fontWeight:400,fontSize:9,letterSpacing:"0.14em",textTransform:"uppercase",borderBottom:`1px solid #2A2A2A`,whiteSpace:"nowrap",background:C.brand,position:"sticky",top:0,zIndex:5,cursor:"pointer"}
  const THF={...TH,cursor:"default"}

  return<div style={{fontFamily:"'Inter',sans-serif",minHeight:"100vh",background:C.cream}}>
    <style>{globalStyle}</style>
    <Toast toasts={toasts}/>

    {/* ── HEADER ── */}
    <div style={{background:C.brand,padding:"0 32px",display:"flex",alignItems:"stretch",justifyContent:"space-between",borderBottom:`1px solid ${C.gold}33`}}>
      <div style={{display:"flex",alignItems:"center",gap:24,padding:"14px 0"}}>
        <div>
          <div style={{fontSize:10,letterSpacing:"0.45em",color:C.gold,textTransform:"uppercase",fontWeight:300,lineHeight:1}}>Saint Germain</div>
          <div style={{fontSize:8,letterSpacing:"0.28em",color:`${C.gold}55`,textTransform:"uppercase",marginTop:3}}>Central de Pedidos</div>
        </div>
        <div style={{width:1,height:32,background:`${C.gold}22`}}/>
        {/* Sync */}
        <div style={{display:"flex",alignItems:"center",gap:7}}>
          <span style={{width:6,height:6,borderRadius:"50%",background:syncDot,display:"inline-block",boxShadow:`0 0 6px ${syncDot}66`}}/>
          <span style={{fontSize:10,color:syncDot,letterSpacing:"0.04em"}}>{syncText}</span>
        </div>
      </div>

      {/* Ações */}
      <div style={{display:"flex",alignItems:"center",gap:8}}>
        {perms?.canImport&&!showImp&&rows.length>0&&(
          <>
            <button onClick={()=>setCompact(c=>!c)} style={{background:"transparent",border:`1px solid ${C.gold}44`,color:compact?C.gold:`${C.gold}88`,borderRadius:6,padding:"6px 12px",fontSize:10,cursor:"pointer",letterSpacing:"0.08em"}}>
              {compact?"⊞":"⊟"}
            </button>
            <button onClick={()=>{const hf=lSrch||lSt!=="Todos"||lTr!=="Todos"||lUrg!=="Todos"||lAc!=="Todos";exportCSV(hf&&tab==="logistica"?filteredLog:rows)}} style={{background:"transparent",border:`1px solid #444`,color:"#888",borderRadius:6,padding:"6px 14px",fontSize:10,cursor:"pointer",letterSpacing:"0.06em"}}>↓ Exportar</button>
            <button onClick={()=>setImporting(true)} style={{background:"transparent",border:`1px solid ${C.gold}`,color:C.gold,borderRadius:6,padding:"6px 14px",fontSize:10,cursor:"pointer",letterSpacing:"0.06em"}}>+ Importar</button>
          </>
        )}
        {perms?.canClear&&rows.length>0&&(
          <button onClick={handleClearAll} style={{background:confirmClear?C.red:"transparent",border:`1px solid ${confirmClear?C.red:"#333"}`,color:confirmClear?C.white:"#555",borderRadius:6,padding:"6px 14px",fontSize:10,cursor:"pointer",letterSpacing:"0.06em"}}>
            {confirmClear?"⚠ Confirmar":"Limpar"}
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
      <div style={{background:C.white,borderBottom:`1px solid ${C.border}`,padding:"0 32px",display:"flex",alignItems:"stretch",boxShadow:"0 1px 0 rgba(0,0,0,0.04)"}}>
        {TABS.map(t=>(
          <button key={t.key} onClick={()=>{setTab(t.key);if(t.key!=="suporte")setSelSup(null)}}
            style={{background:"transparent",border:"none",borderBottom:tab===t.key?`2px solid ${C.gold}`:"2px solid transparent",color:tab===t.key?C.text1:C.text3,padding:"14px 20px",cursor:"pointer",fontSize:10,letterSpacing:"0.14em",textTransform:"uppercase",fontWeight:tab===t.key?500:400,marginBottom:"-1px",display:"flex",alignItems:"center",gap:8,transition:"color .2s"}}>
            {t.label}
            {t.badge!=null&&<span style={{background:tab===t.key?C.brand:C.red,color:C.white,borderRadius:10,padding:"2px 7px",fontSize:9,fontWeight:600,letterSpacing:"0.04em"}}>{t.badge}</span>}
          </button>
        ))}
      </div>
    )}

    {/* ── IMPORT ── */}
    {showImp&&perms?.canImport&&(
      <div style={{padding:48,maxWidth:640,margin:"0 auto"}}>
        <div style={{textAlign:"center",marginBottom:36}}>
          <div style={{fontSize:9,letterSpacing:"0.22em",textTransform:"uppercase",color:C.gold,marginBottom:10}}>{importing?"Adicionar dados":"Bem-vindo"}</div>
          <div style={{fontSize:28,fontWeight:500,color:C.text1,fontFamily:"'Cormorant Garamond',serif",letterSpacing:"0.04em",marginBottom:8}}>Importe seus dados</div>
          <div style={{color:C.text3,fontSize:12,lineHeight:1.6}}>Aceita .csv (cp1252), .xls, .xlsx ou colagem direta do Excel</div>
        </div>
        <div style={{background:C.white,borderRadius:14,border:`1px solid ${C.border}`,padding:24,marginBottom:16,boxShadow:shadow.sm}}>
          <div style={{fontSize:9,color:C.text3,marginBottom:14,fontWeight:500,letterSpacing:"0.12em",textTransform:"uppercase"}}>Mapeamento automático de colunas</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
            {[["No NUVEM","Identificador Ecommerce"],["Destinatário","Destinatário Nome"],["Transportadora","Estratégia de Frete"],["Cód. Rastreio","Rastreador Last Mile"],["Status","Situação"],["Prazo","Prazo Logístico"]].map(([c,a])=>(
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
    {tab==="usuarios"&&perfil==="admin"&&<UsuariosPanel token={token} addToast={addToast}/>}

    {/* ── DASHBOARD ── */}
    {tab==="dashboard"&&!showImp&&(
      <div style={{padding:"28px 40px"}}>
        <div style={{marginBottom:24}}>
          <div style={{fontSize:9,letterSpacing:"0.18em",textTransform:"uppercase",color:C.gold,marginBottom:4}}>Visão geral</div>
          <div style={{fontSize:22,fontWeight:500,color:C.text1,fontFamily:"'Cormorant Garamond',serif",letterSpacing:"0.03em"}}>Dashboard Operacional</div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:12,marginBottom:20}}>
          {[
            {label:"Total de pedidos",      val:rows.length,  sub:"na base de dados"},
            {label:"Em logística",          val:st.log,       sub:`${st.acionar} acionam suporte`},
            {label:"No suporte",            val:ss.total,     sub:`${ss.abertos} abertos`, accent:ss.abertos>0},
            {label:`Parados +${ALERTA_DIAS}d`, val:parados,   sub:"sem movimentação",      accent:parados>0},
            {label:"Entrega no prazo",      val:`${pctNoPrazo}%`, sub:`${noPrazo} de ${entregues.length} entregues`},
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
                    <Bar dataKey="No prazo"   stackId="a" fill={C.green}  name="No prazo"/>
                    <Bar dataKey="Fora prazo" stackId="a" fill={C.amber}  name="Fora prazo"/>
                    <Bar dataKey="Vencidos"   stackId="a" fill={C.red}    name="Vencidos" radius={[0,4,4,0]}/>
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
            {/* Ranking UF */}
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
                        <span style={{fontSize:14,fontWeight:700,color:C.text1,fontFamily:"'Cormorant Garamond',serif",letterSpacing:"0.05em"}}>{u.uf}</span>
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
            {/* Ranking transportadoras */}
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
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12,marginBottom:20}}>
          <KpiCard label="Em logística"   val={st.log}/>
          <KpiCard label="Urgência alta"  val={st.alta}  accent={st.alta>0}/>
          <KpiCard label="Acionar suporte" val={st.acionar} accent={st.acionar>0}/>
        </div>
        {/* Quick filters */}
        <div style={{display:"flex",gap:6,marginBottom:14,flexWrap:"wrap"}}>
          {QFILTERS.map(f=>(
            <button key={f.id} onClick={()=>{setQf(f.id);clearSel()}}
              style={{background:qf===f.id?C.brand:C.white,border:`1px solid ${qf===f.id?C.brand:C.border}`,color:qf===f.id?C.white:C.text2,borderRadius:20,padding:"5px 14px",fontSize:10,cursor:"pointer",fontWeight:qf===f.id?500:400,letterSpacing:"0.06em",boxShadow:qf===f.id?"none":shadow.sm,transition:"all .2s"}}>
              {f.label}{f.id!=="todos"?` (${qCounts[f.id]||0})`:""}
            </button>
          ))}
        </div>
        {/* Filters */}
        <div style={{background:C.white,borderRadius:10,border:`1px solid ${C.border}`,padding:"12px 16px",marginBottom:14,display:"flex",gap:8,flexWrap:"wrap",alignItems:"center",boxShadow:shadow.sm}}>
          <input value={lSrch} onChange={e=>setLSrch(e.target.value)} placeholder="Buscar pedido, destinatário, rastreio..." style={{...INP,flex:1,minWidth:160}}/>
          <select value={lSt}  onChange={e=>setLSt(e.target.value)}  style={INP}>{stOpts.map(o=><option key={o}>{o}</option>)}</select>
          <select value={lTr}  onChange={e=>setLTr(e.target.value)}  style={INP}>{trOpts.map(o=><option key={o}>{o}</option>)}</select>
          <select value={lUrg} onChange={e=>setLUrg(e.target.value)} style={INP}>{["Todos","Alta","Média","Baixa","—"].map(o=><option key={o}>{o}</option>)}</select>
          <select value={lAc}  onChange={e=>setLAc(e.target.value)}  style={INP}>{["Todos","Sim","Avaliar","Não"].map(o=><option key={o}>{o}</option>)}</select>
          {(lSrch||lSt!=="Todos"||lTr!=="Todos"||lUrg!=="Todos"||lAc!=="Todos")&&<button onClick={()=>{setLSrch("");setLSt("Todos");setLTr("Todos");setLUrg("Todos");setLAc("Todos")}} style={{...INP,cursor:"pointer",color:C.red,borderColor:C.redBorder,background:C.redSoft}}>× Limpar</button>}
        </div>
        {/* Bulk bar */}
        {perms?.canSendSupport&&selIds.size>0&&(
          <div style={{background:C.brand,borderRadius:10,padding:"12px 20px",marginBottom:14,display:"flex",alignItems:"center",gap:14,boxShadow:shadow.md}}>
            <span style={{color:"#888",fontSize:12,flex:1}}>{selIds.size} pedido{selIds.size>1?"s":""} selecionado{selIds.size>1?"s":""}</span>
            <button onClick={bulkSend} style={{background:C.gold,border:"none",color:C.white,borderRadius:7,padding:"8px 20px",fontSize:11,cursor:"pointer",fontWeight:500,letterSpacing:"0.08em"}}>Enviar ao Suporte ({selIds.size})</button>
            <button onClick={clearSel} style={{background:"transparent",border:`1px solid #333`,color:"#666",borderRadius:7,padding:"8px 14px",fontSize:11,cursor:"pointer"}}>Cancelar</button>
          </div>
        )}
        {/* Table */}
        <div style={{overflowX:"auto",overflowY:"auto",maxHeight:"54vh",borderRadius:12,border:`1px solid ${C.border}`,boxShadow:shadow.sm}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:compact?11:12,tableLayout:"fixed",minWidth:1040}}>
                          <colgroup><col style={{width:36}}/><col style={{width:76}}/><col style={{width:114}}/><col style={{width:94}}/><col style={{width:104}}/><col style={{width:94}}/><col style={{width:100}}/><col style={{width:64}}/><col style={{width:108}}/><col style={{width:62}}/><col style={{width:72}}/><col style={{width:106}}/><col style={{width:94}}/><col style={{width:36}}/></colgroup>
            <thead>
              <tr>
                <th style={THF}>{perms?.canSendSupport&&<input type="checkbox" onChange={e=>e.target.checked?setSelIds(new Set(pagedLog.map(r=>r.id))):clearSel()} checked={selIds.size>0&&pagedLog.every(r=>selIds.has(r.id))} style={{cursor:"pointer",accentColor:C.gold}}/>}</th>
                {[["nuvem","No NUVEM"],["destinatario","Destinatário"],["transportadora","Transportadora"],["rastreio","Cód. Rastreio"],["status","Status"],["prazo","Prazo / SLA"],["nf","No NF"],["motivo","Motivo (auto)"],["urgencia","Urgência"],["acionar","Acionar?"]].map(([col,label])=>(
                  <th key={col} onClick={()=>toggleSort(col)} style={TH}>{label}<SortIcon col={col} sortCol={sortCol} sortDir={sortDir}/></th>
                ))}
                <th style={THF}>Ação</th><th style={THF}/>
              </tr>
            </thead>
            <tbody>
              {pagedLog.length===0?<tr><td colSpan={13} style={{textAlign:"center",padding:36,color:C.text4}}>Nenhum pedido encontrado</td></tr>
              :pagedLog.map((r,i)=>(
                <tr key={r.id} style={{background:r.isNew?`${C.gold}14`:i%2===0?C.white:C.cream,borderBottom:`1px solid ${C.border}66`,outline:r.isNew?`1px solid ${C.gold}44`:"none"}}>
                  <td style={{padding:`${pd}px 8px`,textAlign:"center"}}>{perms?.canSendSupport&&<input type="checkbox" checked={selIds.has(r.id)} onChange={()=>toggleSel(r.id)} style={{cursor:"pointer",accentColor:C.gold}}/>}</td>
                  <td style={{padding:`${pd}px 14px`,fontWeight:600,color:C.text1,fontSize:11}}>{r.nuvem}</td>
                  <td style={{padding:`${pd}px 14px`,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:C.text1}} title={r.destinatario}>{r.destinatario}</td>
                  <td style={{padding:`${pd}px 14px`,color:C.text2,overflow:"hidden",textOverflow:"ellipsis"}}>{r.transportadora}</td>
                  <td style={{padding:`${pd}px 14px`,color:C.text3,fontFamily:"monospace",fontSize:10,overflow:"hidden",textOverflow:"ellipsis"}}>{r.rastreio}</td>
                  <td style={{padding:`${pd}px 14px`}}><StatusBadge val={r.status}/></td>
                  <td style={{padding:`${pd}px 14px`}}><SlaCell prazo={r.prazo}/></td>
                  <td style={{padding:`${pd}px 14px`,color:C.text2,fontSize:11}}>{r.nf}</td>
                  <td style={{padding:`${pd}px 14px`,color:C.text3,fontSize:10,overflow:"hidden",textOverflow:"ellipsis"}} title={r.motivo}>{r.motivo}</td>
                  <td style={{padding:`${pd}px 14px`}}><Chip val={r.urgencia} styles={urgStyles}/></td>
                  <td style={{padding:`${pd}px 14px`}}><Chip val={r.acionar} styles={acionStyles}/></td>
                  <td style={{padding:`${pd}px 14px`}}><SemMovBadge ultimaMov={r.ultimaMov}/></td>
                  <td style={{padding:`${pd}px 14px`}}>{perms?.canSendSupport&&(
                    <button onClick={()=>upd(r.id,{enviadoSuporte:true,atendimento:"Aberto",sentAt:new Date().toISOString()},{acao:"Enviado ao suporte"})} style={{background:C.cream,border:`1px solid ${C.border}`,color:C.text2,borderRadius:6,padding:"4px 10px",fontSize:10,cursor:"pointer",width:"100%",fontWeight:500,transition:"all .15s"}}>
                      Enviar →
                    </button>
                  )}</td>
                  <td style={{padding:`${pd}px 8px`,textAlign:"center"}}>{perms?.canDelete&&<button onClick={()=>del(r.id)} style={{background:"transparent",border:"none",color:C.text4,cursor:"pointer",fontSize:14}}>×</button>}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginTop:12}}>
          <div style={{fontSize:11,color:C.text4}}>{filteredLog.length===0?"Nenhum resultado":`Mostrando ${((safeP-1)*PAGE_SIZE)+1}–${Math.min(safeP*PAGE_SIZE,filteredLog.length)} de ${filteredLog.length} pedidos`}</div>
          {totalPages>1&&<div style={{display:"flex",gap:4,alignItems:"center"}}>
            <button onClick={()=>setLPage(n=>Math.max(1,n-1))} disabled={safeP===1} style={{...INP,padding:"5px 12px",cursor:safeP===1?"not-allowed":"pointer",opacity:safeP===1?0.4:1}}>‹</button>
            <span style={{fontSize:11,color:C.text3,padding:"0 10px"}}>{safeP} / {totalPages}</span>
            <button onClick={()=>setLPage(n=>Math.min(totalPages,n+1))} disabled={safeP===totalPages} style={{...INP,padding:"5px 12px",cursor:safeP===totalPages?"not-allowed":"pointer",opacity:safeP===totalPages?0.4:1}}>›</button>
          </div>}
        </div>
      </div>
    )}

    {/* ── SUPORTE ── */}
    {tab==="suporte"&&(
      <div style={{display:"flex",minHeight:"calc(100vh - 110px)"}}>
        {/* Fila */}
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
            <div style={{display:"flex",gap:6,marginBottom:selSupIds.size>0?10:0}}>
              <input value={sSrch} onChange={e=>setSSrch(e.target.value)} placeholder="Buscar..." style={{...INP,flex:1,fontSize:11}}/>
              <select value={sAtend} onChange={e=>setSAtend(e.target.value)} style={{...INP,fontSize:11}}>{["Todos","Aberto","Em andamento"].map(o=><option key={o}>{o}</option>)}</select>
              <select value={sUrg}   onChange={e=>setSUrg(e.target.value)}   style={{...INP,fontSize:11}}>{["Todos","Alta","Média","Baixa"].map(o=><option key={o}>{o}</option>)}</select>
            </div>
            {perms?.canOperate&&selSupIds.size>0&&(
              <div style={{background:C.brand,borderRadius:8,padding:"9px 14px",display:"flex",alignItems:"center",gap:8}}>
                <span style={{color:"#888",fontSize:11,flex:1}}>{selSupIds.size} selecionado{selSupIds.size>1?"s":""}</span>
                <button onClick={bulkArchive} style={{background:C.gold,border:"none",color:C.white,borderRadius:6,padding:"5px 12px",fontSize:10,cursor:"pointer",fontWeight:500}}>Arquivar ({selSupIds.size})</button>
                <button onClick={()=>setSelSupIds(new Set())} style={{background:"transparent",border:"none",color:"#666",fontSize:12,cursor:"pointer"}}>✕</button>
              </div>
            )}
          </div>
          <div style={{overflowY:"auto",flex:1}}>
            {ss.total===0?<div style={{textAlign:"center",padding:"56px 20px",color:C.text4}}><div style={{fontSize:32,marginBottom:12,opacity:0.2}}>◎</div><div style={{fontSize:13}}>Fila vazia</div><div style={{fontSize:11,marginTop:4}}>Pedidos enviados da Logística aparecem aqui</div></div>
            :supRows.length===0?<div style={{textAlign:"center",padding:24,color:C.text4,fontSize:12}}>Nenhum resultado</div>
            :supRows.map(r=>{
              const isSel=selSup===r.id
              const acColor=r.urgencia==="Alta"?C.red:r.urgencia==="Média"?C.gold:C.green
              return<div key={r.id} onClick={()=>setSelSup(isSel?null:r.id)}
                style={{padding:"13px 16px 13px 14px",cursor:"pointer",borderBottom:`1px solid ${C.border}55`,borderLeft:`3px solid ${acColor}`,background:isSel?C.amberSoft:r.alertaStatus?`${C.amber}08`:C.white,transition:"background .15s",display:"flex",alignItems:"flex-start",gap:10}}>
                {perms?.canOperate&&<input type="checkbox" checked={selSupIds.has(r.id)} onClick={e=>e.stopPropagation()} onChange={()=>toggleSelSup(r.id)} style={{marginTop:3,cursor:"pointer",accentColor:C.gold,flexShrink:0}}/>}
                <div style={{flex:1,minWidth:0}}>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:4}}>
                    <span style={{fontWeight:600,fontSize:12,color:C.text1}}>{r.nuvem}</span>
                    <TimeOpenBadge sentAt={r.sentAt}/>
                  </div>
                  {r.alertaStatus&&<div style={{background:C.amberSoft,border:`1px solid ${C.amberBorder}`,borderRadius:5,padding:"3px 8px",fontSize:10,color:C.amber,marginBottom:5,display:"flex",alignItems:"center",gap:4}}><span>⚠</span><span style={{fontWeight:500}}>Status alterado — notifique!</span></div>}
                  <div style={{fontSize:11,color:C.text2,marginBottom:6,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.destinatario}</div>
                  <div style={{display:"flex",gap:4,flexWrap:"wrap"}}><StatusBadge val={r.status}/><Chip val={r.urgencia} styles={urgStyles}/><Chip val={r.atendimento} styles={atendStyles}/></div>
                  {r.responsavel&&<div style={{fontSize:9,color:C.text4,marginTop:6,letterSpacing:"0.06em"}}>RESP: {r.responsavel}</div>}
                </div>
              </div>
            })}
          </div>
        </div>

        {/* Painel detalhe */}
        {detail?(
          <div style={{flex:1,display:"flex",flexDirection:"column",overflowY:"auto",background:C.cream}}>
            {/* Header sticky */}
            <div style={{background:C.white,padding:"20px 28px 16px",borderBottom:`1px solid ${C.border}`,position:"sticky",top:0,zIndex:5,boxShadow:shadow.sm}}>
              <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:12}}>
                <div>
                  <div style={{fontSize:9,color:C.gold,letterSpacing:"0.16em",textTransform:"uppercase",marginBottom:5}}>Pedido em atendimento</div>
                  <div style={{fontSize:20,fontWeight:500,color:C.text1,fontFamily:"'Cormorant Garamond',serif",marginBottom:2,letterSpacing:"0.02em"}}>{detail.destinatario}</div>
                  <div style={{fontSize:11,color:C.text3}}>#{detail.nuvem} · NF {detail.nf} · {detail.transportadora}</div>
                </div>
                <button onClick={()=>setSelSup(null)} style={{background:"transparent",border:`1px solid ${C.border}`,color:C.text4,cursor:"pointer",fontSize:16,width:32,height:32,display:"flex",alignItems:"center",justifyContent:"center",borderRadius:8}}>×</button>
              </div>
              <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:14}}>
                <StatusBadge val={detail.status}/><Chip val={detail.urgencia} styles={urgStyles}/><Chip val={detail.atendimento} styles={atendStyles}/><TimeOpenBadge sentAt={detail.sentAt}/>
              </div>
              {detail.alertaStatus&&(
                <div style={{background:C.amberSoft,border:`1px solid ${C.amberBorder}`,borderRadius:10,padding:"12px 16px",marginBottom:14,display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:14}}>
                  <div>
                    <div style={{fontSize:12,fontWeight:600,color:C.amber,marginBottom:4}}>⚠ Status atualizado — notifique o cliente!</div>
                    <div style={{fontSize:11,color:C.amber,marginBottom:2}}>{detail.alertaStatus}</div>
                    <div style={{fontSize:10,color:`${C.amber}99`}}>Use os textos prontos abaixo para contato via WhatsApp ou Zendesk</div>
                  </div>
                  <button onClick={()=>upd(detail.id,{alertaStatus:null},{acao:"Alerta dispensado — cliente notificado"})}
                    style={{background:C.amber,border:"none",color:C.white,borderRadius:7,padding:"7px 14px",fontSize:10,cursor:"pointer",fontWeight:600,whiteSpace:"nowrap",flexShrink:0}}>
                    ✓ Notifiquei
                  </button>
                </div>
              )}
              {perms?.canOperate&&(
                <div style={{display:"flex",gap:8}}>
                  {detail.atendimento==="Aberto"&&<button onClick={()=>handleInitiate(detail.id)} style={{flex:1,background:C.brand,border:"none",color:C.white,borderRadius:8,padding:"10px 0",fontSize:11,cursor:"pointer",fontWeight:500,letterSpacing:"0.1em",textTransform:"uppercase"}}>Iniciar atendimento</button>}
                  {detail.atendimento==="Em andamento"&&<button onClick={()=>handleResolve(detail.id)} style={{flex:1,background:C.gold,border:"none",color:C.white,borderRadius:8,padding:"10px 0",fontSize:11,cursor:"pointer",fontWeight:500,letterSpacing:"0.06em"}}>Marcar como resolvido →</button>}
                  <button onClick={()=>handleReturnLog(detail.id)} style={{background:C.white,border:`1px solid ${C.border}`,color:C.text2,borderRadius:8,padding:"10px 16px",fontSize:11,cursor:"pointer",whiteSpace:"nowrap",boxShadow:shadow.sm}}>← Devolver</button>
                </div>
              )}
            </div>

            <div style={{padding:"22px 28px",flex:1}}>
              {/* Info grid */}
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:18,background:C.white,borderRadius:12,border:`1px solid ${C.border}`,padding:18,boxShadow:shadow.sm}}>
                {[["Rastreio",<span style={{fontFamily:"monospace",fontSize:11,color:C.text1}}>{detail.rastreio||"—"}</span>],["Motivo",detail.motivo],["Transportadora",detail.transportadora]].map(([lbl,val])=>(
                  <div key={lbl}><div style={{fontSize:8,color:C.text4,textTransform:"uppercase",letterSpacing:"0.12em",marginBottom:5,fontWeight:500}}>{lbl}</div><div style={{fontSize:12,color:C.text1}}>{val}</div></div>
                ))}
                <div><div style={{fontSize:8,color:C.text4,textTransform:"uppercase",letterSpacing:"0.12em",marginBottom:5,fontWeight:500}}>Prazo / SLA</div><SlaCell prazo={detail.prazo}/></div>
              </div>

              {perms?.canOperate&&(
                <>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:14}}>
                    {[["Responsável","responsavel","Nome do responsável..."],["Nº Chamado Zendesk","chamado","Ex: #45821"]].map(([lbl,key,ph])=>(
                      <div key={key}>
                        <div style={{fontSize:8,color:C.text4,textTransform:"uppercase",letterSpacing:"0.12em",marginBottom:6,fontWeight:500}}>{lbl}</div>
                        <input value={detail[key]||""} onChange={e=>upd(detail.id,{[key]:e.target.value})} placeholder={ph} style={{...INP,width:"100%",boxSizing:"border-box"}}/>
                      </div>
                    ))}
                  </div>
                  <div style={{marginBottom:20}}>
                    <div style={{fontSize:8,color:C.text4,textTransform:"uppercase",letterSpacing:"0.12em",marginBottom:6,fontWeight:500}}>Observações</div>
                    <textarea value={detail.obs||""} onChange={e=>upd(detail.id,{obs:e.target.value})} placeholder="Anotações do atendimento..." rows={3}
                      style={{width:"100%",borderRadius:8,border:`1px solid ${C.border}`,padding:"10px 12px",fontSize:12,resize:"vertical",fontFamily:"inherit",background:C.white,color:C.text1,boxSizing:"border-box",lineHeight:1.6}}/>
                  </div>
                </>
              )}

              <div style={{height:1,background:C.border,marginBottom:16}}/>

              {/* Templates */}
              <div style={{marginBottom:10,background:C.white,borderRadius:12,border:`1px solid ${C.border}`,overflow:"hidden",boxShadow:shadow.sm}}>
                <button onClick={()=>setOpenTpl(v=>!v)} style={{width:"100%",padding:"14px 18px",background:"transparent",border:"none",cursor:"pointer",display:"flex",justifyContent:"space-between",alignItems:"center",fontSize:11,fontWeight:500,color:C.text1,letterSpacing:"0.04em"}}>
                  <span>✉ Textos prontos para atendimento</span>
                  <span style={{fontSize:16,color:C.text4,fontWeight:300,width:24,textAlign:"center"}}>{openTpl?"−":"+"}</span>
                </button>
                {openTpl&&<div style={{padding:16,borderTop:`1px solid ${C.border}`,background:C.cream}}>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                    {[{ch:"wpp",label:"WhatsApp"},{ch:"zendesk",label:"Zendesk"}].map(({ch,label})=>(
                      <div key={ch} style={{background:C.white,borderRadius:10,border:`1px solid ${C.border}`,overflow:"hidden",boxShadow:shadow.sm}}>
                        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"10px 14px",borderBottom:`1px solid ${C.border}`,background:C.brand}}>
                          <span style={{fontSize:10,color:C.gold,fontWeight:500,letterSpacing:"0.1em",textTransform:"uppercase"}}>{label}</span>
                          <CopyBtn text={getTemplate(detail,ch)} label="Copiar"/>
                        </div>
                        <pre style={{margin:0,padding:"12px 14px",fontSize:11,color:C.text2,whiteSpace:"pre-wrap",fontFamily:"inherit",lineHeight:1.7,maxHeight:160,overflowY:"auto"}}>{getTemplate(detail,ch)}</pre>
                      </div>
                    ))}
                  </div>
                </div>}
              </div>

              {/* Histórico */}
              <div style={{background:C.white,borderRadius:12,border:`1px solid ${C.border}`,overflow:"hidden",boxShadow:shadow.sm}}>
                <button onClick={()=>setOpenHist(v=>!v)} style={{width:"100%",padding:"14px 18px",background:"transparent",border:"none",cursor:"pointer",display:"flex",justifyContent:"space-between",alignItems:"center",fontSize:11,fontWeight:500,color:C.text1,letterSpacing:"0.04em"}}>
                  <span>Histórico{detail.historico.length>0?` (${detail.historico.length})`:""}</span>
                  <span style={{fontSize:16,color:C.text4,fontWeight:300,width:24,textAlign:"center"}}>{openHist?"−":"+"}</span>
                </button>
                {openHist&&<div style={{padding:16,borderTop:`1px solid ${C.border}`,background:C.cream}}>
                  {detail.historico.length===0?<div style={{fontSize:12,color:C.text4}}>Nenhuma ação registrada.</div>
                  :<div style={{display:"flex",flexDirection:"column",gap:10}}>
                    {detail.historico.map((h,i)=>(
                      <div key={i} style={{display:"flex",gap:12,alignItems:"flex-start"}}>
                        <div style={{width:6,height:6,borderRadius:"50%",background:C.gold,flexShrink:0,marginTop:4}}/>
                        <div style={{flex:1}}>
                          <div style={{fontSize:10,color:C.text4,marginBottom:2}}>{h.ts}</div>
                          <div style={{fontSize:11,color:C.text2}}>{h.acao}</div>
                        </div>
                      </div>
                    ))}
                  </div>}
                </div>}
              </div>
            </div>
          </div>
        ):ss.total>0?<div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center",background:C.cream}}>
          <div style={{textAlign:"center"}}>
            <div style={{width:48,height:1,background:C.border,margin:"0 auto 20px"}}/>
            <div style={{fontSize:10,letterSpacing:"0.18em",textTransform:"uppercase",color:C.text4}}>Selecione um pedido para atender</div>
            <div style={{fontSize:9,color:C.text4,marginTop:6,letterSpacing:"0.06em"}}>ou selecione vários para arquivar em lote</div>
          </div>
        </div>:null}
      </div>
    )}

    {/* ── ARQUIVADOS ── */}
    {tab==="arquivados"&&(
      <div style={{padding:"24px 32px"}}>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12,marginBottom:20}}>
          <KpiCard label="Total arquivados" val={arch}/>
          <KpiCard label="Resolvidos hoje" val={rows.filter(r=>{if(r.atendimento!=="Resolvido")return false;const h=r.historico.find(x=>x.acao&&(x.acao.includes("Resolvido")||x.acao.includes("Arquivado")));return h&&h.ts&&h.ts.startsWith(new Date().toLocaleDateString("pt-BR"))}).length}/>
          <KpiCard label="Com observações" val={baseArch.filter(r=>r.obs&&r.obs.trim()).length}/>
        </div>
        {arch===0?<div style={{textAlign:"center",padding:"56px 0",color:C.text4}}><div style={{fontSize:32,marginBottom:12,opacity:0.2}}>◎</div><div style={{fontSize:14}}>Nenhum atendimento arquivado</div></div>:(
          <div>
            <div style={{marginBottom:14}}><input value={aSrch} onChange={e=>setASrch(e.target.value)} placeholder="Buscar nos arquivados..." style={{...INP,width:"100%",padding:"10px 14px",boxSizing:"border-box",boxShadow:shadow.sm}}/></div>
            <div style={{overflowX:"auto",overflowY:"auto",maxHeight:"60vh",borderRadius:12,border:`1px solid ${C.border}`,boxShadow:shadow.sm}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:compact?11:12,tableLayout:"fixed",minWidth:900}}>
                <colgroup><col style={{width:80}}/><col style={{width:140}}/><col style={{width:98}}/><col style={{width:98}}/><col style={{width:108}}/><col style={{width:64}}/><col style={{width:100}}/><col style={{width:88}}/><col style={{width:82}}/><col style={{width:90}}/></colgroup>
                <thead><tr>{["No NUVEM","Destinatário","Transportadora","Status","Motivo","Urgência","Prazo / SLA","Chamado","Responsável","Ações"].map(h=><th key={h} style={THF}>{h}</th>)}</tr></thead>
                <tbody>
                  {archRows.length===0?<tr><td colSpan={10} style={{textAlign:"center",padding:32,color:C.text4}}>Nenhum resultado</td></tr>
                  :archRows.map((r,i)=>(
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
                      <td style={{padding:`${pd}px 14px`}}>{perms?.canOperate&&<button onClick={()=>upd(r.id,{atendimento:"Em andamento"},{acao:"Reaberto dos arquivados"})} style={{background:C.cream,border:`1px solid ${C.border}`,color:C.text2,borderRadius:6,padding:"4px 12px",fontSize:10,cursor:"pointer",fontWeight:500}}>Reabrir</button>}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div style={{marginTop:10,fontSize:11,color:C.text4}}>{archRows.length} de {arch} arquivados</div>
          </div>
        )}
      </div>
    )}
  </div>
}
