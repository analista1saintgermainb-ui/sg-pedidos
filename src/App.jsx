import { useState, useEffect, useRef, useCallback } from 'react'
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell } from 'recharts'
import * as XLSX from 'xlsx'
import { createClient } from '@supabase/supabase-js'

// ─── Design System (mantido igual) ─────────────────────────────────
const CL = { brand: "#0C0C0C", brandSoft: "#1A1A1A", gold: "#B8974A", goldLight: "#D4AF6A", goldDim: "#8C7038", cream: "#F8F5EF", creamDark: "#F0EDE5", white: "#FFFFFF", border: "#E8E3D8", borderDark: "#D4CFC4", text1: "#1A1A1A", text2: "#5C5750", text3: "#9C9690", text4: "#C4C0B8", red: "#C0392B", redSoft: "#F9ECEB", redBorder: "#EBCBC8", green: "#2E7D50", greenSoft: "#EAF4EE", greenBorder: "#C0DCCB", amber: "#8C6D1F", amberSoft: "#FDF6E3", amberBorder: "#E8D5A3", blue: "#1A5276", blueSoft: "#EAF2FB", blueBorder: "#AACDE6" }
const CD = { brand: "#F0EDE5", brandSoft: "#D4CFC4", gold: "#D4AF6A", goldLight: "#E8C97A", goldDim: "#B8974A", cream: "#0F0F0F", creamDark: "#1A1A1A", white: "#1E1E1E", border: "#2E2E2E", borderDark: "#3A3A3A", text1: "#F0EDE5", text2: "#C4C0B8", text3: "#7A7670", text4: "#4A4640", red: "#E05555", redSoft: "#2A1212", redBorder: "#4A2020", green: "#4AB870", greenSoft: "#0A2015", greenBorder: "#1A4030", amber: "#C89830", amberSoft: "#281E00", amberBorder: "#4A3800", blue: "#5A9AD4", blueSoft: "#0A1825", blueBorder: "#1A3A55" }
const C = { ...CL }
function applyTheme(dark) { const src = dark ? CD : CL; Object.keys(src).forEach(k => { C[k] = src[k] }) }
const shadow = { sm: "0 1px 3px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04)", md: "0 4px 12px rgba(0,0,0,0.08), 0 2px 4px rgba(0,0,0,0.04)", lg: "0 8px 24px rgba(0,0,0,0.10), 0 4px 8px rgba(0,0,0,0.06)" }
const getGlobalStyle = () => `@import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;500;600&family=Inter:wght@300;400;500;600&display=swap'); * { box-sizing: border-box; margin: 0; padding: 0; } body { background: ${C.cream}; color: ${C.text1}; } ::-webkit-scrollbar { width: 4px; height: 4px; } ::-webkit-scrollbar-track { background: ${C.creamDark}; } ::-webkit-scrollbar-thumb { background: ${C.borderDark}; border-radius: 2px; } ::-webkit-scrollbar-thumb:hover { background: ${C.gold}; } select, input, textarea, button { font-family: 'Inter', sans-serif; } tr:hover td { background: ${C.creamDark} !important; }`

// ─── Supabase Client (NOVO) ──────────────────────────────────────
const supabase = createClient(SUPA_URL, SUPA_KEY)

// ─── Suas funções existentes (signIn, signOut, createUser, dbLoadFast, dbUpsert, dbDelete, dbClear, loadUsuarios, saveUsuario, deleteUsuario) permanecem IGUAIS ───
// (copie exatamente como estavam, sem alterações)
// ... [código de SUPA_URL, SUPA_KEY, SH, aSH, etc. permanece idêntico] ...

// ─── Permissões, ALERTA_DIAS, HEADER_MAP, norm, findIdx, uniq, QFILTERS, PAGE_SIZE, helpers, parseData, applyQF, applySortRows, exportCSV, getTemplate, classificarProblema, TRANSP_LINKS, PROBLEMA_CONFIG, estilos, etc. permanecem IGUAIS ───

// ─── Componentes: Chip, StatusBadge, SlaCell, SituacaoPrazoBadge, SemMovBadge, TimeOpenBadge, KpiCard, CopyBtn, SortIcon, Toast, getINP, HeaderProblema, SugestaoSistema, buildMailto, AcoesRapidas, TimelineHistorico, LoginScreen, BoxlinkSettings, UsuariosPanel, KanbanSuporteView, FilterBar, etc. permanecem IGUAIS ───

// ─── Boxlink API Integration (mantido igual) ─────────────────────
const BOXLINK_API = "https://api.boxlink.com.br"
const BOXLINK_TOKEN_KEY = "sg_boxlink_token"
function getBoxlinkToken() { return localStorage.getItem(BOXLINK_TOKEN_KEY) || "" }
function setBoxlinkToken(t) { localStorage.setItem(BOXLINK_TOKEN_KEY, t) }
// ... (mapBoxlinkRow, BOXLINK_TRACKING_PATHS, fetchBoxlinkPage, syncBoxlinkFull iguais) ...

// ======================= NOVA IMPLEMENTAÇÃO DO REALTIME =======================

export default function App() {
  const [session, setSession] = useState(null)
  const [perfil, setPerfil] = useState(null)
  const [nomeAtendente, setNomeAtendente] = useState("")
  const [loadingPerfil, setLoadingPerfil] = useState(false)
  const [rows, setRows] = useState([])
  const [tab, setTab] = useState(null)
  const [paste, setPaste] = useState("")
  const [importing, setImporting] = useState(false)
  const [loadingData, setLoadingData] = useState(false)
  const [compact, setCompact] = useState(false)
  const [dark, setDark] = useState(false)
  const [toasts, setToasts] = useState([])
  // Filtros logística
  const [lSrch, setLSrch] = useState(""); const [lSt, setLSt] = useState("Todos")
  const [lTr, setLTr] = useState("Todos"); const [lUrg, setLUrg] = useState("Todos")
  const [lAc, setLAc] = useState("Todos"); const [lSitPrazo, setLSitPrazo] = useState("Todos"); const [qf, setQf] = useState("todos")
  const [lShowFilters, setLShowFilters] = useState(false)
  // Filtros suporte
  const [sResp, setSResp] = useState("Todos"); const [sShowFilters, setSShowFilters] = useState(false)
  const [lPage, setLPage] = useState(1)
  const [selIds, setSelIds] = useState(new Set())
  const [sortCol, setSortCol] = useState(null); const [sortDir, setSortDir] = useState("asc")
  const [sSrch, setSSrch] = useState(""); const [sAtend, setSAtend] = useState("Todos")
  const [sUrg, setSUrg] = useState("Todos")
  const [selSup, setSelSup] = useState(null)
  const [supView, setSupView] = useState('lista')
  const [selSupIds, setSelSupIds] = useState(new Set())
  const [openTpl, setOpenTpl] = useState(false); const [openHist, setOpenHist] = useState(false)
  const [aSrch, setASrch] = useState("")
  const [aPage, setAPage] = useState(1)
  // Status de sincronização
  const [syncStatus, setSyncStatus] = useState("idle")
  const [lastSync, setLastSync] = useState(null)
  // NOVO: status da conexão Realtime
  const [realtimeStatus, setRealtimeStatus] = useState("connecting")
  // Boxlink
  const [bxToken, setBxToken] = useState(() => getBoxlinkToken())
  const [bxStatus, setBxStatus] = useState("idle")
  const [bxCountdown, setBxCountdown] = useState(15 * 60)

  const saveTimer = useRef(null);
  const fileRef = useRef()
  const token = session?.access_token

  const addToast = useCallback((msg, type = "ok", ms = 4000) => {
    const id = Date.now(); setToasts(p => [...p, { id, msg, type }]); setTimeout(() => setToasts(p => p.filter(t => t.id !== id)), ms)
  }, [])

  const handleLogin = async data => { /* igual */ }
  const handleLogout = async () => { /* igual */ }
  const perms = perfil ? PERMS[perfil] : null

  // ─── (1) CARGA INICIAL E SUBSCRIÇÃO REALTIME (substitui polling) ───
  useEffect(() => {
    if (!token) return

    let subscription = null
    let isMounted = true

    const loadInitialData = async () => {
      setSyncStatus("loading")
      setLoadingData(true)
      try {
        // Busca todos os pedidos via Supabase client (mais rápido e tipado)
        const { data, error } = await supabase
          .from('pedidos')
          .select('*')
          .order('id', { ascending: true })

        if (error) throw error

        const mapped = data.map(row => ({ ...row.dados, id: row.id }))
        const fixRows = (rows) => rows.map(r => ({
          ...r,
          isNew: false,
          atendimento: isEntregue(r.status) && !r.enviadoSuporte ? "Resolvido" : r.atendimento,
          enviadoSuporte: isEntregue(r.status) && !r.enviadoSuporte ? false : r.enviadoSuporte,
        }))

        if (isMounted) {
          setRows(fixRows(mapped))
          setLastSync(new Date())
          setSyncStatus("idle")
        }
      } catch (err) {
        if (isMounted) {
          setSyncStatus("error")
          addToast("Erro ao carregar dados: " + err.message, "error", 8000)
        }
      } finally {
        if (isMounted) setLoadingData(false)
      }
    }

    loadInitialData()

    // Inscreve para mudanças em tempo real
    subscription = supabase
      .channel('pedidos-realtime')
      .on(
        'postgres_changes',
        { event: '*', schema: 'public', table: 'pedidos' },
        async (payload) => {
          // Evita processar se a mudança foi feita pelo próprio cliente (opcional)
          if (payload.eventType === 'INSERT') {
            const newRow = { ...payload.new.dados, id: payload.new.id }
            setRows(prev => {
              if (prev.some(r => r.id === newRow.id)) return prev
              return [...prev, { ...newRow, isNew: true }]
            })
            addToast(`📦 Novo pedido ${newRow.nuvem} adicionado`, "warn")
          } else if (payload.eventType === 'UPDATE') {
            const updatedRow = { ...payload.new.dados, id: payload.new.id }
            setRows(prev => prev.map(r => r.id === updatedRow.id ? { ...updatedRow, isNew: true } : r))
            addToast(`✏️ Pedido ${updatedRow.nuvem} atualizado`, "warn")
          } else if (payload.eventType === 'DELETE') {
            const deletedId = payload.old.id
            setRows(prev => prev.filter(r => r.id !== deletedId))
            addToast(`🗑️ Pedido removido`, "warn")
          }
        }
      )
      .subscribe((status) => {
        if (status === 'SUBSCRIBED') setRealtimeStatus('connected')
        else if (status === 'CHANNEL_ERROR') setRealtimeStatus('error')
        else setRealtimeStatus('connecting')
      })

    return () => {
      if (subscription) subscription.unsubscribe()
      isMounted = false
    }
  }, [token, addToast]) // <- dependência apenas do token

  // ─── (2) REMOVA O ANTIGO useEffect DE POLLING (aquele com setInterval) ───
  // Simplesmente não inclua o código antigo. O polling foi totalmente removido.

  // Boxlink auto-sync (mantido igual)
  useEffect(() => {
    if (!bxToken || !token) return
    const doBoxlinkSync = async (bToken, silent = false) => { /* igual */ }
    doBoxlinkSync(bxToken, true)
    const cd = setInterval(() => setBxCountdown(p => { if (p <= 1) { doBoxlinkSync(bxToken, true); return 15 * 60 } return p - 1 }), 1000)
    return () => clearInterval(cd)
  }, [bxToken, token])

  // Salvamento automático (mantido igual)
  useEffect(() => {
    if (!token || rows.length === 0) return
    if (saveTimer.current) clearTimeout(saveTimer.current)
    saveTimer.current = setTimeout(async () => {
      setSyncStatus("saving")
      try {
        await dbUpsert(rows, token)
      } catch (e) {
        setSyncStatus("error")
        addToast("Erro ao salvar: " + e.message, "error", 8000)
        setTimeout(() => setSyncStatus("idle"), 4000)
        return
      }
      setLastSync(new Date())
      setSyncStatus("saved")
      setTimeout(() => setSyncStatus("idle"), 2500)
    }, 1200)
  }, [rows, token, addToast])

  // Efeito para "isNew" expirar (mantido igual)
  useEffect(() => {
    if (!rows.some(r => r.isNew)) return
    const t = setTimeout(() => setRows(p => p.map(r => ({ ...r, isNew: false }))), 6000)
    return () => clearTimeout(t)
  }, [rows])

  // Demais useEffects (páginação, reset de filtros, etc.) permanecem IGUAIS

  // O resto do componente App (renderização, lógica de filtros, etc.) permanece IDÊNTICO
  // ... (todo o JSX abaixo permanece igual, apenas remova qualquer referência a `countdown` no header)

  // No header, onde exibia "Sync em Xs", substitua por um indicador de conexão Realtime
  // Exemplo de alteração no header (dentro do bloco que mostra sync status):
  // 
  // <div style={{display:"flex",alignItems:"center",gap:7}}>
  //   <span style={{width:6,height:6,borderRadius:"50%",background:realtimeStatus==="connected"?"#27ae60":realtimeStatus==="error"?C.red:C.gold}}/>
  //   <span style={{fontSize:10,color:realtimeStatus==="connected"?"#27ae60":realtimeStatus==="error"?C.red:C.gold}}>
  //     {realtimeStatus==="connected"?"Ao vivo":realtimeStatus==="error"?"Offline":"Conectando..."}
  //   </span>
  // </div>

  // O restante do JSX não precisa ser alterado, apenas remova o `countdown` e `setCountdown`.

  return ( /* JSX existente, com a alteração do indicador de sincronismo */ )
}
