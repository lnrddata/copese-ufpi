import { useEffect, useState } from 'react'
import * as XLSX from 'xlsx'
import './App.css'

const UiIcon = ({ name, size = 18 }) => {
  const paths = {
    dashboard: <><rect x="3" y="3" width="6" height="6" rx="1" /><rect x="15" y="3" width="6" height="6" rx="1" /><rect x="3" y="15" width="6" height="6" rx="1" /><rect x="15" y="15" width="6" height="6" rx="1" /></>,
    list: <><path d="M5 6h14M5 12h14M5 18h14" /><circle cx="3" cy="6" r=".8" fill="currentColor" stroke="none" /><circle cx="3" cy="12" r=".8" fill="currentColor" stroke="none" /><circle cx="3" cy="18" r=".8" fill="currentColor" stroke="none" /></>,
    vacancies: <><path d="M5 4.5h5v6H5zM4 10.5h7v3H4zM14 4.5h5v6h-5zM13 10.5h7v3h-7z" /></>,
    occupied: <><circle cx="12" cy="12" r="8.5" /><path d="m8.2 12 2.5 2.5 5.2-5.2" /></>,
    remaining: <><path d="M8 4.5h8v7H8zM6 11.5h12v3H6z" /></>,
    waiting: <><path d="M5 5.5h14v13H5z" /><path d="M8 9h8M8 13h8M8 17h5" /></>,
  }
  return <svg className="ui-icon" width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{paths[name]}</svg>
}


// ============================================================
// FONTE ÚNICA DE DADOS
// O dashboard lê o arquivo diretamente de /public.
// No GitHub Pages, o arquivo será servido em:
// /Dados_copese.xlsx
// ============================================================

const DATA_FILE = `${import.meta.env.BASE_URL}Dados_copese.xlsx`

const emptyDashboardData = {
  courses: [],
  calls: [],
  modalityRows: [],
  modalityByCourse: [],
  waitingList: [],
  convocacaoMotivoByCandidate: {},
}

const normalizeKey = (value) =>
  String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/\s+/g, ' ')
    .trim()

const toNumber = (value) => {
  if (value === null || value === undefined || value === '') return 0
  const n = Number(String(value).replace(',', '.'))
  return Number.isFinite(n) ? n : 0
}

const toRows = (sheet) => {
  if (!sheet) return []
  return XLSX.utils.sheet_to_json(sheet, { defval: null })
}

function buildDashboardData(workbook) {
  const vagas = toRows(workbook.Sheets.vagas)
  const convocacoes = toRows(workbook.Sheets.convocacoes)
  const listaEspera = toRows(workbook.Sheets.lista_espera)

  const callOrder = [
    'Chamada Regular',
    '1º Chamada da lista de espera',
    '2º Chamada da lista de espera',
    '3º Chamada da lista de espera',
  ]

  // Vagas por curso/campus: a aba "vagas" traz as etapas do processo.
  const courseMap = new Map()

  vagas.forEach((row) => {
    const key = `${row.campus_centro}||${row.curso}`
    if (!courseMap.has(key)) {
      courseMap.set(key, {
        name: row.curso,
        campus: row.campus_centro,
        vacancies: 0,
        occupied: 0,
        remaining: 0,
        firstCall: 0,
        firstCallOccupied: 0,
        secondCall: 0,
        secondCallOccupied: 0,
        thirdCall: 0,
        thirdCallOccupied: 0,
        waitingTotal: 0,
        awaiting: 0,
        waitingConvoked: 0,
      })
    }

    const course = courseMap.get(key)
    const vacancies = toNumber(row.vagas)
    const occupied = toNumber(row.ocupacoes)
    const remaining = toNumber(row.remanescentes)

    if (row.chamada === 'Chamada Regular') {
      course.vacancies += vacancies
      course.occupied += occupied
      course.remaining += remaining
    } else if (row.chamada === '1º Chamada da lista de espera') {
      course.firstCall += vacancies
      course.firstCallOccupied += occupied
    } else if (row.chamada === '2º Chamada da lista de espera') {
      course.secondCall += vacancies
      course.secondCallOccupied += occupied
    } else if (row.chamada === '3º Chamada da lista de espera') {
      course.thirdCall += vacancies
      course.thirdCallOccupied += occupied
    }
  })

  // OCUPAÇÃO ATUAL REAL:
  // cada vaga da Chamada Regular é uma vaga da oferta original.
  // As chamadas posteriores preenchem somente as vagas que ainda estavam
  // disponíveis. Portanto, para cada curso/campus:
  // ocupação atual = ocupação da Chamada Regular + ocupações posteriores,
  // limitada à oferta original. Nunca se considera uma nova chamada como
  // uma nova oferta de vagas para aumentar o denominador.
  const additionalOccupiedByCourse = new Map()
  vagas
    .filter((row) => row.chamada !== 'Chamada Regular')
    .forEach((row) => {
      const key = `${row.campus_centro}||${row.curso}`
      additionalOccupiedByCourse.set(
        key,
        (additionalOccupiedByCourse.get(key) || 0) + toNumber(row.ocupacoes)
      )
    })

  courseMap.forEach((course, key) => {
    const additionalOccupied = additionalOccupiedByCourse.get(key) || 0
    const regularOccupied = course.occupied
    course.occupied = Math.min(course.vacancies, regularOccupied + additionalOccupied)
    course.remaining = Math.max(0, course.vacancies - course.occupied)
  })

  const courses = [...courseMap.values()]
    .sort((a, b) =>
      a.campus.localeCompare(b.campus, 'pt-BR') ||
      a.name.localeCompare(b.name, 'pt-BR')
    )

  // Resumo por chamada.
  const calls = callOrder
    .map((name) => {
      const rows = vagas.filter((row) => row.chamada === name)
      if (!rows.length) return null
      const vacancies = rows.reduce((sum, row) => sum + toNumber(row.vagas), 0)
      const occupied = rows.reduce((sum, row) => sum + toNumber(row.ocupacoes), 0)
      const remaining = rows.reduce((sum, row) => sum + toNumber(row.remanescentes), 0)
      return {
        name,
        vacancies,
        occupied,
        remaining,
        summoned: vacancies,
      }
    })
    .filter(Boolean)

  // Ocupação por modalidade — situação acumulada do processo.
  // A ocupação soma as ocupações registradas em todas as chamadas.
  // As vagas remanescentes são as da última chamada registrada para cada
  // curso/campus/modalidade. Dessa forma, a situação final não duplica vagas
  // entre chamadas e permanece compatível com a oferta total do processo.
  const modalityAccumulator = new Map()
  vagas.forEach((row) => {
    const campus = String(row.campus_centro ?? '').trim()
    const course = String(row.curso ?? '').trim()
    const modality = String(row.modalidade ?? '').trim()
    const key = `${campus}||${course}||${modality}`
    const current = modalityAccumulator.get(key) || {
      campus,
      course,
      modality,
      occupiedAcrossCalls: 0,
      latestCallRank: -1,
      latestRemaining: 0,
      hasRegularBase: false,
    }

    if (row.chamada === 'Chamada Regular') {
      current.hasRegularBase = true
    }

    current.occupiedAcrossCalls += toNumber(row.ocupacoes)

    const rowCallRank = callOrder.indexOf(row.chamada)
    if (rowCallRank >= current.latestCallRank) {
      current.latestCallRank = rowCallRank
      current.latestRemaining = toNumber(row.remanescentes)
    }

    modalityAccumulator.set(key, current)
  })

  const modalityByCourse = [...modalityAccumulator.values()]
    .filter((row) => row.hasRegularBase)
    .map((row) => {
      const occupied = row.occupiedAcrossCalls
      const remaining = Math.max(0, row.latestRemaining)
      return {
        campus: row.campus,
        course: row.course,
        modality: row.modality,
        vacancies: occupied + remaining,
        occupied,
        remaining,
      }
    })

  // Resumo por modalidade para a tabela "Ocupação por modalidade".
  const modalityRows = [...modalityByCourse
    .reduce((modalityMap, row) => {
      const current = modalityMap.get(row.modality) || {
        name: row.modality,
        vacancies: 0,
        occupied: 0,
        remaining: 0,
      }
      current.vacancies += row.vacancies
      current.occupied += row.occupied
      current.remaining += row.remaining
      modalityMap.set(row.modality, current)
      return modalityMap
    }, new Map())
    .values()]

  // Lista de espera.
  const waitingList = listaEspera.map((row) => ({
    campus: row.campus_centro,
    course: row.curso,
    courseShort: row.curso_ab || row.curso,
    shift: row.Turno,
    registration: String(row.inscricao ?? ''),
    candidate: row.candidato,
    modality: row.modalidade_conc,
    score: row.nota == null || row.nota === '' ? null : toNumber(row.nota),
    classification: toNumber(row.classificacao),
    call: row.chamada || '-',
    status: row.situação || 'Aguardando convocação',
    calledModality: row.modalidade_conv || '-',
  }))

  // Motivos de não matrícula: derivados diretamente da aba "convocacoes".
  const convocacaoMotivoByCandidate = {}
  convocacoes.forEach((row) => {
    const candidateKey = normalizeKey(row.candidato)
    const motivo = row.motivo_n_matriculado
    if (candidateKey && motivo) {
      convocacaoMotivoByCandidate[candidateKey] = motivo
    }
  })

  return {
    courses,
    calls,
    modalityRows,
    modalityByCourse,
    waitingList,
    convocacaoMotivoByCandidate,
  }
}
const number = (value) => Number(value || 0).toLocaleString('pt-BR')
const normalizeText = (value) => String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/\s+/g, ' ').trim()
const percent = (value, total) => total ? `${Math.round((value / total) * 100)}%` : '0%'

/*
 * Gráfico reconstruído em uma única área SVG.
 * Barras, linha, pontos e percentuais usam as mesmas coordenadas.
 * Os dados recebidos são somente de leitura; nenhum cálculo do dashboard
 * é alterado neste componente.
 */
const CumulativeOccupancyChart = ({ snapshots, capacity, formatNumber, formatRate }) => {
  const WIDTH = 600
  const HEIGHT = 320
  const LINE_TOP = 18
  const LINE_HEIGHT = 48
  const BARS_TOP = 82
  const BARS_BOTTOM = 238
  const LABEL_TOP = 257
  const count = Math.max(snapshots.length, 1)
  const safeCapacity = Math.max(Number(capacity) || 1, 1)
  const barWidth = 58

  const clamp = (value, min = 0, max = 1) => Math.min(max, Math.max(min, Number(value) || 0))
  const xForIndex = (index) => ((index + 0.5) / count) * WIDTH
  const yForRate = (rate) => LINE_TOP + ((100 - clamp(rate, 0, 100)) / 100) * LINE_HEIGHT
  const yForValue = (value) => BARS_BOTTOM - clamp(value / safeCapacity) * (BARS_BOTTOM - BARS_TOP)
  const linePoints = snapshots.map((call, index) => `${xForIndex(index)},${yForRate(call.rate)}`).join(' ')

  const labelLines = (name) => {
    if (name === 'Chamada Regular') return ['Chamada', 'Regular']
    if (name.includes('1º')) return ['1º Chamada da lista de', 'espera']
    if (name.includes('2º')) return ['2º Chamada da lista de', 'espera']
    if (name.includes('3º')) return ['3º Chamada da lista de', 'espera']
    return [name]
  }

  return (
    <div className="cumulative-visual" role="img" aria-label="Evolução acumulada das vagas ocupadas e remanescentes por chamada">
      <svg className="cumulative-visual-svg" viewBox={`0 0 ${WIDTH} ${HEIGHT}`} preserveAspectRatio="xMidYMid meet" aria-hidden="true">
        <line className="cumulative-visual-baseline" x1="0" y1={BARS_BOTTOM} x2={WIDTH} y2={BARS_BOTTOM} />

        {snapshots.map((call, index) => {
          const x = xForIndex(index)
          const occupiedY = yForValue(call.occupied)
          const remainingHeight = Math.max(0, occupiedY - BARS_TOP)
          const occupiedHeight = Math.max(0, BARS_BOTTOM - occupiedY)
          const lines = labelLines(call.name)
          return (
            <g className="cumulative-visual-column" key={call.name}>
              <rect className="cumulative-visual-remaining" x={x - barWidth / 2} y={BARS_TOP} width={barWidth} height={remainingHeight} rx="7" />
              <rect className="cumulative-visual-occupied" x={x - barWidth / 2} y={occupiedY} width={barWidth} height={occupiedHeight} rx="6" />
              {call.remaining > 0 && <text className="cumulative-visual-remaining-value" x={x} y={BARS_TOP + Math.max(16, remainingHeight / 2 + 5)} textAnchor="middle">{formatNumber(call.remaining)}</text>}
              {call.occupied > 0 && <text className="cumulative-visual-occupied-value" x={x} y={occupiedY + Math.max(16, occupiedHeight / 2 + 5)} textAnchor="middle">{formatNumber(call.occupied)}</text>}
              <text className="cumulative-visual-category" x={x} y={LABEL_TOP} textAnchor="middle">
                {lines.map((line, lineIndex) => <tspan key={`${call.name}-${line}`} x={x} dy={lineIndex === 0 ? 0 : 17}>{line}</tspan>)}
              </text>
            </g>
          )
        })}

        <polyline className="cumulative-visual-rate-line" points={linePoints} fill="none" />
        {snapshots.map((call, index) => {
          const x = xForIndex(index)
          const y = yForRate(call.rate)
          return (
            <g className="cumulative-visual-rate-point" key={`rate-${call.name}`}>
              <circle cx={x} cy={y} r="4" />
              <text x={x} y={y - 10} textAnchor="middle">{formatRate(call.rate)}</text>
            </g>
          )
        })}
      </svg>

      <div className="cumulative-visual-legend" aria-label="Legenda do gráfico de ocupação acumulada">
        <span><i className="cumulative-visual-key cumulative-visual-key-occupied" />Ocupadas acumuladas</span>
        <span><i className="cumulative-visual-key cumulative-visual-key-remaining" />Vagas remanescentes</span>
        <span><i className="cumulative-visual-key cumulative-visual-key-rate" />Taxa de ocupação acumulada</span>
      </div>
    </div>
  )
}

function App() {
  const [search, setSearch] = useState('')
  const [campus, setCampus] = useState('Todos')
  const [activePage, setActivePage] = useState('overview')
  const [waitingSearch, setWaitingSearch] = useState('')
  const [waitingCourse, setWaitingCourse] = useState('Todos')
  const [waitingCampus, setWaitingCampus] = useState('Todos')
  const [waitingModality, setWaitingModality] = useState('Todos')
  const [waitingCalledModality, setWaitingCalledModality] = useState('Todos')
  const [waitingStatus, setWaitingStatus] = useState('Todos')
  const [theme, setTheme] = useState(() => localStorage.getItem('acompanha-theme') || 'light')
  const [dashboardData, setDashboardData] = useState(emptyDashboardData)
  const [dataLoading, setDataLoading] = useState(true)
  const [dataError, setDataError] = useState('')

  useEffect(() => {
    document.documentElement.dataset.theme = theme
    localStorage.setItem('acompanha-theme', theme)
  }, [theme])

  useEffect(() => {
    let cancelled = false

    const loadWorkbook = async () => {
      try {
        setDataLoading(true)
        setDataError('')

        const response = await fetch(`${DATA_FILE}?v=${Date.now()}`, { cache: 'no-store' })
        if (!response.ok) {
          throw new Error(`Não foi possível carregar ${DATA_FILE} (HTTP ${response.status}).`)
        }

        const buffer = await response.arrayBuffer()
        const workbook = XLSX.read(buffer, { type: 'array' })
        const data = buildDashboardData(workbook)

        if (!cancelled) {
          setDashboardData(data)
        }
      } catch (error) {
        console.error('Erro ao carregar a planilha do dashboard:', error)
        if (!cancelled) {
          setDataError(error?.message || 'Erro ao carregar a planilha.')
        }
      } finally {
        if (!cancelled) {
          setDataLoading(false)
        }
      }
    }

    loadWorkbook()

    return () => {
      cancelled = true
    }
  }, [])

  const {
    courses,
    calls,
    modalityRows,
    modalityByCourse,
    waitingList,
    convocacaoMotivoByCandidate,
  } = dashboardData

  if (dataLoading) {
    return (
      <div className="app-shell theme-light" style={{ minHeight: '100vh', alignItems: 'center', justifyContent: 'center' }}>
        <div className="panel" style={{ maxWidth: 520, margin: 24, textAlign: 'center' }}>
          <strong>Carregando dados da UFPI…</strong>
          <p className="panel-caption">Lendo a planilha Dados_copese.xlsx.</p>
        </div>
      </div>
    )
  }

  if (dataError) {
    return (
      <div className="app-shell theme-light" style={{ minHeight: '100vh', alignItems: 'center', justifyContent: 'center' }}>
        <div className="panel" style={{ maxWidth: 620, margin: 24 }}>
          <strong>Não foi possível carregar os dados.</strong>
          <p className="panel-caption">{dataError}</p>
          <p className="panel-caption">Verifique se Dados_copese.xlsx está dentro da pasta public do projeto.</p>
        </div>
      </div>
    )
  }

  const uniqueWaitingByCourse = waitingList.reduce((courseMap, candidate) => {
    const courseKey = candidate.course
    const candidateKey = candidate.registration || candidate.candidate
    if (!courseMap.has(courseKey)) courseMap.set(courseKey, new Map())
    const candidatesInCourse = courseMap.get(courseKey)
    const wasCalled = candidate.status === 'Convocado' || (candidate.calledModality && candidate.calledModality !== '-')
    const wasFirstCall = candidate.call === '1º Lista de espera' && Boolean(wasCalled)
    const wasSecondCall = candidate.call === '2º Lista de espera' && Boolean(wasCalled)
    const wasThirdCall = (candidate.call === '3º Lista de espera' || candidate.call === '3ª Lista de espera' || candidate.call === '3º Chamada da lista de espera') && Boolean(wasCalled)
    const existing = candidatesInCourse.get(candidateKey)
    if (!existing) {
      candidatesInCourse.set(candidateKey, {
        called: Boolean(wasCalled),
        firstCall: Boolean(wasFirstCall),
        secondCall: Boolean(wasSecondCall),
        thirdCall: Boolean(wasThirdCall),
      })
    } else {
      existing.called = existing.called || Boolean(wasCalled)
      existing.firstCall = existing.firstCall || Boolean(wasFirstCall)
      existing.secondCall = existing.secondCall || Boolean(wasSecondCall)
      existing.thirdCall = existing.thirdCall || Boolean(wasThirdCall)
    }
    return courseMap
  }, new Map())

  const coursesWithDistinctWaiting = courses.map((course) => {
    const uniqueCandidates = [...(uniqueWaitingByCourse.get(course.name)?.values() || [])]
    const firstCall = uniqueCandidates.filter((candidate) => candidate.firstCall).length
    const secondCall = uniqueCandidates.filter((candidate) => candidate.secondCall).length
    const thirdCall = uniqueCandidates.filter((candidate) => candidate.thirdCall).length
    const called = uniqueCandidates.filter((candidate) => candidate.called).length
    return {
      ...course,
      waitingTotal: uniqueCandidates.length,
      firstCall,
      secondCall,
      thirdCall,
      awaiting: uniqueCandidates.length - called,
    }
  })

  const filteredCourses = coursesWithDistinctWaiting.filter((course) => {
    const matchesSearch = `${course.name} ${course.campus}`.toLowerCase().includes(search.toLowerCase())
    const matchesCampus = campus === 'Todos' || course.campus === campus
    return matchesSearch && matchesCampus
  })

  const hasOverviewFilter = Boolean(search.trim()) || campus !== 'Todos'

  const totals = filteredCourses.reduce((acc, course) => ({
    vacancies: acc.vacancies + course.vacancies,
    occupied: acc.occupied + course.occupied,
    remaining: acc.remaining + course.remaining,
    firstCall: acc.firstCall + course.firstCall,
    firstCallOccupied: acc.firstCallOccupied + course.firstCallOccupied,
    secondCall: acc.secondCall + course.secondCall,
    secondCallOccupied: acc.secondCallOccupied + course.secondCallOccupied,
    thirdCall: acc.thirdCall + course.thirdCall,
    thirdCallOccupied: acc.thirdCallOccupied + course.thirdCallOccupied,
    waitingTotal: acc.waitingTotal + course.waitingTotal,
    awaiting: acc.awaiting + course.awaiting,
  }), {
    vacancies: 0, occupied: 0, remaining: 0,
    firstCall: 0, firstCallOccupied: 0,
    secondCall: 0, secondCallOccupied: 0,
    thirdCall: 0, thirdCallOccupied: 0,
    waitingTotal: 0, awaiting: 0,
  })

  const filteredCalls = calls.map((call) => {
    if (!hasOverviewFilter) return call
    if (call.name === 'Chamada Regular') {
      return {
        ...call,
        vacancies: totals.vacancies,
        occupied: totals.occupied,
        remaining: totals.remaining,
        summoned: totals.vacancies,
      }
    }
    if (call.name === '1º Chamada da lista de espera') {
      return {
        ...call,
        vacancies: totals.firstCall,
        occupied: totals.firstCallOccupied,
        remaining: Math.max(0, totals.firstCall - totals.firstCallOccupied),
        summoned: totals.firstCall,
      }
    }
    if (call.name === '2º Chamada da lista de espera') {
      return {
        ...call,
        vacancies: totals.secondCall,
        occupied: totals.secondCallOccupied,
        remaining: Math.max(0, totals.secondCall - totals.secondCallOccupied),
        summoned: totals.secondCall,
      }
    }
    if (call.name === '3º Chamada da lista de espera') {
      return {
        ...call,
        vacancies: totals.thirdCall,
        occupied: totals.thirdCallOccupied,
        remaining: Math.max(0, totals.thirdCall - totals.thirdCallOccupied),
        summoned: totals.thirdCall,
      }
    }
    return call
  })

  // A ocupação por modalidade já foi calculada rigorosamente em
  // buildDashboardData(), no nível campus + curso + modalidade, usando a
  // capacidade da Chamada Regular e a ocupação acumulada das chamadas.
  // Não se deve reconstruí-la a partir da lista de espera, pois isso
  // pode duplicar candidatos e atribuir uma ocupação à modalidade errada.

  const currentModalityByCourse = modalityByCourse
  const modalitySource = currentModalityByCourse

  const filteredModalityRows = [...modalitySource
    .filter((row) => {
      const searchTerm = normalizeText(search)
      if (!hasOverviewFilter) return true
      const matchesSearch = !searchTerm || normalizeText(`${row.course} ${row.campus}`).includes(searchTerm)
      const matchesCampus = campus === 'Todos' || row.campus === campus
      return matchesSearch && matchesCampus
    })
    .reduce((modalityMap, row) => {
      const current = modalityMap.get(row.modality) || { name: row.modality, vacancies: 0, occupied: 0, remaining: 0 }
      current.vacancies += row.vacancies
      current.occupied += row.occupied
      current.remaining = Math.max(0, current.vacancies - current.occupied)
      modalityMap.set(row.modality, current)
      return modalityMap
    }, new Map())
    .values()]

  const filteredWaitingRows = waitingList.filter((candidate) => {
    const term = normalizeText(waitingSearch)
    const normalizedName = normalizeText(candidate.candidate)
    const matchesSearch = !term || normalizedName.includes(term)
    const matchesCourse = waitingCourse === 'Todos' || candidate.course === waitingCourse
    const matchesCampus = waitingCampus === 'Todos' || candidate.campus === waitingCampus
    const matchesModality = waitingModality === 'Todos' || candidate.modality === waitingModality
    const matchesCalledModality = waitingCalledModality === 'Todos' || candidate.calledModality === waitingCalledModality
    const matchesStatus = waitingStatus === 'Todos' || candidate.status === waitingStatus
    return matchesSearch && matchesCourse && matchesCampus && matchesModality && matchesCalledModality && matchesStatus
  })

  const filteredWaiting = Array.from(filteredWaitingRows.reduce((map, candidate) => {
    const key = candidate.registration || candidate.candidate
    const existing = map.get(key)
    if (!existing) {
      map.set(key, { ...candidate, modalities: [candidate.modality], calledModalities: candidate.calledModality !== '-' ? [candidate.calledModality] : [], classifications: [{ modality: candidate.modality, value: candidate.classification }] })
    } else {
      if (!existing.modalities.includes(candidate.modality)) existing.modalities.push(candidate.modality)
      if (candidate.calledModality !== '-' && !existing.calledModalities.includes(candidate.calledModality)) existing.calledModalities.push(candidate.calledModality)
      if (!existing.classifications.some((item) => item.modality === candidate.modality)) existing.classifications.push({ modality: candidate.modality, value: candidate.classification })
      if (candidate.status === 'Convocado') existing.status = 'Convocado'
    }
    return map
  }, new Map()).values()).map((candidate) => ({ ...candidate, modality: candidate.modalities.join(' / '), calledModality: candidate.calledModalities.join(' / ') || '—' })).sort((a, b) => {
    const campusOrder = a.campus.localeCompare(b.campus, 'pt-BR')
    if (campusOrder) return campusOrder
    const courseOrder = a.course.localeCompare(b.course, 'pt-BR')
    if (courseOrder) return courseOrder
    const rankA = Math.min(...a.classifications.map((item) => Number(item.value) || Number.POSITIVE_INFINITY))
    const rankB = Math.min(...b.classifications.map((item) => Number(item.value) || Number.POSITIVE_INFINITY))
    return rankA - rankB || a.candidate.localeCompare(b.candidate, 'pt-BR')
  })

  const waitingStats = filteredWaiting.reduce((acc, candidate) => ({
    total: acc.total + 1,
    waiting: acc.waiting + (candidate.status === 'Aguardando convocação' ? 1 : 0),
    called: acc.called + (candidate.status === 'Convocado' ? 1 : 0),
  }), { total: 0, waiting: 0, called: 0 })
  // Expande as modalidades de cada candidato em linhas independentes.
  // O candidato continua sendo contado uma única vez nos KPIs; apenas a
  // apresentação da tabela é desdobrada por modalidade/classificação.
  const expandedWaiting = filteredWaiting.flatMap((candidate) => (
    candidate.classifications.map((classification) => ({
      ...candidate,
      motivoNaoMatricula: convocacaoMotivoByCandidate[normalizeKey(candidate.candidate)] || null,
      rowModality: classification.modality,
      rowClassification: classification.value,
    }))
  ))

  // Motivos vêm exclusivamente da aba "convocacoes". Cada candidato é contado
  // uma única vez no recorte atual da Lista de Espera.
  const motivosNaoMatricula = filteredWaiting.reduce((acc, candidate) => {
    const motivo = convocacaoMotivoByCandidate[normalizeKey(candidate.candidate)]
    if (motivo === 'Não Compareceu') acc.naoCompareceu += 1
    if (motivo === 'Matrícula Indeferida') acc.matriculaIndeferida += 1
    return acc
  }, { naoCompareceu: 0, matriculaIndeferida: 0 })
  const visibleWaiting = expandedWaiting.slice(0, 250)

  // Evolução acumulada respeitando o recorte atual.
  // O ponto da Chamada Regular deve começar pela ocupação REGULAR do recorte,
  // e somente as chamadas posteriores acrescentam novas ocupações.
  // `totals.occupied` representa a ocupação ATUAL e, portanto, não pode ser
  // usada como ponto inicial do gráfico, pois já inclui as chamadas posteriores.
  const regularCall = filteredCalls.find((call) => call.name === 'Chamada Regular') || { vacancies: 0, occupied: 0, remaining: 0 }
  const processCapacity = regularCall.vacancies || 1
  const regularOccupied = Math.min(
    processCapacity,
    filteredCourses.reduce((sum, course) => sum + (
      course.occupied -
      course.firstCallOccupied -
      course.secondCallOccupied -
      course.thirdCallOccupied
    ), 0)
  )

  let cumulativeOccupied = regularOccupied
  const occupationSnapshots = filteredCalls.map((call) => {
    const isRegular = call.name === 'Chamada Regular'
    const newlyOccupied = isRegular ? 0 : Math.max(0, toNumber(call.occupied))

    if (!isRegular) {
      cumulativeOccupied = Math.min(processCapacity, cumulativeOccupied + newlyOccupied)
    }

    const remaining = Math.max(0, processCapacity - cumulativeOccupied)
    const rate = processCapacity ? (cumulativeOccupied / processCapacity) * 100 : 0

    return {
      ...call,
      occupied: cumulativeOccupied,
      remaining,
      rate,
    }
  })
  // KPI = situação atual do recorte, nunca uma taxa acumulada acima da
  // capacidade original. O valor vem diretamente dos cursos filtrados.
  const currentOccupied = Math.min(totals.vacancies, totals.occupied)
  const currentRemaining = Math.max(0, totals.vacancies - currentOccupied)
  const formatRate = (value) => `${value.toFixed(1).replace('.', ',')}%`
  


  return (
    <div className={`app-shell ${theme === 'dark' ? 'theme-dark' : 'theme-light'}`}>
      <aside className="sidebar">
        <div className="brand institutional-sidebar-brand"><img className="sidebar-linear-logo" src={`${import.meta.env.BASE_URL}logo-ufpi.png`} alt="Logo linear da Universidade Federal do Piauí" /><div className="sidebar-institution-copy"><strong>UFPI · PREG</strong></div></div>
        <nav className="navigation">
          <button className={`navigation-item ${activePage === 'overview' ? 'active' : ''}`} onClick={() => setActivePage('overview')}><UiIcon name="dashboard" size={16} /> <span>Visão geral</span></button>
          <button className={`navigation-item ${activePage === 'waiting' ? 'active' : ''}`} onClick={() => setActivePage('waiting')}><UiIcon name="list" size={16} /> <span>Lista de espera</span></button>
        </nav>
      </aside>

      <main className="main-content">
                {activePage === 'waiting' ? (
          <section className="waiting-page">
            <section className="page-heading compact-heading"><div><span className="eyebrow">CONSULTA INDIVIDUAL</span><h1>Lista de Espera</h1><p>Consulte a classificação publicada oficialmente pela UFPI, organizada por curso, campus e modalidade.</p></div><div className="source-card"><span>Fonte oficial</span><strong>Base publicada pela UFPI</strong></div></section>
            <section className="waiting-kpis" aria-label="Indicadores da lista de espera"><article className="kpi-card"><span className="kpi-label"><i className="data-icon candidates-icon" aria-hidden="true">◎</i> Manifestaram interesse</span><strong>{number(waitingStats.total)}</strong><small>Lista de espera ativa</small></article><article className="kpi-card"><span className="kpi-label"><i className="data-icon waiting-icon" aria-hidden="true">◷</i> Aguardando convocação</span><strong>{number(waitingStats.waiting)}</strong><small>Classificados sem chamada registrada</small></article><article className="kpi-card highlight"><span className="kpi-label"><i className="data-icon called-icon" aria-hidden="true">✓</i> Convocados</span><strong>{number(waitingStats.called)}</strong><small>Chamada registrada na base oficial</small></article></section>
            <section className="panel waiting-panel"><div className="panel-heading"><div><span className="eyebrow">BUSCA E FILTROS</span><h2>Localize seu registro</h2><p className="panel-caption">Os nomes são exibidos conforme a publicação oficial.</p></div></div><div className="waiting-filters"><label className="filter-control search-filter"><span>Nome do candidato</span><input value={waitingSearch} onChange={(event) => setWaitingSearch(event.target.value)} placeholder="Digite o nome" aria-label="Buscar por nome do candidato" /></label><label className="filter-control"><span>Curso</span><select value={waitingCourse} onChange={(event) => setWaitingCourse(event.target.value)}><option>Todos</option>{[...new Set(waitingList.map((item) => item.course))].map((item) => <option key={item}>{item}</option>)}</select></label><label className="filter-control"><span>Campus/Centro</span><select value={waitingCampus} onChange={(event) => setWaitingCampus(event.target.value)}><option>Todos</option>{[...new Set(waitingList.map((item) => item.campus))].map((item) => <option key={item}>{item}</option>)}</select></label><label className="filter-control"><span>Modalidade de concorrência</span><select value={waitingModality} onChange={(event) => setWaitingModality(event.target.value)}><option>Todos</option>{[...new Set(waitingList.map((item) => item.modality))].map((item) => <option key={item}>{item}</option>)}</select></label><label className="filter-control"><span>Modalidade de convocação</span><select value={waitingCalledModality} onChange={(event) => setWaitingCalledModality(event.target.value)}><option>Todos</option>{[...new Set(waitingList.map((item) => item.calledModality))].map((item) => <option key={item}>{item}</option>)}</select></label><label className="filter-control"><span>Situação</span><select value={waitingStatus} onChange={(event) => setWaitingStatus(event.target.value)}><option>Todos</option><option>Aguardando convocação</option><option>Convocado</option></select></label>{(waitingSearch || waitingCourse !== 'Todos' || waitingCampus !== 'Todos' || waitingModality !== 'Todos' || waitingCalledModality !== 'Todos' || waitingStatus !== 'Todos') && <button className="clear-button" type="button" onClick={() => { setWaitingSearch(''); setWaitingCourse('Todos'); setWaitingCampus('Todos'); setWaitingModality('Todos'); setWaitingCalledModality('Todos'); setWaitingStatus('Todos') }} aria-label="Limpar filtros da lista de espera">Limpar filtros</button>}</div></section>
            <section className="panel waiting-panel"><div className="panel-heading"><div><span className="eyebrow">LISTA PUBLICADA</span><h2>Classificação dos candidatos</h2></div><span className="panel-caption">{number(filteredWaiting.length)} candidatos · {number(expandedWaiting.length)} registros por modalidade</span></div><div className="table-container waiting-table-container">{filteredWaiting.length === 0 ? <div className="empty-state" role="status">Nenhum candidato encontrado com os filtros atuais.</div> : <table key={`${waitingSearch}|${waitingCourse}|${waitingCampus}|${waitingModality}|${waitingCalledModality}|${waitingStatus}`} className="waiting-table"><thead><tr><th>Nome</th><th>Curso</th><th>Campus/Centro</th><th>Modalidade de convocação</th><th>Modalidade de concorrência</th><th>Classificação</th><th>Nota</th><th>Situação</th></tr></thead><tbody>{visibleWaiting.map((candidate) => <tr key={`${candidate.registration}-${candidate.course}-${candidate.rowModality}`}><td><strong>{candidate.candidate}</strong><small>Inscrição: {candidate.registration}</small></td><td>{candidate.courseShort}</td><td>{candidate.campus}</td><td><strong className="called-modality">{candidate.calledModality}</strong></td><td>{candidate.rowModality}</td><td>{number(candidate.rowClassification)}º</td><td>{candidate.score == null ? '—' : candidate.score.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td><td><span className={`waiting-status ${candidate.status === 'Convocado' ? 'called' : 'pending'}`}>{candidate.status}</span></td></tr>)}</tbody></table>} </div>{expandedWaiting.length > 250 && <div className="waiting-limit">Exibindo os primeiros 250 registros por modalidade. Use os filtros para localizar um candidato específico.</div>}<div className="modality-footer"><span>.</span></div></section>
          </section>
        ) : (
          <>
        <section className="page-heading"><div className="institutional-lockup"><img src={`${import.meta.env.BASE_URL}logo-ufpi.png`} alt="Brasão da Universidade Federal do Piauí" /><div><span className="institution-name">UNIVERSIDADE FEDERAL DO PIAUÍ</span><span className="institution-unit">PRÓ-REITORIA DE ENSINO DE GRADUAÇÃO</span><h1>Acompanhamento da lista de espera</h1><p>Monitoramento institucional de vagas, ocupações e convocações.</p></div></div><div className="institutional-actions"><button className="theme-toggle" type="button" onClick={() => setTheme((current) => current === 'light' ? 'dark' : 'light')} aria-label={theme === 'light' ? 'Ativar tema escuro' : 'Ativar tema claro'} title={theme === 'light' ? 'Tema escuro' : 'Tema claro'}><span aria-hidden="true">{theme === 'light' ? '☾' : '☀'}</span><strong>{theme === 'light' ? 'Escuro' : 'Claro'}</strong></button></div></section>

        <section className="filters" aria-label="Filtros da visão geral">
          <label className="filter-control">
            <span>Campus/Centro</span>
            <select value={campus} onChange={(event) => setCampus(event.target.value)} aria-label="Filtrar por campus/centro">
              <option>Todos</option>
              <option>CAFS</option>
              <option>CCN</option>
              <option>CCE</option>
            </select>
          </label>
          <label className="filter-control search-filter">
            <span>Curso</span>
            <input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Todos os cursos" aria-label="Filtrar por curso" />
          </label>
          {(search || campus !== 'Todos') && <button className="clear-button" type="button" onClick={() => { setSearch(''); setCampus('Todos') }}>Limpar filtros</button>}
          {hasOverviewFilter && <div className="active-filter-summary" aria-live="polite"><span>Filtro ativo</span>{campus !== 'Todos' && <b>Campus/Centro: {campus}</b>}{search && <b>Curso: {search}</b>}</div>}
        </section>

        <section className="kpi-grid" aria-label="Indicadores principais">
          <article className="kpi-card overview-kpi">
            <div className="kpi-icon kpi-icon-blue"><UiIcon name="vacancies" /></div>
            <div><span className="kpi-label">Vagas ofertadas</span><strong>{number(totals.vacancies)}</strong><small>Chamada regular</small></div>
          </article>
          <article className="kpi-card overview-kpi">
            <div className="kpi-icon kpi-icon-green"><UiIcon name="occupied" /></div>
            <div><span className="kpi-label">Vagas ocupadas</span><strong>{number(currentOccupied)}</strong><small>{percent(currentOccupied, totals.vacancies)} da oferta</small></div>
          </article>
          <article className="kpi-card overview-kpi">
            <div className="kpi-icon kpi-icon-amber"><UiIcon name="remaining" /></div>
            <div><span className="kpi-label">Vagas remanescentes</span><strong>{number(currentRemaining)}</strong><small>{percent(currentRemaining, totals.vacancies)} da oferta</small></div>
          </article>
          <article className="kpi-card overview-kpi">
            <div className="kpi-icon kpi-icon-slate"><UiIcon name="waiting" /></div>
            <div><span className="kpi-label">Lista de espera</span><strong>{number(totals.waitingTotal)}</strong><small className="waiting-kpi-detail">{number(totals.firstCall)} convocados na 1ª · {number(totals.secondCall)} na 2ª · {number(totals.thirdCall)} na 3ª chamada</small></div>
          </article>
        </section>

        <section className="overview-layout">
          <div className="overview-layout-column overview-layout-left overview-grid overview-analysis-grid">
          <article className="panel demand-panel">
            <div className="panel-heading">
              <div><span className="eyebrow">DISPONIBILIDADE POR CURSO</span><h2>Onde ainda existem vagas?</h2></div>
              <span className="panel-caption">Situação atual</span>
            </div>
            <div className="bars">
              {filteredCourses.length === 0 ? <div className="empty-state" role="status">Nenhum curso encontrado com os filtros atuais.</div> : filteredCourses.map((course) => {
                // `course.occupied` já contém a ocupação da Chamada Regular
                // somada às ocupações efetivamente registradas nas chamadas posteriores.
                // Não somar firstCallOccupied/secondCallOccupied/thirdCallOccupied novamente aqui.
                const courseCurrentOccupied = course.occupied
                const courseCurrentRemaining = Math.max(0, course.vacancies - courseCurrentOccupied)
                const occupiedRate = course.vacancies ? (courseCurrentOccupied / course.vacancies) * 100 : 0
                const remainingRate = course.vacancies ? (courseCurrentRemaining / course.vacancies) * 100 : 0
                return <div className="bar-row" key={`${course.campus}||${course.name}`}>
                  <div><span>{course.name}</span><b>{number(courseCurrentRemaining)} vagas disponíveis</b></div>
                  <div className="bar-track"><i style={{ width: `${Math.min(occupiedRate, 100)}%` }} /></div>
                  <div className="bar-meta"><span>{number(courseCurrentOccupied)} ocupadas</span><strong>{remainingRate.toFixed(1).replace('.', ',')}% livres</strong></div>
                </div>
              })}
            </div>
            <div className="demand-legend"><span><i className="demand-occupied-key" /> Ocupadas</span><span><i className="demand-empty-key" /> Disponíveis</span></div>
          </article>

          <section className="panel cumulative-panel process-chart-panel overview-process-panel">
          <div className="panel-heading process-heading">
            <div>
              <span className="eyebrow">EVOLUÇÃO DO PROCESSO</span>
              <h2>Ocupação acumulada por chamada</h2>
              {hasOverviewFilter && <p className="panel-caption">Evolução recalculada para o recorte selecionado.</p>}
            </div>
          </div>

          <CumulativeOccupancyChart
            snapshots={occupationSnapshots}
            capacity={processCapacity}
            formatNumber={number}
            formatRate={formatRate}
          />
        

          
        </section>
          </div>
          <div className="overview-layout-column overview-layout-right overview-grid overview-analysis-grid">
          <section className="panel modality-panel overview-modality-panel">
            <div className="panel-heading">
              <div><span className="eyebrow">DISTRIBUIÇÃO DAS VAGAS</span><h2>Ocupação por modalidade</h2></div>
              {hasOverviewFilter && <span className="modality-source"><i className="source-icon" aria-hidden="true">✓</i> Recorte por curso/campus</span>}
            </div>
            <div className="table-container modality-table-container">
              <table className="modality-table">
                <thead><tr><th>Modalidade</th><th>Vagas<br />ofertadas</th><th>Vagas<br />ocupadas</th><th>Vagas<br />remanescentes</th><th>Progresso<br />da ocupação</th></tr></thead>
                <tbody>{(hasOverviewFilter ? filteredModalityRows : modalityRows).map((row) => {
                  const rate = row.vacancies ? (row.occupied / row.vacancies) * 100 : 0
                  return <tr key={row.name}>
                    <td><strong>{row.name}</strong></td>
                    <td>{number(row.vacancies)}</td>
                    <td>{number(row.occupied)}</td>
                    <td>{number(row.remaining)}</td>
                    <td><div className="modality-progress">
                      <div className="modality-progress-label">
                        <span>{rate.toFixed(1).replace('.', ',')}%</span>
                        <small>{number(row.occupied)} de {number(row.vacancies)}</small>
                      </div>
                      <div className="modality-track"><i style={{ width: `${Math.min(rate, 100)}%` }} /></div>
                    </div></td>
                  </tr>
                })}</tbody>
              </table>
            </div>
          </section>

          <section className="panel table-panel"><div className="panel-heading"><div><span className="eyebrow">VAGAS POR CURSO</span><h2>Distribuição das vagas e convocações</h2></div></div><div className="table-container"><table><thead><tr><th>Curso</th><th>Campus</th><th>Vagas</th><th>Ocupadas</th><th>Remanescentes</th><th>Convocados<br />1ª chamada</th><th>Convocados<br />2ª chamada</th><th>Convocados<br />3ª chamada</th></tr></thead><tbody>{filteredCourses.map((course) => { const courseCurrentOccupied = course.occupied; const courseCurrentRemaining = Math.max(0, course.vacancies - courseCurrentOccupied); return <tr key={`${course.campus}||${course.name}`}><td><strong>{course.name}</strong></td><td>{course.campus}</td><td>{number(course.vacancies)}</td><td>{number(courseCurrentOccupied)}</td><td>{number(courseCurrentRemaining)}</td><td>{number(course.firstCall)}</td><td>{number(course.secondCall)}</td><td>{number(course.thirdCall)}</td></tr> })}</tbody></table></div></section>
          </div>
        </section>

        

        <footer>UFPI · PREG · 2026</footer>
        </>)}
      </main>
    </div>
  )
}

export default App
