import { useState, useEffect } from 'react'
import ExcelJS from 'exceljs'
import { checkBackendHealth } from '../api'

function AutoDB() {
  const [shift, setShift] = useState('currentDay')
  const [shiftDate, setShiftDate] = useState('')
  const [dbFrom, setDbFrom] = useState('')
  const [dbTo, setDbTo] = useState('')
  const [inzhener, setInzhener] = useState('')
  const [tex, setTex] = useState('')
  const [loading, setLoading] = useState(false)
  const [status, setStatus] = useState('')

  const [techServiceList, setTechServiceList] = useState([])
  const [wcsEngineersList, setWcsEngineersList] = useState([])
  const [showAddTech, setShowAddTech] = useState(false)
  const [showAddEngineer, setShowAddEngineer] = useState(false)
  const [newTechName, setNewTechName] = useState('')
  const [newEngineerName, setNewEngineerName] = useState('')

  const [backendAvailable, setBackendAvailable] = useState(null)

  // Распарсенные данные, которые можно править прямо в проекте
  const [excelData, setExcelData] = useState([])
  const [shipRows, setShipRows] = useState([])
  const [recvRows, setRecvRows] = useState([])
  const [palletsInput, setPalletsInput] = useState('')
  const [palletStats, setPalletStats] = useState({ paired: 0, unpaired: 0, total: 0 })

  const getApiBase = () => 'http://172.30.1.249:3000'

  const loadSavedData = () => {
    try {
      const savedTech = localStorage.getItem('techServiceList')
      const savedEng = localStorage.getItem('wcsEngineersList')
      setTechServiceList(savedTech ? JSON.parse(savedTech) : [{ id: 1, name: 'Лёша, Владимир' }])
      setWcsEngineersList(savedEng ? JSON.parse(savedEng) : [{ id: 1, name: 'Aser' }])

      const savedInzhener = localStorage.getItem('dbInzhener') || localStorage.getItem('excelInzhener')
      const savedTex = localStorage.getItem('dbTex') || localStorage.getItem('excelTex')
      if (savedInzhener) setInzhener(savedInzhener)
      if (savedTex) setTex(savedTex)

      const savedExcelData = localStorage.getItem('dbExcelData')
      const savedShipRows = localStorage.getItem('dbShipmentRows')
      const savedRecvRows = localStorage.getItem('dbReceiptRows')
      const savedPalletsInput = localStorage.getItem('dbShiftPalletsInput')
      if (savedExcelData) setExcelData(JSON.parse(savedExcelData))
      if (savedShipRows) setShipRows(JSON.parse(savedShipRows))
      if (savedRecvRows) setRecvRows(JSON.parse(savedRecvRows))
      if (savedPalletsInput) setPalletsInput(savedPalletsInput)
    } catch (err) {
      console.error('Error loading saved AutoDB data:', err)
    }
  }

  const recheckBackend = async () => {
    setBackendAvailable(null)
    const ok = await checkBackendHealth()
    setBackendAvailable(ok)
  }

  useEffect(() => {
    const now = new Date()
    const hour = now.getHours()
    const autoShift = (hour >= 8 && hour < 20) ? 'currentDay' : 'currentNight'
    setShift(autoShift)
    updateDates(autoShift)
    loadSavedData()
    recheckBackend()
  }, [])

  useEffect(() => {
    setPalletStats(getPalletStats(palletsInput))
  }, [palletsInput])

  const handleShiftChange = (e) => {
    setShift(e.target.value)
    updateDates(e.target.value)
  }

  const updateDates = (shiftType) => {
    const now = new Date()
    let from, to

    switch (shiftType) {
      case 'currentDay':
        from = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 8, 0, 0)
        to = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 20, 0, 0)
        break
      case 'currentNight':
        from = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 20, 0, 0)
        to = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1, 8, 0, 0)
        break
      case 'lastDay':
        from = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1, 8, 0, 0)
        to = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1, 20, 0, 0)
        break
      case 'lastNight':
        from = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1, 20, 0, 0)
        to = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 8, 0, 0)
        break
    }

    setDbFrom(formatDateTimeLocal(from))
    setDbTo(formatDateTimeLocal(to))
    setShiftDate(formatDate(from))
  }

  const formatDateTimeLocal = (date) => {
    const offset = date.getTimezoneOffset() * 60000
    const localISOTime = (new Date(date - offset)).toISOString().slice(0, 16)
    return localISOTime
  }

  const formatDate = (date) => {
    const year = date.getFullYear()
    const month = String(date.getMonth() + 1).padStart(2, '0')
    const day = String(date.getDate()).padStart(2, '0')
    return `${year}-${month}-${day}`
  }

  const handleDateChange = (e) => {
    setShiftDate(e.target.value)
    if (shift === 'currentDay' || shift === 'lastDay') {
      const date = new Date(e.target.value + 'T08:00')
      const to = new Date(e.target.value + 'T20:00')
      setDbFrom(formatDateTimeLocal(date))
      setDbTo(formatDateTimeLocal(to))
    } else {
      const date = new Date(e.target.value + 'T20:00')
      const nextDay = new Date(date)
      nextDay.setDate(nextDay.getDate() + 1)
      nextDay.setHours(8, 0, 0)
      setDbFrom(formatDateTimeLocal(date))
      setDbTo(formatDateTimeLocal(nextDay))
    }
  }

  const handleAddTech = () => {
    if (!newTechName.trim()) return
    const next = [...techServiceList, { id: Date.now() + Math.random(), name: newTechName.trim() }]
    setTechServiceList(next)
    localStorage.setItem('techServiceList', JSON.stringify(next))
    setTex(newTechName.trim())
    setNewTechName('')
    setShowAddTech(false)
  }

  const handleAddEngineer = () => {
    if (!newEngineerName.trim()) return
    const next = [...wcsEngineersList, { id: Date.now() + Math.random(), name: newEngineerName.trim() }]
    setWcsEngineersList(next)
    localStorage.setItem('wcsEngineersList', JSON.stringify(next))
    setInzhener(newEngineerName.trim())
    setNewEngineerName('')
    setShowAddEngineer(false)
  }

  const fetchJSON = async (url) => {
    const response = await fetch(url, {
      headers: { 'Content-Type': 'application/json' }
    })
    if (!response.ok) throw new Error(`HTTP ${response.status}`)
    return response.json()
  }

  const extractHMS = (datetimeStr) => {
    if (!datetimeStr) return ''
    const match = String(datetimeStr).match(/\d{2}:\d{2}:\d{2}/)
    return match ? match[0] : ''
  }

  const parseHMS = (hms) => {
    if (!hms) return null
    const parts = hms.split(':')
    if (parts.length < 2) return null
    const h = parseInt(parts[0])
    const m = parseInt(parts[1])
    const s = parseInt(parts[2] || 0)
    if (isNaN(h) || isNaN(m) || isNaN(s)) return null
    return h * 3600 + m * 60 + s
  }

  const calcTotalMinutes = (startHMS, endHMS) => {
    const startSec = parseHMS(startHMS)
    const endSec = parseHMS(endHMS)
    if (startSec === null || endSec === null) return ''
    let d = endSec - startSec
    if (d < 0) d += 86400
    return Math.round(d / 60)
  }

  const getPalletStats = (input) => {
    const lines = (input || '').trim().split('\n').filter(line => line.trim())
    const pallets = {}
    for (const line of lines) {
      const trimmed = line.trim()
      if (trimmed.length < 3) continue
      const type = trimmed.slice(-2)
      const id = trimmed.slice(0, -2)
      if (!pallets[id]) pallets[id] = { d1: false, d2: false }
      if (type === 'D1') pallets[id].d1 = true
      if (type === 'D2') pallets[id].d2 = true
    }
    let paired = 0
    let unpaired = 0
    for (const id in pallets) {
      if (pallets[id].d1 && pallets[id].d2) paired++
      else unpaired++
    }
    return { paired, unpaired, total: paired * 2 + unpaired }
  }

  const loadFromDatabase = async () => {
    if (!dbFrom || !dbTo) {
      setStatus('❌ Укажите начало и окончание смены')
      return false
    }

    const base = getApiBase()
    setLoading(true)
    setStatus('⏳ Загрузка из БД...')

    try {
      const qs = `from=${encodeURIComponent(dbFrom)}&to=${encodeURIComponent(dbTo)}`
      const [faultRes, shipRes, recvRes, palletRes] = await Promise.all([
        fetchJSON(`${base}/api/fault-report?${qs}`),
        fetchJSON(`${base}/api/shipment?${qs}`),
        fetchJSON(`${base}/api/receipt?${qs}`),
        fetchJSON(`${base}/api/pallets?${qs}`)
      ])

      // ---- 1. Ошибки штабелеров ----
      const parsedRows = faultRes.rows
        .filter(r => r.stacker)
        .map(r => ({
          id: Date.now() + Math.random(),
          time: extractHMS(r.beginTime),
          endTime: extractHMS(r.finishTime),
          startISO: (r.beginTime || '').substring(0, 19),
          stacker: r.stacker,
          statuses: Array(10).fill('').map((_, i) => i + 1 === r.stacker ? 'F06' : '')
        }))

      function shiftSortKey(isoStr) {
        const d = new Date(isoStr)
        const h = d.getHours()
        if (h < 8) return d.getTime() + 24 * 3600 * 1000
        return d.getTime()
      }
      parsedRows.sort((a, b) => shiftSortKey(a.startISO) - shiftSortKey(b.startISO))

      const excelData = parsedRows.map(row => ({
        id: row.id,
        time: row.time,
        endTime: row.endTime,
        startISO: row.startISO,
        stacker: row.stacker,
        statuses: row.statuses
      }))

      // ---- 2. Отгрузка ----
      const newShipRows = shipRes.rows.map(r => {
        const startTime = extractHMS(r.start_time)
        const endTime = extractHMS(r.end_time)
        return {
          id: Date.now() + Math.random(),
          kisNum: r.bill_no || '',
          palletQty: r.qty || 0,
          startTime,
          endTime,
          totalTime: calcTotalMinutes(startTime, endTime)
        }
      })

      // ---- 3. Приемка ----
      const newRecvRows = recvRes.rows.map(r => {
        const startTime = extractHMS(r.start_time)
        const endTime = extractHMS(r.end_time)
        return {
          id: Date.now() + Math.random(),
          kisNum: r.bill_no || '',
          palletQty: r.qty || 0,
          startTime,
          endTime,
          totalTime: calcTotalMinutes(startTime, endTime)
        }
      })

      // ---- 4. Паллеты D1/D2 ----
      const palletLines = []
      palletRes.paired.forEach(p => { palletLines.push(p.id + 'D1'); palletLines.push(p.id + 'D2'); })
      palletRes.unpaired.forEach(p => { palletLines.push(p.id + (p.missing === 'D2' ? 'D1' : 'D2')); })
      const shiftPalletsInput = palletLines.join('\n')

      setExcelData(excelData)
      setShipRows(newShipRows)
      setRecvRows(newRecvRows)
      setPalletsInput(shiftPalletsInput)
      setPalletStats(getPalletStats(shiftPalletsInput))

      // Сохраняем в localStorage для использования в ExcelGenerator
      localStorage.setItem('dbExcelData', JSON.stringify(excelData))
      localStorage.setItem('dbShipmentRows', JSON.stringify(newShipRows))
      localStorage.setItem('dbReceiptRows', JSON.stringify(newRecvRows))
      localStorage.setItem('dbShiftPalletsInput', shiftPalletsInput)
      localStorage.setItem('dbInzhener', inzhener || 'Aser')
      localStorage.setItem('dbTex', tex || 'Лёша, Владимир')

      setStatus(`✅ Загружено: ошибок ${excelData.length}, отгрузка ${newShipRows.length}, приемка ${newRecvRows.length}, паллет ${palletRes.totalRows}`)
      return true
    } catch (err) {
      console.error('loadFromDatabase error:', err)
      setStatus('❌ Ошибка: ' + err.message)
      return false
    } finally {
      setLoading(false)
    }
  }

  const handleLoadFromDb = async () => {
    await loadFromDatabase()
  }

  const handleDownload = async () => {
    if (!excelData.length && !shipRows.length && !recvRows.length && !palletsInput.trim()) {
      setStatus('❌ Нет данных для генерации. Сначала загрузите из БД или введите вручную.')
      return
    }

    setLoading(true)
    setStatus('⏳ Генерация Excel файлов...')
    try {
      localStorage.setItem('dbExcelData', JSON.stringify(excelData))
      localStorage.setItem('dbShipmentRows', JSON.stringify(shipRows))
      localStorage.setItem('dbReceiptRows', JSON.stringify(recvRows))
      localStorage.setItem('dbShiftPalletsInput', palletsInput)
      localStorage.setItem('dbInzhener', inzhener || 'Aser')
      localStorage.setItem('dbTex', tex || 'Лёша, Владимир')
      localStorage.setItem('excelInzhener', inzhener || 'Aser')
      localStorage.setItem('excelTex', tex || 'Лёша, Владимир')

      await generateExcelWithDBData(excelData, inzhener, tex)
      await new Promise(r => setTimeout(r, 600))
      await generateShiftReportWithDBData(shipRows, recvRows, palletsInput, excelData, inzhener, tex)

      setStatus('✅ Файлы скачаны!')
      setTimeout(() => setStatus(''), 3000)
    } catch (err) {
      console.error('Error generating Excel:', err)
      setStatus('❌ Ошибка генерации: ' + err.message)
    } finally {
      setLoading(false)
    }
  }

  // ---- Редактирование распарсенных данных ----

  const updateExcelRow = (index, field, value) => {
    setExcelData(prev => {
      const next = [...prev]
      const row = { ...next[index] }
      if (field === 'stacker') {
        const num = parseInt(value) || 1
        row.stacker = num
        row.statuses = Array(10).fill('').map((_, i) => i + 1 === num ? 'F06' : '')
      } else {
        row[field] = value
      }
      next[index] = row
      return next
    })
  }

  const addExcelRow = () => {
    setExcelData(prev => [...prev, {
      id: Date.now() + Math.random(),
      time: '',
      endTime: '',
      startISO: '',
      stacker: 1,
      statuses: Array(10).fill('').map((_, i) => i === 0 ? 'F06' : '')
    }])
  }

  const updateShipRow = (index, field, value) => {
    setShipRows(prev => {
      const next = [...prev]
      const row = { ...next[index], [field]: value }
      if ((field === 'startTime' || field === 'endTime') && row.startTime && row.endTime) {
        row.totalTime = calcTotalMinutes(row.startTime, row.endTime)
      }
      next[index] = row
      return next
    })
  }

  const addShipRow = () => {
    setShipRows(prev => [...prev, { id: Date.now() + Math.random(), kisNum: '', palletQty: '', startTime: '', endTime: '', totalTime: '' }])
  }

  const updateRecvRow = (index, field, value) => {
    setRecvRows(prev => {
      const next = [...prev]
      const row = { ...next[index], [field]: value }
      if ((field === 'startTime' || field === 'endTime') && row.startTime && row.endTime) {
        row.totalTime = calcTotalMinutes(row.startTime, row.endTime)
      }
      next[index] = row
      return next
    })
  }

  const addRecvRow = () => {
    setRecvRows(prev => [...prev, { id: Date.now() + Math.random(), kisNum: '', palletQty: '', startTime: '', endTime: '', totalTime: '' }])
  }

  // ---- Генерация Excel с данными из БД ----

  const generateExcelWithDBData = async (data, inzhenerVal, texVal) => {
    const workbook = new ExcelJS.Workbook()
    const ws = workbook.addWorksheet('Отчёт')

    const RED_HDR  = 'FFFFCCCC'
    const GREY_BG  = 'FFBFBFBF'
    const TIME_BG  = 'FFBDD7EE'
    const RED      = 'FFFF0000'
    const YELLOW   = 'FFFFFF00'
    const GREEN    = 'FF92D050'
    const VOTS_BG  = 'FF00B0F0'
    const ERR_BG   = 'FFFF9999'
    const WHITE    = 'FFFFFFFF'
    const BLACK    = 'FF000000'

    const FONT = 'Calibri'
    const FSIZE = 12
    const BORDER = { top:{style:'thin',color:{argb:BLACK}}, bottom:{style:'thin',color:{argb:BLACK}}, left:{style:'thin',color:{argb:BLACK}}, right:{style:'thin',color:{argb:BLACK}} }

    ws.columns = [
      { key: 'time', width: 14 },
      { key: 's1',   width: 16 },
      { key: 's2',   width: 16 },
      { key: 's3',   width: 16 },
      { key: 's4',   width: 16 },
      { key: 's5',   width: 16 },
      { key: 's6',   width: 16 },
      { key: 's7',   width: 16 },
      { key: 's8',   width: 16 },
      { key: 's9',   width: 16 },
      { key: 's10',  width: 16 },
    ]

    const cellStyle = (bgColor, bold = true) => ({
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } },
      font: { name: FONT, size: FSIZE, bold: bold, color: { argb: BLACK } },
      alignment: { horizontal: 'center', vertical: 'middle', wrapText: false },
      border: BORDER
    })

    const timeStyle = () => ({
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: TIME_BG } },
      font: { name: FONT, size: FSIZE, bold: true, color: { argb: RED } },
      alignment: { horizontal: 'center', vertical: 'middle' },
      border: BORDER
    })

    const headers = ['Время','1-штабелер','2-штабелер','3-штабелер','4-штабелер','5-штабелер','6-штабелер','7-штабелер','8-штабелер','9-штабелер','10-штабелер']
    const headerRow = ws.addRow(headers)
    headerRow.eachCell(cell => { cell.style = cellStyle(RED_HDR, true) })
    headerRow.height = 18

    function calcVotsTime(startHMS, endHMS) {
      const toSec = t => { const p = t.split(':'); return parseInt(p[0])*3600+parseInt(p[1])*60+parseInt(p[2]||0); }
      const fromSec = s => { const h=Math.floor(s/3600),m=Math.floor((s%3600)/60),sec=s%60; return `${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}:${String(sec).padStart(2,'0')}`; }
      let s = toSec(startHMS), e = toSec(endHMS)
      if (e < s) e += 86400
      const diffSec = e - s
      const diffMin = diffSec / 60
      if (diffMin <= 15) {
        return fromSec(e - 120)
      } else {
        return fromSec(s + Math.round(diffSec / 2))
      }
    }

    for (let i = 0; i < data.length; i++) {
      const row = data[i]
      const votsTime = calcVotsTime(row.time, row.endTime)

      const f06Row = ws.addRow([row.time, ...row.statuses.map(s => s === 'F06' ? 'F06' : '')])
      f06Row.height = 18
      f06Row.getCell(1).style = timeStyle()
      row.statuses.forEach((s, idx) => {
        f06Row.getCell(idx + 2).style = cellStyle(s === 'F06' ? RED : GREY_BG, true)
      })

      const texRow = ws.addRow([votsTime, ...row.statuses.map(s => s === 'F06' ? 'тех' : '')])
      texRow.height = 18
      texRow.getCell(1).style = timeStyle()
      row.statuses.forEach((s, idx) => {
        texRow.getCell(idx + 2).style = cellStyle(s === 'F06' ? YELLOW : GREY_BG, true)
      })

      const readyRow = ws.addRow([row.endTime || '', ...row.statuses.map(s => s === 'F06' ? 'готов к работе' : '')])
      readyRow.height = 18
      readyRow.getCell(1).style = timeStyle()
      row.statuses.forEach((s, idx) => {
        readyRow.getCell(idx + 2).style = cellStyle(s === 'F06' ? GREEN : GREY_BG, true)
      })
    }

    const legendStartCol = 13
    const legendData = [
      { color: RED,     text: 'F06',            desc: 'Код ошибки' },
      { color: GREEN,   text: 'готов к работе', desc: 'Время снятия ошибки со штабелера.' },
      { color: VOTS_BG, text: 'ВОТС',           desc: 'Время ожидания технической службы.' },
      { color: ERR_BG,  text: 'В ошибке',       desc: 'Общее время простоя штабелера по причине ошибки.' },
    ]

    ws.getColumn(legendStartCol).width = 18
    ws.getColumn(legendStartCol + 1).width = 48

    legendData.forEach((item, idx) => {
      const wsRow = ws.getRow(idx + 2)
      const isRedLabel = item.text === 'ВОТС' || item.text === 'В ошибке'
      const labelCell = wsRow.getCell(legendStartCol)
      labelCell.value = item.text
      labelCell.style = {
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: isRedLabel ? WHITE : item.color } },
        font: { name: FONT, size: FSIZE, bold: true, color: { argb: isRedLabel ? RED : BLACK } },
        alignment: { horizontal: 'center', vertical: 'middle' },
        border: BORDER
      }
      const descCell = wsRow.getCell(legendStartCol + 1)
      descCell.value = item.desc
      descCell.style = {
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: GREY_BG } },
        font: { name: FONT, size: FSIZE, bold: true, color: { argb: BLACK } },
        alignment: { horizontal: 'left', vertical: 'middle' },
        border: BORDER
      }
    })

    let autoErrorMinutes = 0
    let autoVOTSMinutes = 0
    for (const row of data) {
      if (row.time && row.endTime) {
        const toSec = t => {
          const p = t.split(':')
          return parseInt(p[0]) * 3600 + parseInt(p[1]) * 60 + parseInt(p[2] || 0)
        }
        let diff = toSec(row.endTime) - toSec(row.time)
        if (diff < 0) diff += 24 * 3600
        const diffMinutes = Math.round(diff / 60)
        autoErrorMinutes += diffMinutes

        if (diffMinutes <= 15) {
          autoVOTSMinutes += Math.max(0, diffMinutes - 2)
        } else {
          autoVOTSMinutes += Math.round(diffMinutes / 2)
        }
      }
    }

    const totals = [
      ['В ошибке',    autoErrorMinutes ? `${autoErrorMinutes} мин` : ''],
      ['ВОТС',        autoVOTSMinutes ? `${autoVOTSMinutes} мин` : ''],
      ['Инженер WCS', inzhenerVal],
      ['Тех служба',  texVal],
    ]

    totals.forEach(([label, value]) => {
      const r = ws.addRow([label, value])
      r.height = 18
      r.getCell(1).style = cellStyle(WHITE, true)
      r.getCell(2).style = cellStyle(WHITE, true)
    })

    const buffer = await workbook.xlsx.writeBuffer()
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = 'stackers_report.xlsx'
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }

  const generateShiftReportWithDBData = async (shipRows, recvRows, palletsInput, errorData, inzhenerVal, texVal) => {
    const workbook = new ExcelJS.Workbook()
    const ws = workbook.addWorksheet('Отчёт смены')

    const GREEN_HDR  = 'FF92D050'
    const GREY_HDR   = 'FFD9D9D9'
    const YELLOW_TOT = 'FFFFFF99'
    const BEIGE_HDR  = 'FFFFCCCC'
    const BEIGE_ROW  = 'FFFFCCCC'
    const BLUE_HDR   = 'FFDAE8FC'
    const BLUE_ROW   = 'FFDAE8FC'
    const WHITE      = 'FFFFFFFF'
    const BLACK      = 'FF000000'

    const FONT  = 'Calibri'
    const FSIZE = 12
    const BD    = { top:{style:'thin',color:{argb:BLACK}}, bottom:{style:'thin',color:{argb:BLACK}}, left:{style:'thin',color:{argb:BLACK}}, right:{style:'thin',color:{argb:BLACK}} }

    function cs(bg, bold=true, align='center') {
      return {
        fill: { type:'pattern', pattern:'solid', fgColor:{argb:bg} },
        font: { name:FONT, size:FSIZE, bold, color:{argb:BLACK} },
        alignment: { horizontal:align, vertical:'middle' },
        border: BD
      }
    }

    ws.columns = [
      {width:26},{width:20},{width:16},{width:18},{width:22}
    ]

    function addSectionHeader(title, bgColor, cols=5) {
      const r = ws.addRow([title,'','','',''])
      if (cols > 1) ws.mergeCells(`A${r.number}:${String.fromCharCode(64+cols)}${r.number}`)
      r.getCell(1).style = { ...cs(bgColor, true, 'left') }
      r.height = 20
    }

    function addTableHeader() {
      const r = ws.addRow(['Номер Кис','Количество паллет','Время начала','Время окончания','Общее время (минут)'])
      r.eachCell(c => { c.style = cs(GREY_HDR, true) })
      r.height = 16
    }

    function addDataRows(rows) {
      let totalQty = 0, totalMin = 0
      rows.forEach(row => {
        const diffVal = row.totalTime !== '' ? row.totalTime : ''
        const r = ws.addRow([row.kisNum, row.palletQty, row.startTime, row.endTime, diffVal])
        r.eachCell((c, col) => {
          c.style = cs(WHITE, false, col === 1 ? 'left' : 'center')
        })
        r.height = 16
        const q = parseInt(row.palletQty)
        if (!isNaN(q)) totalQty += q
        if (row.totalTime !== '') totalMin += parseInt(row.totalTime) || 0
      })
      const tr = ws.addRow(['Общее количество', totalQty, '', '', `${totalMin} мин`])
      tr.getCell(1).style = cs(YELLOW_TOT, true, 'left')
      tr.getCell(2).style = cs(YELLOW_TOT, true, 'center')
      tr.getCell(3).style = cs(YELLOW_TOT, false, 'center')
      tr.getCell(4).style = cs(YELLOW_TOT, false, 'center')
      tr.getCell(5).style = cs(YELLOW_TOT, true, 'center')
      tr.height = 16
    }

    // ОТГРУЗКА
    addSectionHeader('Отгрузка', GREEN_HDR)
    addTableHeader()
    addDataRows(shipRows)
    ws.addRow([])

    // ПРИЕМКА
    addSectionHeader('Приемка', GREEN_HDR)
    addTableHeader()
    addDataRows(recvRows)
    ws.addRow([])

    let autoErrCount = 0
    let autoErrTime = 0
    let autoVOTSTime = 0
    if (errorData.length > 0) {
      autoErrCount = errorData.length
      for (const row of errorData) {
        if (row.time && row.endTime) {
          const toSec = t => {
            const p = t.split(':')
            return parseInt(p[0]) * 3600 + parseInt(p[1]) * 60 + parseInt(p[2] || 0)
          }
          let diff = toSec(row.endTime) - toSec(row.time)
          if (diff < 0) diff += 24 * 3600
          const diffMinutes = Math.round(diff / 60)
          autoErrTime += diffMinutes

          if (diffMinutes <= 15) {
            autoVOTSTime += Math.max(0, diffMinutes - 2)
          } else {
            autoVOTSTime += Math.round(diffMinutes / 2)
          }
        }
      }
    }

    // БЛОК 1: ОШИБКА (розовый — 2 колонки)
    addSectionHeader('Ошибка', BEIGE_HDR, 2)

    const errRows = [
      ['Количество ошибок',                autoErrCount],
      ['Время когда Робот был в ошибке',   `${autoErrTime} минут`],
      ['Время ожидания тех службы (ВОТС)', `${autoVOTSTime} минут`],
      ['Имя тех службы',                   texVal],
      ['Инженер WCS',                      inzhenerVal],
    ]
    errRows.forEach(([label, val]) => {
      const r = ws.addRow([label, val])
      r.getCell(1).style = cs(BEIGE_ROW, true, 'left')
      r.getCell(2).style = cs(BEIGE_ROW, true, 'center')
      r.height = 16
    })

    ws.addRow([])

    // Паллеты
    const stats = getPalletStats(palletsInput)
    const { paired, unpaired, total } = stats
    const pairedPct = total > 0 ? Math.round((paired * 2 / total) * 100) + '%' : '0%'
    const unpairedPct = total > 0 ? Math.round((unpaired / total) * 100) + '%' : '0%'

    // БЛОК 2: ПАЛЛЕТЫ (голубой — 3 колонки)
    addSectionHeader('Парные / Непарные паллеты', BLUE_HDR, 3)
    const pallData = [
      { label:'Парные',          qty:paired,   pct:pairedPct,   bg:BLUE_ROW },
      { label:'Не парные',       qty:unpaired, pct:unpairedPct, bg:BLUE_ROW },
      { label:'Общий кол палет', qty:total,    pct:'',          bg:WHITE    },
    ]
    pallData.forEach(({label, qty, pct, bg}) => {
      const r = ws.addRow(pct ? [label, qty, pct] : [label, qty])
      r.getCell(1).style = cs(bg, true, 'left')
      r.getCell(2).style = cs(bg, true, 'center')
      if (pct) r.getCell(3).style = cs(bg, true, 'center')
      r.height = 16
    })

    const buffer = await workbook.xlsx.writeBuffer()
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = 'shift_report.xlsx'
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }

  if (backendAvailable === false) {
    return (
      <div className="card">
        <div className="card-header">
          <div>
            <h2 className="card-title">Автозагрузка отчётов из БД</h2>
            <p className="card-subtitle">Автоматическая загрузка данных из WMS/WCS и генерация Excel отчётов</p>
          </div>
        </div>
        <div style={{ padding: '20px' }}>
          <div style={{ background: 'rgba(239, 68, 68, 0.1)', border: '1px solid rgba(239, 68, 68, 0.4)', borderRadius: '10px', padding: '20px', textAlign: 'center' }}>
            <div style={{ fontSize: '32px', marginBottom: '8px' }}>⚠️</div>
            <h3 style={{ fontSize: '16px', fontWeight: 700, color: '#ef4444', margin: '0 0 8px' }}>Backend не работает</h3>
            <p style={{ color: 'var(--text-secondary)', margin: '0 0 16px' }}>
              Не удалось подключиться к серверу. Эта страница требует запущенный backend с доступом к БД.
            </p>
            <button onClick={recheckBackend} className="btn btn-secondary">Повторить попытку</button>
          </div>
        </div>
      </div>
    )
  }

  if (backendAvailable === null) {
    return (
      <div className="card">
        <div style={{ padding: '40px', textAlign: 'center', color: 'var(--text-secondary)' }}>
          Проверка соединения с backend...
        </div>
      </div>
    )
  }

  return (
    <div className="card">
      <div className="card-header">
        <div>
          <h2 className="card-title">Автозагрузка отчётов из БД</h2>
          <p className="card-subtitle">Автоматическая загрузка данных из WMS/WCS и генерация Excel отчётов</p>
        </div>
      </div>

      <div style={{ padding: '20px' }}>
        <div style={{ background: 'var(--surface)', border: '2px solid var(--primary)', borderRadius: '10px', padding: '16px', marginBottom: '20px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '14px' }}>
            <div style={{ width: '4px', height: '18px', background: 'var(--primary)', borderRadius: '2px' }}></div>
            <h3 style={{ fontSize: '14px', fontWeight: 700, margin: 0, color: 'var(--text-primary)' }}>
              ⚡ Автозагрузка данных из БД (WMS/WCS)
            </h3>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px', marginBottom: '12px' }}>
            <div className="form-group" style={{ margin: 0 }}>
              <label className="form-label">Смена</label>
              <select
                id="shiftSelect"
                value={shift}
                onChange={handleShiftChange}
                className="form-input"
              >
                <option value="currentDay">Текущая дневная (8:00-20:00)</option>
                <option value="currentNight">Текущая ночная (20:00-8:00)</option>
                <option value="lastDay">Прошлая дневная (8:00-20:00)</option>
                <option value="lastNight">Прошлая ночная (20:00-8:00)</option>
              </select>
            </div>
            <div className="form-group" style={{ margin: 0 }}>
              <label className="form-label">Дата</label>
              <input
                id="shiftDate"
                type="date"
                value={shiftDate}
                onChange={handleDateChange}
                className="form-input"
              />
            </div>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px', marginBottom: '12px' }}>
            <div className="form-group" style={{ margin: 0 }}>
              <label className="form-label">Начало смены</label>
              <input
                id="dbFrom"
                type="datetime-local"
                value={dbFrom}
                onChange={(e) => setDbFrom(e.target.value)}
                className="form-input"
              />
            </div>
            <div className="form-group" style={{ margin: 0 }}>
              <label className="form-label">Окончание смены</label>
              <input
                id="dbTo"
                type="datetime-local"
                value={dbTo}
                onChange={(e) => setDbTo(e.target.value)}
                className="form-input"
              />
            </div>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px', marginBottom: '12px' }}>
            <div className="form-group" style={{ margin: 0 }}>
              <label className="form-label">Инженер WCS</label>
              {!showAddEngineer ? (
                <div style={{ display: 'flex', gap: '6px' }}>
                  <select
                    id="dbInzhener"
                    value={inzhener}
                    onChange={(e) => setInzhener(e.target.value)}
                    className="form-input"
                    style={{ flex: 1 }}
                  >
                    <option value="">Выберите из списка</option>
                    {wcsEngineersList.map((w) => (
                      <option key={w.id} value={w.name}>{w.name}</option>
                    ))}
                  </select>
                  <button
                    type="button"
                    onClick={() => setShowAddEngineer(true)}
                    className="btn btn-secondary"
                    style={{ padding: '0 14px', fontSize: '18px', lineHeight: 1 }}
                    title="Добавить нового инженера"
                  >
                    +
                  </button>
                </div>
              ) : (
                <div style={{ display: 'flex', gap: '6px' }}>
                  <input
                    type="text"
                    autoFocus
                    value={newEngineerName}
                    onChange={(e) => setNewEngineerName(e.target.value)}
                    onKeyDown={(e) => { if (e.key === 'Enter') handleAddEngineer() }}
                    className="form-input"
                    style={{ flex: 1 }}
                    placeholder="Новое имя инженера"
                  />
                  <button type="button" onClick={handleAddEngineer} className="btn btn-primary" style={{ padding: '0 14px' }}>
                    ✓
                  </button>
                  <button
                    type="button"
                    onClick={() => { setShowAddEngineer(false); setNewEngineerName('') }}
                    className="btn btn-secondary"
                    style={{ padding: '0 14px' }}
                  >
                    ✕
                  </button>
                </div>
              )}
            </div>
            <div className="form-group" style={{ margin: 0 }}>
              <label className="form-label">Тех служба</label>
              {!showAddTech ? (
                <div style={{ display: 'flex', gap: '6px' }}>
                  <select
                    id="dbTex"
                    value={tex}
                    onChange={(e) => setTex(e.target.value)}
                    className="form-input"
                    style={{ flex: 1 }}
                  >
                    <option value="">Выберите из списка</option>
                    {techServiceList.map((t) => (
                      <option key={t.id} value={t.name}>{t.name}</option>
                    ))}
                  </select>
                  <button
                    type="button"
                    onClick={() => setShowAddTech(true)}
                    className="btn btn-secondary"
                    style={{ padding: '0 14px', fontSize: '18px', lineHeight: 1 }}
                    title="Добавить нового сотрудника"
                  >
                    +
                  </button>
                </div>
              ) : (
                <div style={{ display: 'flex', gap: '6px' }}>
                  <input
                    type="text"
                    autoFocus
                    value={newTechName}
                    onChange={(e) => setNewTechName(e.target.value)}
                    onKeyDown={(e) => { if (e.key === 'Enter') handleAddTech() }}
                    className="form-input"
                    style={{ flex: 1 }}
                    placeholder="Новое имя сотрудника"
                  />
                  <button type="button" onClick={handleAddTech} className="btn btn-primary" style={{ padding: '0 14px' }}>
                    ✓
                  </button>
                  <button
                    type="button"
                    onClick={() => { setShowAddTech(false); setNewTechName('') }}
                    className="btn btn-secondary"
                    style={{ padding: '0 14px' }}
                  >
                    ✕
                  </button>
                </div>
              )}
            </div>
          </div>

          <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap', alignItems: 'center' }}>
            <button
              id="dbLoadBtn"
              onClick={handleLoadFromDb}
              disabled={loading}
              className="btn btn-secondary"
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ marginRight: '6px' }}>
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                <polyline points="7 10 12 15 17 10"/>
                <line x1="12" y1="15" x2="12" y2="3"/>
              </svg>
              Загрузить из БД
            </button>
            <button
              id="dbDownloadBtn"
              onClick={handleDownload}
              disabled={loading}
              className="btn btn-primary"
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ marginRight: '6px' }}>
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                <polyline points="7 10 12 15 17 10"/>
                <line x1="12" y1="15" x2="12" y2="3"/>
              </svg>
              Скачать 2 файла Excel
            </button>
            <span id="dbLoadStatus" style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>{status}</span>
          </div>
        </div>

        {/* ---- Редактирование распарсенных данных ---- */}
        {(excelData.length > 0 || shipRows.length > 0 || recvRows.length > 0 || palletsInput) && (
          <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: '10px', padding: '16px', marginBottom: '20px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '14px' }}>
              <div style={{ width: '4px', height: '18px', background: '#22c55e', borderRadius: '2px' }}></div>
              <h3 style={{ fontSize: '14px', fontWeight: 700, margin: 0, color: 'var(--text-primary)' }}>
                Распарсенные данные (можно редактировать перед скачиванием)
              </h3>
            </div>

            {/* Ошибки штабелеров */}
            <div style={{ marginBottom: '18px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '10px' }}>
                <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--text-primary)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Ошибки штабелеров</span>
                <button onClick={addExcelRow} className="btn btn-secondary" style={{ padding: '3px 12px', fontSize: '12px' }}>+ Добавить</button>
              </div>
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                  <thead style={{ background: 'var(--background)' }}>
                    <tr>
                      <th style={{ padding: '8px', border: '1px solid var(--border)', textAlign: 'left' }}>Время начала</th>
                      <th style={{ padding: '8px', border: '1px solid var(--border)', textAlign: 'left' }}>Время окончания</th>
                      <th style={{ padding: '8px', border: '1px solid var(--border)', textAlign: 'center' }}>Штабелер</th>
                      <th style={{ padding: '8px', border: '1px solid var(--border)', textAlign: 'center', width: '60px' }}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {excelData.map((row, index) => (
                      <tr key={row.id || index}>
                        <td style={{ padding: '6px', border: '1px solid var(--border)' }}>
                          <input type="text" value={row.time} onChange={(e) => updateExcelRow(index, 'time', e.target.value)} className="form-input" placeholder="09:41:05" style={{ fontSize: '12px', padding: '6px' }} />
                        </td>
                        <td style={{ padding: '6px', border: '1px solid var(--border)' }}>
                          <input type="text" value={row.endTime} onChange={(e) => updateExcelRow(index, 'endTime', e.target.value)} className="form-input" placeholder="10:16:30" style={{ fontSize: '12px', padding: '6px' }} />
                        </td>
                        <td style={{ padding: '6px', border: '1px solid var(--border)', textAlign: 'center' }}>
                          <select value={row.stacker} onChange={(e) => updateExcelRow(index, 'stacker', e.target.value)} className="form-input" style={{ fontSize: '12px', padding: '6px' }}>
                            {Array.from({ length: 10 }, (_, i) => i + 1).map(n => (
                              <option key={n} value={n}>{n}</option>
                            ))}
                          </select>
                        </td>
                        <td style={{ padding: '6px', border: '1px solid var(--border)', textAlign: 'center' }}>
                          <button onClick={() => setExcelData(excelData.filter((_, i) => i !== index))} style={{ padding: '6px 10px', background: 'var(--danger,#ef4444)', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer' }}>✕</button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Отгрузка */}
            <div style={{ marginBottom: '18px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '10px' }}>
                <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--text-primary)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Отгрузка</span>
                <button onClick={addShipRow} className="btn btn-secondary" style={{ padding: '3px 12px', fontSize: '12px' }}>+ Добавить</button>
              </div>
              {shipRows.map((row, index) => (
                <div key={row.id} style={{ display: 'grid', gridTemplateColumns: '2fr 1fr 1fr 1fr auto', gap: '8px', marginBottom: '8px', alignItems: 'end' }}>
                  <div className="form-group" style={{ margin: 0 }}>
                    <label className="form-label" style={{ fontSize: '11px' }}>Номер Кис</label>
                    <input type="text" value={row.kisNum} onChange={(e) => updateShipRow(index, 'kisNum', e.target.value)} className="form-input" placeholder="АЛЦ-0114435" style={{ fontSize: '12px', padding: '6px' }} />
                  </div>
                  <div className="form-group" style={{ margin: 0 }}>
                    <label className="form-label" style={{ fontSize: '11px' }}>Кол-во паллет</label>
                    <input type="text" value={row.palletQty} onChange={(e) => updateShipRow(index, 'palletQty', e.target.value)} className="form-input" placeholder="40" style={{ fontSize: '12px', padding: '6px' }} />
                  </div>
                  <div className="form-group" style={{ margin: 0 }}>
                    <label className="form-label" style={{ fontSize: '11px' }}>Время начала</label>
                    <input type="text" value={row.startTime} onChange={(e) => updateShipRow(index, 'startTime', e.target.value)} className="form-input" placeholder="09:41:05" style={{ fontSize: '12px', padding: '6px' }} />
                  </div>
                  <div className="form-group" style={{ margin: 0 }}>
                    <label className="form-label" style={{ fontSize: '11px' }}>Время окончания</label>
                    <input type="text" value={row.endTime} onChange={(e) => updateShipRow(index, 'endTime', e.target.value)} className="form-input" placeholder="10:16:30" style={{ fontSize: '12px', padding: '6px' }} />
                  </div>
                  <button onClick={() => setShipRows(shipRows.filter(r => r.id !== row.id))} style={{ padding: '8px 10px', background: 'var(--danger,#ef4444)', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', height: '36px', marginBottom: 0 }}>✕</button>
                </div>
              ))}
            </div>

            {/* Приемка */}
            <div style={{ marginBottom: '18px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '10px' }}>
                <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--text-primary)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Приемка</span>
                <button onClick={addRecvRow} className="btn btn-secondary" style={{ padding: '3px 12px', fontSize: '12px' }}>+ Добавить</button>
              </div>
              {recvRows.map((row, index) => (
                <div key={row.id} style={{ display: 'grid', gridTemplateColumns: '2fr 1fr 1fr 1fr auto', gap: '8px', marginBottom: '8px', alignItems: 'end' }}>
                  <div className="form-group" style={{ margin: 0 }}>
                    <label className="form-label" style={{ fontSize: '11px' }}>Номер Кис</label>
                    <input type="text" value={row.kisNum} onChange={(e) => updateRecvRow(index, 'kisNum', e.target.value)} className="form-input" placeholder="АЛЦ-0114435" style={{ fontSize: '12px', padding: '6px' }} />
                  </div>
                  <div className="form-group" style={{ margin: 0 }}>
                    <label className="form-label" style={{ fontSize: '11px' }}>Кол-во паллет</label>
                    <input type="text" value={row.palletQty} onChange={(e) => updateRecvRow(index, 'palletQty', e.target.value)} className="form-input" placeholder="40" style={{ fontSize: '12px', padding: '6px' }} />
                  </div>
                  <div className="form-group" style={{ margin: 0 }}>
                    <label className="form-label" style={{ fontSize: '11px' }}>Время начала</label>
                    <input type="text" value={row.startTime} onChange={(e) => updateRecvRow(index, 'startTime', e.target.value)} className="form-input" placeholder="09:41:05" style={{ fontSize: '12px', padding: '6px' }} />
                  </div>
                  <div className="form-group" style={{ margin: 0 }}>
                    <label className="form-label" style={{ fontSize: '11px' }}>Время окончания</label>
                    <input type="text" value={row.endTime} onChange={(e) => updateRecvRow(index, 'endTime', e.target.value)} className="form-input" placeholder="10:16:30" style={{ fontSize: '12px', padding: '6px' }} />
                  </div>
                  <button onClick={() => setRecvRows(recvRows.filter(r => r.id !== row.id))} style={{ padding: '8px 10px', background: 'var(--danger,#ef4444)', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', height: '36px', marginBottom: 0 }}>✕</button>
                </div>
              ))}
            </div>

            {/* Паллеты */}
            <div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '10px' }}>
                <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--text-primary)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Паллеты D1/D2</span>
              </div>
              <div className="form-group">
                <label className="form-label">Список паллет (ID + D1 или ID + D2, каждый с новой строки)</label>
                <textarea value={palletsInput} onChange={(e) => setPalletsInput(e.target.value)} className="form-input" rows="4" placeholder="L01X01Y73Z01D1&#10;L01X01Y73Z01D2&#10;L02X02Y74Z01D1&#10;..."></textarea>
              </div>
              {(palletStats.total > 0 || palletsInput.trim()) && (
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3,1fr)', gap: '10px', marginTop: '10px' }}>
                  <div style={{ background: 'var(--background)', border: '1px solid var(--border)', borderRadius: '8px', padding: '12px', textAlign: 'center' }}>
                    <div style={{ fontSize: '10px', color: 'var(--text-secondary)', marginBottom: '4px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Парные</div>
                    <div style={{ fontSize: '20px', fontWeight: 700, color: '#22c55e' }}>{palletStats.paired}</div>
                  </div>
                  <div style={{ background: 'var(--background)', border: '1px solid var(--border)', borderRadius: '8px', padding: '12px', textAlign: 'center' }}>
                    <div style={{ fontSize: '10px', color: 'var(--text-secondary)', marginBottom: '4px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Непарные</div>
                    <div style={{ fontSize: '20px', fontWeight: 700, color: '#ef4444' }}>{palletStats.unpaired}</div>
                  </div>
                  <div style={{ background: 'var(--background)', border: '1px solid var(--border)', borderRadius: '8px', padding: '12px', textAlign: 'center' }}>
                    <div style={{ fontSize: '10px', color: 'var(--text-secondary)', marginBottom: '4px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Всего</div>
                    <div style={{ fontSize: '20px', fontWeight: 700, color: 'var(--text-primary)' }}>{palletStats.total}</div>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  )
}

export default AutoDB
