import { useState, useEffect } from 'react'
import ExcelJS from 'exceljs'

function ExcelGenerator() {
  const [input, setInput] = useState('')
  const [data, setData] = useState([])
  const [inzhener, setInzhener] = useState('')
  const [tex, setTex] = useState('')
  const [fileName, setFileName] = useState('stackers_report')
  const [shiftFileName, setShiftFileName] = useState('shift_report')
  const [errorCount, setErrorCount] = useState(0)
  const [errorTime, setErrorTime] = useState(0)
  const [votsTime, setVotsTime] = useState(0)
  const [shipRows, setShipRows] = useState([])
  const [recvRows, setRecvRows] = useState([])
  const [shiftPalletsInput, setShiftPalletsInput] = useState('')
  const [shiftPaired, setShiftPaired] = useState(0)
  const [shiftUnpaired, setShiftUnpaired] = useState(0)
  const [shiftTotal, setShiftTotal] = useState(0)
  const [shipmentExcelData, setShipmentExcelData] = useState([])
  const [receiptExcelData, setReceiptExcelData] = useState([])

  // Загрузка данных из localStorage (если были загружены через AutoDB)
  useEffect(() => {
    const excelInzhener = localStorage.getItem('excelInzhener')
    const excelTex = localStorage.getItem('excelTex')
    const dbExcelData = localStorage.getItem('dbExcelData')
    const dbShipmentRows = localStorage.getItem('dbShipmentRows')
    const dbReceiptRows = localStorage.getItem('dbReceiptRows')
    const dbShiftPalletsInput = localStorage.getItem('dbShiftPalletsInput')

    if (excelInzhener) setInzhener(excelInzhener)
    if (excelTex) setTex(excelTex)
    if (dbExcelData) {
      const parsed = JSON.parse(dbExcelData)
      setData(parsed)
      setErrorCount(parsed.length)
    }
    if (dbShipmentRows) setShipRows(JSON.parse(dbShipmentRows))
    if (dbReceiptRows) setRecvRows(JSON.parse(dbReceiptRows))
    if (dbShiftPalletsInput) setShiftPalletsInput(dbShiftPalletsInput)
  }, [])

  const equipmentToStacker = {
    'SC0101': 1,
    'SC0201': 2,
    'SC0301': 3,
    'SC0401': 4,
    'SC0501': 5,
    'SC0601': 6,
    'SC0701': 7,
    'SC0801': 8,
    'SC0901': 9,
    'SC1001': 10
  }

  const parseExcelData = () => {
    const lines = input.trim().split('\n').filter(line => line.trim())
    
    if (lines.length === 0) {
      alert('Пожалуйста, вставьте данные')
      return
    }

    const parsedRows = []
    
    for (const line of lines) {
      const parts = line.trim().split(/\t/)
      if (parts.length < 2) continue

      const equipmentId = parts[0].trim()
      const dateMatches = line.match(/\d{4}\.\d{2}\.\d{2} \d{2}:\d{2}:\d{2}/g)
      if (!dateMatches || dateMatches.length < 2) continue

      const startDateStr = dateMatches[0].trim()
      const endDateStr = dateMatches[1].trim()
      const startTime = startDateStr.split(' ')[1]
      const endTime = endDateStr.split(' ')[1]
      const startISO = startDateStr.replace(/\./g, '-').replace(' ', 'T')

      const stackerNum = equipmentToStacker[equipmentId] || null
      if (!stackerNum) continue

      parsedRows.push({
        startTime,
        endTime,
        startISO,
        stacker: stackerNum
      })
    }

    if (parsedRows.length === 0) {
      alert('Не удалось распознать данные')
      return
    }

    // Сортировка с учётом ночной смены
    function shiftSortKey(isoStr) {
      const d = new Date(isoStr)
      const h = d.getHours()
      if (h < 8) return d.getTime() + 24 * 3600 * 1000
      return d.getTime()
    }
    parsedRows.sort((a, b) => shiftSortKey(a.startISO) - shiftSortKey(b.startISO))

    // Автоматический расчёт статистики
    let totalErrorMinutes = 0
    let totalVOTSMinutes = 0
    
    for (const row of parsedRows) {
      const toSec = t => {
        const p = t.split(':')
        return parseInt(p[0]) * 3600 + parseInt(p[1]) * 60 + parseInt(p[2] || 0)
      }
      let diff = toSec(row.endTime) - toSec(row.startTime)
      if (diff < 0) diff += 24 * 3600
      const diffMinutes = Math.round(diff / 60)
      totalErrorMinutes += diffMinutes

      if (diffMinutes <= 15) {
        totalVOTSMinutes += Math.max(0, diffMinutes - 2)
      } else {
        totalVOTSMinutes += Math.round(diffMinutes / 2)
      }
    }

    setData(parsedRows)
    setErrorCount(parsedRows.length)
    setErrorTime(totalErrorMinutes)
    setVotsTime(totalVOTSMinutes)
  }

  const clearExcelData = () => {
    setInput('')
    setData([])
    setErrorCount(0)
    setErrorTime(0)
    setVotsTime(0)
    setShipRows([])
    setRecvRows([])
    setShiftPalletsInput('')
    setShiftPaired(0)
    setShiftUnpaired(0)
    setShiftTotal(0)
  }

  const addShipmentRow = () => {
    setShipRows([...shipRows, { id: Date.now(), kisNum: '', palletQty: '', startTime: '', endTime: '', totalTime: '' }])
  }

  const addReceiptRow = () => {
    setRecvRows([...recvRows, { id: Date.now(), kisNum: '', palletQty: '', startTime: '', endTime: '', totalTime: '' }])
  }

  const parseShipmentExcel = async (e) => {
    const file = e.target.files[0]
    if (!file) return

    try {
      const reader = new FileReader()
      reader.onload = async (event) => {
        const buffer = event.target.result
        const workbook = new ExcelJS.Workbook()
        await workbook.xlsx.load(buffer)
        
        const worksheet = workbook.worksheets[0]
        const data = []
        
        const headers = []
        worksheet.getRow(1).eachCell(cell => headers.push(cell.value))
        
        const startTimeIdx = headers.findIndex(h => h && h.toString().includes('Время начала'))
        const endTimeIdx = headers.findIndex(h => h && h.toString().includes('Время окончания'))
        const machineNumIdx = headers.findIndex(h => h && h.toString().includes('Номер машины'))
        const notesIdx = headers.findIndex(h => h && h.toString().includes('Примечания'))
        
        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber === 1) return
          
          const startTime = startTimeIdx >= 0 ? row.getCell(startTimeIdx + 1).value : ''
          const endTime = endTimeIdx >= 0 ? row.getCell(endTimeIdx + 1).value : ''
          const machineNum = machineNumIdx >= 0 ? row.getCell(machineNumIdx + 1).value : ''
          const notes = notesIdx >= 0 ? row.getCell(notesIdx + 1).value : ''
          
          if (startTime || endTime) {
            data.push({ startTime, endTime, machineNum, notes })
          }
        })
        
        setShipmentExcelData(data)
        alert(`Загружено ${data.length} строк. Нажмите "Парсировать данные" для добавления.`)
      }
      reader.readAsArrayBuffer(file)
    } catch (error) {
      alert('Ошибка загрузки файла: ' + error.message)
    }
  }

  const parseReceiptExcel = async (e) => {
    const file = e.target.files[0]
    if (!file) return

    try {
      const reader = new FileReader()
      reader.onload = async (event) => {
        const buffer = event.target.result
        const workbook = new ExcelJS.Workbook()
        await workbook.xlsx.load(buffer)
        
        const worksheet = workbook.worksheets[0]
        const data = []
        
        const headers = []
        worksheet.getRow(1).eachCell(cell => headers.push(cell.value))
        
        const startTimeIdx = headers.findIndex(h => h && h.toString().includes('Время начала'))
        const endTimeIdx = headers.findIndex(h => h && h.toString().includes('Время окончания'))
        const machineNumIdx = headers.findIndex(h => h && h.toString().includes('Номер машины'))
        const notesIdx = headers.findIndex(h => h && h.toString().includes('Примечания'))
        
        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber === 1) return
          
          const startTime = startTimeIdx >= 0 ? row.getCell(startTimeIdx + 1).value : ''
          const endTime = endTimeIdx >= 0 ? row.getCell(endTimeIdx + 1).value : ''
          const machineNum = machineNumIdx >= 0 ? row.getCell(machineNumIdx + 1).value : ''
          const notes = notesIdx >= 0 ? row.getCell(notesIdx + 1).value : ''
          
          if (startTime || endTime) {
            data.push({ startTime, endTime, machineNum, notes })
          }
        })
        
        setReceiptExcelData(data)
        alert(`Загружено ${data.length} строк. Нажмите "Парсировать данные" для добавления.`)
      }
      reader.readAsArrayBuffer(file)
    } catch (error) {
      alert('Ошибка загрузки файла: ' + error.message)
    }
  }

  const formatKIS = (value) => {
    if (!value) return ''
    const str = value.toString().trim()
    if (str.startsWith('ALЦ-') || str.startsWith('алц-')) return str
    if (/^\d+$/.test(str)) return 'ALЦ-' + str
    return str
  }

  const extractTime = (value) => {
    if (!value) return ''
    
    if (value instanceof Date) {
      const h = String(value.getUTCHours()).padStart(2, '0')
      const m = String(value.getUTCMinutes()).padStart(2, '0')
      const s = String(value.getUTCSeconds()).padStart(2, '0')
      return `${h}:${m}:${s}`
    }
    
    const str = value.toString().trim()
    const match = str.match(/\d{2}:\d{2}:\d{2}/)
    return match ? match[0] : str
  }

  const parseShipmentData = () => {
    if (!shipmentExcelData.length) {
      alert('Сначала загрузите Excel файл')
      return
    }
    
    const newRows = []
    for (const row of shipmentExcelData) {
      const kis = formatKIS(row.notes || row.machineNum)
      const startTime = extractTime(row.startTime)
      const endTime = extractTime(row.endTime)
      
      if (kis || startTime) {
        newRows.push({
          id: Date.now() + Math.random(),
          kisNum: kis,
          palletQty: '',
          startTime,
          endTime,
          totalTime: ''
        })
      }
    }
    
    setShipRows(newRows)
    setShipmentExcelData([])
  }

  const parseReceiptData = () => {
    if (!receiptExcelData.length) {
      alert('Сначала загрузите Excel файл')
      return
    }
    
    const newRows = []
    for (const row of receiptExcelData) {
      const kis = formatKIS(row.notes || row.machineNum)
      const startTime = extractTime(row.startTime)
      const endTime = extractTime(row.endTime)
      
      if (kis || startTime) {
        newRows.push({
          id: Date.now() + Math.random(),
          kisNum: kis,
          palletQty: '',
          startTime,
          endTime,
          totalTime: ''
        })
      }
    }
    
    setRecvRows(newRows)
    setReceiptExcelData([])
  }

  const calcShiftPallets = () => {
    const lines = shiftPalletsInput.trim().split('\n').filter(line => line.trim())
    const pallets = {}
    
    for (const line of lines) {
      const trimmed = line.trim()
      if (trimmed.length < 3) continue
      const type = trimmed.slice(-2)
      const id = trimmed.slice(0, -2)
      
      if (!pallets[id]) {
        pallets[id] = { d1: false, d2: false }
      }
      if (type === 'D1') pallets[id].d1 = true
      if (type === 'D2') pallets[id].d2 = true
    }

    let paired = 0
    let unpaired = 0

    for (const id in pallets) {
      const pallet = pallets[id]
      if (pallet.d1 && pallet.d2) {
        paired++
      } else {
        unpaired++
      }
    }

    const total = paired * 2 + unpaired
    setShiftPaired(paired)
    setShiftUnpaired(unpaired)
    setShiftTotal(total)
  }

  const generateExcel = async () => {
    if (data.length === 0) {
      alert('Нет данных для генерации Excel файла')
      return
    }

    // Автоматический расчет времени ошибок и ВОТС из data
    let autoErrorMinutes = 0
    let autoVOTSMinutes = 0
    for (const row of data) {
      if (row.startTime && row.endTime) {
        const toSec = t => {
          const p = t.split(':')
          return parseInt(p[0]) * 3600 + parseInt(p[1]) * 60 + parseInt(p[2] || 0)
        }
        let diff = toSec(row.endTime) - toSec(row.startTime)
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
    const autoErrCount = data.length

    const vOshibkeRaw = errorTime || autoErrorMinutes || 0
    const vOTSRaw = votsTime || autoVOTSMinutes || 0
    const inzhenerVal = inzhener || ''
    const texVal = tex || ''
    const fileNameVal = (fileName.trim() || 'stackers_report') + '.xlsx'

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

    for (let i = 0; i < data.length; i++) {
      const row = data[i]
      const votsTime = calcVotsTime(row.startTime, row.endTime)
      const statuses = Array(10).fill('').map((_, idx) => row.stacker === idx + 1 ? 'F06' : '')

      const f06Row = ws.addRow([row.startTime, ...statuses.map(s => s === 'F06' ? 'F06' : '')])
      f06Row.height = 18
      f06Row.getCell(1).style = timeStyle()
      statuses.forEach((s, idx) => {
        f06Row.getCell(idx + 2).style = cellStyle(s === 'F06' ? RED : GREY_BG, true)
      })

      const texRow = ws.addRow([votsTime, ...statuses.map(s => s === 'F06' ? 'тех' : '')])
      texRow.height = 18
      texRow.getCell(1).style = timeStyle()
      statuses.forEach((s, idx) => {
        texRow.getCell(idx + 2).style = cellStyle(s === 'F06' ? YELLOW : GREY_BG, true)
      })

      const readyRow = ws.addRow([row.endTime || '', ...statuses.map(s => s === 'F06' ? 'готов к работе' : '')])
      readyRow.height = 18
      readyRow.getCell(1).style = timeStyle()
      statuses.forEach((s, idx) => {
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

    const totals = [
      ['В ошибке',    vOshibkeRaw ? `${vOshibkeRaw} мин` : ''],
      ['ВОТС',        vOTSRaw ? `${vOTSRaw} мин` : ''],
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
    a.download = fileNameVal
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }

  const generateShiftReport = async () => {
    // Автоматический расчет из data если поля пустые
    let autoErrCount = 0
    let autoErrTime = 0
    let autoVOTSTime = 0
    if (data.length > 0) {
      autoErrCount = data.length
      for (const row of data) {
        if (row.startTime && row.endTime) {
          const toSec = t => {
            const p = t.split(':')
            return parseInt(p[0]) * 3600 + parseInt(p[1]) * 60 + parseInt(p[2] || 0)
          }
          let diff = toSec(row.endTime) - toSec(row.startTime)
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

    const errCount = errorCount || autoErrCount || ''
    const errTimeVal = errorTime || autoErrTime || ''
    const vots = votsTime || autoVOTSTime || ''
    const techName = tex || ''
    const wcsName = inzhener || ''
    const fileNameVal = (shiftFileName.trim() || 'shift_report') + '.xlsx'

    // Паллеты
    calcShiftPallets()
    const paired = shiftPaired || 0
    const unpaired = shiftUnpaired || 0
    const total = shiftTotal || 0
    const pairedPct = total > 0 ? Math.round((paired * 2 / total) * 100) + '%' : '0%'
    const unpairedPct = total > 0 ? Math.round((unpaired / total) * 100) + '%' : '0%'

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

    // БЛОК 1: ОШИБКА (розовый — 2 колонки)
    addSectionHeader('Ошибка', BEIGE_HDR, 2)

    const errRows = [
      ['Количество ошибок',                errCount],
      ['Время когда Робот был в ошибке',   `${errTimeVal} минут`],
      ['Время ожидания тех службы (ВОТС)', `${vots} минут`],
      ['Имя тех службы',                   techName],
      ['Инженер WCS',                      wcsName],
    ]
    const EMPTY = { font:{name:FONT,size:FSIZE}, alignment:{horizontal:'center',vertical:'middle'} }
    errRows.forEach(([label, val]) => {
      const r = ws.addRow([label, val])
      r.getCell(1).style = cs(BEIGE_ROW, true, 'left')
      r.getCell(2).style = cs(BEIGE_ROW, true, 'center')
      r.height = 16
    })

    ws.addRow([])

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
    a.download = fileNameVal
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }

  const generateBothReports = async () => {
    await generateExcel()
    await new Promise(r => setTimeout(r, 600))
    await generateShiftReport()
  }

  return (
    <div className="card">
      <div className="card-header">
        <div>
          <h2 className="card-title">Генератор Excel отчётов</h2>
          <p className="card-subtitle">Заполните данные и скачайте нужные файлы</p>
        </div>
      </div>

      <div style={{ padding: '20px' }}>
        {/* Общие данные */}
        <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: '10px', padding: '16px', marginBottom: '20px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '14px' }}>
            <div style={{ width: '4px', height: '18px', background: 'var(--primary)', borderRadius: '2px' }}></div>
            <h3 style={{ fontSize: '14px', fontWeight: 700, margin: 0, color: 'var(--text-primary)' }}>Общие данные</h3>
            <span style={{ fontSize: '11px', color: 'var(--text-secondary)', marginLeft: '4px' }}>используются в обоих файлах</span>
          </div>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
            <div className="form-group" style={{ margin: 0 }}>
              <label className="form-label">Инженер WCS</label>
              <input type="text" value={inzhener} onChange={(e) => setInzhener(e.target.value)} className="form-input" placeholder="Aser" />
            </div>
            <div className="form-group" style={{ margin: 0 }}>
              <label className="form-label">Тех служба</label>
              <input type="text" value={tex} onChange={(e) => setTex(e.target.value)} className="form-input" placeholder="Лёша, Владимир" />
            </div>
          </div>
        </div>

        {/* Отчёт по ошибкам */}
        <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: '10px', padding: '16px', marginBottom: '20px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '14px' }}>
            <div style={{ width: '4px', height: '18px', background: '#ef4444', borderRadius: '2px' }}></div>
            <h3 style={{ fontSize: '14px', fontWeight: 700, margin: 0, color: 'var(--text-primary)' }}>Отчёт по ошибкам штабелеров</h3>
          </div>

          <div className="form-group">
            <label className="form-label">Данные из отчёта (вставьте сырые строки)</label>
            <textarea value={input} onChange={(e) => setInput(e.target.value)} className="form-input" rows="4" placeholder="SC0401	Conveyor	1006940137864695808	ScFault1	DH01	超边报警	2026.04.18 21:22:37	2026.04.18 21:47:44"></textarea>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '10px', marginBottom: '14px' }}>
            <div style={{ background: 'var(--background)', border: '1px solid var(--border)', borderRadius: '8px', padding: '12px', textAlign: 'center' }}>
              <div style={{ fontSize: '10px', color: 'var(--text-secondary)', marginBottom: '4px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Количество ошибок</div>
              <div style={{ fontSize: '22px', fontWeight: 700, color: 'var(--text-primary)' }}>{errorCount || '—'}</div>
              <div style={{ fontSize: '10px', color: 'var(--text-secondary)', marginTop: '2px' }}>авто из данных</div>
            </div>
            <div style={{ background: 'var(--background)', border: '1px solid var(--border)', borderRadius: '8px', padding: '12px', textAlign: 'center' }}>
              <div style={{ fontSize: '10px', color: 'var(--text-secondary)', marginBottom: '4px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Время в ошибке</div>
              <div style={{ fontSize: '22px', fontWeight: 700, color: 'var(--text-primary)' }}>{errorTime || '—'}</div>
              <div style={{ fontSize: '10px', color: 'var(--text-secondary)', marginTop: '2px' }}>мин · авто</div>
            </div>
            <div style={{ background: 'var(--background)', border: '1px solid var(--border)', borderRadius: '8px', padding: '12px', textAlign: 'center' }}>
              <div style={{ fontSize: '10px', color: 'var(--text-secondary)', marginBottom: '4px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>ВОТС (мин)</div>
              <div style={{ fontSize: '22px', fontWeight: 700, color: 'var(--text-primary)' }}>{votsTime || '—'}</div>
              <div style={{ fontSize: '10px', color: 'var(--text-secondary)', marginTop: '2px' }}>мин · авто</div>
            </div>
          </div>

          <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap', alignItems: 'center' }}>
            <button onClick={parseExcelData} className="btn btn-secondary">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="20 6 9 17 4 12"/></svg>
              Парсить данные
            </button>
            <button onClick={clearExcelData} className="btn btn-secondary">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
              Очистить
            </button>
          </div>

          <div style={{ marginTop: '12px' }}>
            <div className="form-group" style={{ margin: 0 }}>
              <label className="form-label">Название файла (ошибки)</label>
              <input type="text" value={fileName} onChange={(e) => setFileName(e.target.value)} className="form-input" placeholder="stackers_report" />
            </div>
          </div>
        </div>

        {/* Отчёт смены */}
        <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: '10px', padding: '16px', marginBottom: '20px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '14px' }}>
            <div style={{ width: '4px', height: '18px', background: '#22c55e', borderRadius: '2px' }}></div>
            <h3 style={{ fontSize: '14px', fontWeight: 700, margin: 0, color: 'var(--text-primary)' }}>Отчёт смены (Дашборд)</h3>
          </div>

          {/* Отгрузка */}
          <div style={{ marginBottom: '16px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '10px' }}>
              <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--text-primary)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Отгрузка</span>
              <button onClick={addShipmentRow} className="btn btn-secondary" style={{ padding: '3px 12px', fontSize: '12px' }}>+ Добавить</button>
              <button onClick={() => document.getElementById('shipmentExcelInput').click()} className="btn btn-secondary" style={{ padding: '3px 12px', fontSize: '12px' }}>📁 Загрузить Excel</button>
              <input type="file" id="shipmentExcelInput" accept=".xlsx,.xls" style={{ display: 'none' }} onChange={parseShipmentExcel} />
              <button onClick={parseShipmentData} className="btn btn-secondary" style={{ padding: '3px 12px', fontSize: '12px' }}>📊 Парсировать данные</button>
            </div>
            {shipRows.map((row, index) => (
              <div key={row.id} style={{ display: 'grid', gridTemplateColumns: '2fr 1fr 1fr 1fr auto', gap: '8px', marginBottom: '8px', alignItems: 'end' }}>
                <div className="form-group" style={{ margin: 0 }}>
                  <label className="form-label" style={{ fontSize: '11px' }}>Номер Кис</label>
                  <input type="text" value={row.kisNum} onChange={(e) => { const newRows = [...shipRows]; newRows[index].kisNum = e.target.value; setShipRows(newRows); }} className="form-input" placeholder="АЛЦ-0114435" style={{ fontSize: '12px', padding: '6px' }} />
                </div>
                <div className="form-group" style={{ margin: 0 }}>
                  <label className="form-label" style={{ fontSize: '11px' }}>Кол-во паллет</label>
                  <input type="text" value={row.palletQty} onChange={(e) => { const newRows = [...shipRows]; newRows[index].palletQty = e.target.value; setShipRows(newRows); }} className="form-input" placeholder="40" style={{ fontSize: '12px', padding: '6px' }} />
                </div>
                <div className="form-group" style={{ margin: 0 }}>
                  <label className="form-label" style={{ fontSize: '11px' }}>Время начала</label>
                  <input type="text" value={row.startTime} onChange={(e) => { const newRows = [...shipRows]; newRows[index].startTime = e.target.value; setShipRows(newRows); }} className="form-input" placeholder="09:41:05" style={{ fontSize: '12px', padding: '6px' }} />
                </div>
                <div className="form-group" style={{ margin: 0 }}>
                  <label className="form-label" style={{ fontSize: '11px' }}>Время окончания</label>
                  <input type="text" value={row.endTime} onChange={(e) => { const newRows = [...shipRows]; newRows[index].endTime = e.target.value; setShipRows(newRows); }} className="form-input" placeholder="10:16:30" style={{ fontSize: '12px', padding: '6px' }} />
                </div>
                <button onClick={() => setShipRows(shipRows.filter(r => r.id !== row.id))} style={{ padding: '8px 10px', background: 'var(--danger,#ef4444)', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', height: '36px', marginBottom: 0 }}>✕</button>
              </div>
            ))}
          </div>

          {/* Приемка */}
          <div style={{ marginBottom: '16px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '10px' }}>
              <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--text-primary)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Приемка</span>
              <button onClick={addReceiptRow} className="btn btn-secondary" style={{ padding: '3px 12px', fontSize: '12px' }}>+ Добавить</button>
              <button onClick={() => document.getElementById('receiptExcelInput').click()} className="btn btn-secondary" style={{ padding: '3px 12px', fontSize: '12px' }}>📁 Загрузить Excel</button>
              <input type="file" id="receiptExcelInput" accept=".xlsx,.xls" style={{ display: 'none' }} onChange={parseReceiptExcel} />
              <button onClick={parseReceiptData} className="btn btn-secondary" style={{ padding: '3px 12px', fontSize: '12px' }}>📊 Парсировать данные</button>
            </div>
            {recvRows.map((row, index) => (
              <div key={row.id} style={{ display: 'grid', gridTemplateColumns: '2fr 1fr 1fr 1fr auto', gap: '8px', marginBottom: '8px', alignItems: 'end' }}>
                <div className="form-group" style={{ margin: 0 }}>
                  <label className="form-label" style={{ fontSize: '11px' }}>Номер Кис</label>
                  <input type="text" value={row.kisNum} onChange={(e) => { const newRows = [...recvRows]; newRows[index].kisNum = e.target.value; setRecvRows(newRows); }} className="form-input" placeholder="АЛЦ-0114435" style={{ fontSize: '12px', padding: '6px' }} />
                </div>
                <div className="form-group" style={{ margin: 0 }}>
                  <label className="form-label" style={{ fontSize: '11px' }}>Кол-во паллет</label>
                  <input type="text" value={row.palletQty} onChange={(e) => { const newRows = [...recvRows]; newRows[index].palletQty = e.target.value; setRecvRows(newRows); }} className="form-input" placeholder="40" style={{ fontSize: '12px', padding: '6px' }} />
                </div>
                <div className="form-group" style={{ margin: 0 }}>
                  <label className="form-label" style={{ fontSize: '11px' }}>Время начала</label>
                  <input type="text" value={row.startTime} onChange={(e) => { const newRows = [...recvRows]; newRows[index].startTime = e.target.value; setRecvRows(newRows); }} className="form-input" placeholder="09:41:05" style={{ fontSize: '12px', padding: '6px' }} />
                </div>
                <div className="form-group" style={{ margin: 0 }}>
                  <label className="form-label" style={{ fontSize: '11px' }}>Время окончания</label>
                  <input type="text" value={row.endTime} onChange={(e) => { const newRows = [...recvRows]; newRows[index].endTime = e.target.value; setRecvRows(newRows); }} className="form-input" placeholder="10:16:30" style={{ fontSize: '12px', padding: '6px' }} />
                </div>
                <button onClick={() => setRecvRows(recvRows.filter(r => r.id !== row.id))} style={{ padding: '8px 10px', background: 'var(--danger,#ef4444)', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', height: '36px', marginBottom: 0 }}>✕</button>
              </div>
            ))}
          </div>

          {/* Паллеты */}
          <div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '10px' }}>
              <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--text-primary)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Паллеты</span>
              <button onClick={calcShiftPallets} className="btn btn-secondary" style={{ padding: '3px 12px', fontSize: '12px' }}>Рассчитать</button>
            </div>
            <div className="form-group">
              <label className="form-label">Список паллет (ID\tD1 или ID\tD2)</label>
              <textarea value={shiftPalletsInput} onChange={(e) => setShiftPalletsInput(e.target.value)} className="form-input" rows="4" placeholder="L01X01Y73Z01D1&#10;L01X01Y73Z01D2&#10;L02X02Y74Z01D1&#10;..."></textarea>
            </div>
            {shiftTotal > 0 && (
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3,1fr)', gap: '10px', marginTop: '10px' }}>
                <div style={{ background: 'var(--background)', border: '1px solid var(--border)', borderRadius: '8px', padding: '12px', textAlign: 'center' }}>
                  <div style={{ fontSize: '10px', color: 'var(--text-secondary)', marginBottom: '4px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Парные</div>
                  <div style={{ fontSize: '20px', fontWeight: 700, color: '#22c55e' }}>{shiftPaired}</div>
                </div>
                <div style={{ background: 'var(--background)', border: '1px solid var(--border)', borderRadius: '8px', padding: '12px', textAlign: 'center' }}>
                  <div style={{ fontSize: '10px', color: 'var(--text-secondary)', marginBottom: '4px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Непарные</div>
                  <div style={{ fontSize: '20px', fontWeight: 700, color: '#ef4444' }}>{shiftUnpaired}</div>
                </div>
                <div style={{ background: 'var(--background)', border: '1px solid var(--border)', borderRadius: '8px', padding: '12px', textAlign: 'center' }}>
                  <div style={{ fontSize: '10px', color: 'var(--text-secondary)', marginBottom: '4px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Всего</div>
                  <div style={{ fontSize: '20px', fontWeight: 700, color: 'var(--text-primary)' }}>{shiftTotal}</div>
                </div>
              </div>
            )}
          </div>

          <div style={{ marginTop: '12px' }}>
            <div className="form-group" style={{ margin: 0 }}>
              <label className="form-label">Название файла (дашборд)</label>
              <input type="text" value={shiftFileName} onChange={(e) => setShiftFileName(e.target.value)} className="form-input" placeholder="shift_report" />
            </div>
          </div>
        </div>

        <div style={{ paddingTop: '20px', borderTop: '1px solid var(--border)', marginTop: '4px' }}>
          <button onClick={generateBothReports} className="btn btn-primary" style={{ width: '100%', padding: '14px', fontSize: '15px', fontWeight: 600 }}>
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            Скачать оба файла
          </button>
        </div>
      </div>
    </div>
  )
}

export default ExcelGenerator
