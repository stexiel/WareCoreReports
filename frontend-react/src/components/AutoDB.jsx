import { useState, useEffect } from 'react'
import ExcelJS from 'exceljs'

function AutoDB() {
  const [shift, setShift] = useState('currentDay')
  const [shiftDate, setShiftDate] = useState('')
  const [dbFrom, setDbFrom] = useState('')
  const [dbTo, setDbTo] = useState('')
  const [inzhener, setInzhener] = useState('')
  const [tex, setTex] = useState('')
  const [loading, setLoading] = useState(false)
  const [status, setStatus] = useState('')

  // Автоопределение текущей смены при загрузке
  useEffect(() => {
    const now = new Date()
    const hour = now.getHours()
    // Дневная смена: 08:00 - 20:00
    // Ночная смена: 20:00 - 08:00
    const autoShift = (hour >= 8 && hour < 20) ? 'currentDay' : 'currentNight'
    setShift(autoShift)
    updateDates(autoShift)
  }, [])

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
      const to = new Date(date.getTime() + 12 * 60 * 60 * 1000)
      setDbFrom(formatDateTimeLocal(date))
      setDbTo(formatDateTimeLocal(to))
    }
  }

  const downloadExcel = async (workbook, fileName) => {
    const buffer = await workbook.xlsx.writeBuffer()
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    const url = window.URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = fileName
    a.click()
    window.URL.revokeObjectURL(url)
  }

  const getApiBase = () => {
    return 'http://localhost:3000'
  }

  const fetchJSON = async (url) => {
    const token = localStorage.getItem('authToken')
    const headers = {
      'Content-Type': 'application/json'
    }
    if (token) {
      headers['Authorization'] = `Bearer ${token}`
    }
    const response = await fetch(url, { headers })
    if (!response.ok) throw new Error(`HTTP ${response.status}`)
    return await response.json()
  }

  const extractHMS = (datetimeStr) => {
    if (!datetimeStr) return ''
    const match = datetimeStr.match(/\d{2}:\d{2}:\d{2}/)
    return match ? match[0] : ''
  }

  const loadFromDatabase = async () => {
    const from = dbFrom
    const to = dbTo

    if (!from || !to) {
      setStatus('❌ Укажите начало и окончание смены')
      return
    }

    const base = getApiBase()
    setLoading(true)
    setStatus('⏳ Загрузка из БД...')

    console.log('Запрос к БД с параметрами:', from, '-', to)

    try {
      const qs = `from=${encodeURIComponent(from)}&to=${encodeURIComponent(to)}`
      const [faultRes, shipRes, recvRes, palletRes] = await Promise.all([
        fetchJSON(`${base}/api/fault-report?${qs}`),
        fetchJSON(`${base}/api/shipment?${qs}`),
        fetchJSON(`${base}/api/receipt?${qs}`),
        fetchJSON(`${base}/api/pallets?${qs}`)
      ])

      console.log('Получено из БД ошибок:', faultRes.rows.length)
      console.log('Пример данных из БД:', faultRes.rows[0])

      // Преобразуем даты в объекты Date для фильтрации
      // from и to в локальном формате (без timezone), интерпретируем как локальное время
      const fromDate = new Date(from)
      const toDate = new Date(to)

      // Переводим from/to в UTC для сравнения с данными из БД (которые в UTC)
      const fromDateUTC = new Date(fromDate.getTime() - fromDate.getTimezoneOffset() * 60000)
      const toDateUTC = new Date(toDate.getTime() - toDate.getTimezoneOffset() * 60000)

      console.log('Фильтр по времени (local):', fromDate.toISOString(), '-', toDate.toISOString())
      console.log('Фильтр по времени (UTC):', fromDateUTC.toISOString(), '-', toDateUTC.toISOString())

      // ---- 1. Ошибки штабелеров (eqpt_fault_record) ----
      const parsedRows = faultRes.rows
        .filter(r => {
          if (!r.stacker) return false
          // Данные из БД в UTC формате (с Z), сравниваем с UTC диапазоном
          const beginTime = new Date(r.beginTime)
          const inRange = beginTime >= fromDateUTC && beginTime <= toDateUTC
          console.log('Проверка записи:', r.beginTime, 'inRange:', inRange)
          if (!inRange) {
            console.log('Отфильтровано:', r.beginTime)
          }
          return inRange
        })
        .map(r => ({
          time: extractHMS(r.beginTime),
          endTime: extractHMS(r.finishTime),
          startISO: (r.beginTime || '').substring(0, 19),
          rawStart: extractHMS(r.beginTime),
          rawEnd: extractHMS(r.finishTime),
          stacker: r.stacker
        }))

      // Сортировка с учётом ночной смены 20:00–08:00
      function shiftSortKey(isoStr) {
        const d = new Date(isoStr)
        const h = d.getHours()
        if (h < 8) return d.getTime() + 24 * 3600 * 1000
        return d.getTime()
      }
      parsedRows.sort((a, b) => shiftSortKey(a.startISO) - shiftSortKey(b.startISO))

      // Строим excelData с statuses как в старой версии
      const excelData = parsedRows.map(row => ({
        time: row.time,
        endTime: row.endTime,
        startISO: row.startISO,
        stacker: row.stacker,
        statuses: Array(10).fill('').map((_, i) => i + 1 === row.stacker ? 'F06' : '')
      }))

      // ---- 2. Отгрузка (out_delivering_bill) ----
      const newShipRows = []
      shipRes.rows.forEach(r => {
        // Фильтрация по времени на стороне клиента (UTC)
        const startTime = new Date(r.start_time)
        if (startTime >= fromDateUTC && startTime <= toDateUTC) {
          newShipRows.push({
            id: Date.now() + Math.random(),
            kisNum: r.bill_no || '',
            palletQty: r.qty || 0,
            startTime: extractHMS(r.start_time),
            endTime: extractHMS(r.end_time),
            totalTime: ''
          })
        }
      })

      // ---- 3. Приемка (in_receiving_bill) ----
      const newRecvRows = []
      recvRes.rows.forEach(r => {
        // Фильтрация по времени на стороне клиента (UTC)
        const startTime = new Date(r.start_time)
        if (startTime >= fromDateUTC && startTime <= toDateUTC) {
          newRecvRows.push({
            id: Date.now() + Math.random(),
            kisNum: r.bill_no || '',
            palletQty: r.qty || 0,
            startTime: extractHMS(r.start_time),
            endTime: extractHMS(r.end_time),
            totalTime: ''
          })
        }
      })

      // ---- 4. Паллеты D1/D2 (inv_container) ----
      const palletLines = []
      palletRes.paired.forEach(p => { palletLines.push(p.id + 'D1'); palletLines.push(p.id + 'D2'); })
      palletRes.unpaired.forEach(p => { palletLines.push(p.id + (p.missing === 'D2' ? 'D1' : 'D2')); })
      const shiftPalletsInput = palletLines.join('\n')

      // Сохраняем данные в localStorage для использования в ExcelGenerator
      localStorage.setItem('dbExcelData', JSON.stringify(excelData))
      localStorage.setItem('dbShipmentRows', JSON.stringify(newShipRows))
      localStorage.setItem('dbReceiptRows', JSON.stringify(newRecvRows))
      localStorage.setItem('dbShiftPalletsInput', shiftPalletsInput)
      localStorage.setItem('dbInzhener', inzhener || 'Aser')
      localStorage.setItem('dbTex', tex || 'Лёша, Владимир')

      setStatus(`✅ Загружено: ошибок ${parsedRows.length}, отгрузка ${shipRes.rows.length}, приемка ${recvRes.rows.length}, паллет ${palletRes.totalRows}`)
    } catch (err) {
      console.error('loadFromDatabase error:', err)
      setStatus('❌ Ошибка: ' + err.message)
    } finally {
      setLoading(false)
    }
  }

  const downloadBothFromDb = async () => {
    await loadFromDatabase()
    if (status.startsWith('❌')) return

    // Копируем данные из полей автозагрузки в поля генератора
    localStorage.setItem('excelInzhener', inzhener || 'Aser')
    localStorage.setItem('excelTex', tex || 'Лёша, Владимир')

    // Генерируем и скачиваем оба Excel файла
    setStatus('⏳ Генерация Excel файлов...')
    try {
      // Импортируем функции из ExcelGenerator
      const { generateExcel, generateShiftReport } = await import('./ExcelGenerator.jsx')

      // Загружаем данные из localStorage
      const dbExcelData = JSON.parse(localStorage.getItem('dbExcelData') || '[]')
      const dbShipmentRows = JSON.parse(localStorage.getItem('dbShipmentRows') || '[]')
      const dbReceiptRows = JSON.parse(localStorage.getItem('dbReceiptRows') || '[]')
      const dbShiftPalletsInput = localStorage.getItem('dbShiftPalletsInput') || ''

      // Вызываем генерацию файлов напрямую
      await generateExcelWithDBData(dbExcelData, inzhener, tex)
      await new Promise(r => setTimeout(r, 600))
      await generateShiftReportWithDBData(dbShipmentRows, dbReceiptRows, dbShiftPalletsInput, dbExcelData, inzhener, tex)

      setStatus('✅ Файлы скачаны!')
      setTimeout(() => setStatus(''), 3000)
    } catch (err) {
      console.error('Error generating Excel:', err)
      setStatus('❌ Ошибка генерации: ' + err.message)
    }
  }

  // Вспомогательные функции для генерации Excel с данными из БД
  const generateExcelWithDBData = async (data, inzhenerVal, texVal) => {
    const ExcelJS = await import('exceljs')
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

    // Автоматический расчет времени ошибок и ВОТС
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
    const ExcelJS = await import('exceljs')
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

    // Автоматический расчет ошибок
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
    const EMPTY = { font:{name:FONT,size:FSIZE}, alignment:{horizontal:'center',vertical:'middle'} }
    errRows.forEach(([label, val]) => {
      const r = ws.addRow([label, val])
      r.getCell(1).style = cs(BEIGE_ROW, true, 'left')
      r.getCell(2).style = cs(BEIGE_ROW, true, 'center')
      r.height = 16
    })

    ws.addRow([])

    // Паллеты
    const lines = palletsInput.trim().split('\n').filter(line => line.trim())
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

  useEffect(() => {
    updateDates('currentDay')
  }, [])

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
              <input
                id="dbInzhener"
                type="text"
                value={inzhener}
                onChange={(e) => setInzhener(e.target.value)}
                className="form-input"
                placeholder="Aser"
              />
            </div>
            <div className="form-group" style={{ margin: 0 }}>
              <label className="form-label">Тех служба</label>
              <input
                id="dbTex"
                type="text"
                value={tex}
                onChange={(e) => setTex(e.target.value)}
                className="form-input"
                placeholder="Лёша, Владимир"
              />
            </div>
          </div>

          <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap', alignItems: 'center' }}>
            <button
              id="dbDownloadBtn"
              onClick={downloadBothFromDb}
              disabled={loading}
              className="btn btn-primary"
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                <polyline points="7 10 12 15 17 10"/>
                <line x1="12" y1="15" x2="12" y2="3"/>
              </svg>
              Загрузить и скачать 2 файла
            </button>
            <span id="dbLoadStatus" style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>{status}</span>
          </div>
        </div>
      </div>
    </div>
  )
}

export default AutoDB
