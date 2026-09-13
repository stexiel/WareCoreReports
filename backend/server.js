require('dotenv').config();
const express = require('express');
const cors = require('cors');
const path = require('path');
const rateLimit = require('express-rate-limit');
const fs = require('fs');
const https = require('https');
const { queryWMS, queryWCS } = require('./db');

const app = express();
const PORT = parseInt(process.env.PORT || '3000', 10);

// Драйвер mssql хранит datetime-поля БД как UTC-компоненты Date объекта
// (то есть getUTCHours() и т.д. соответствуют реальному значению в БД).
// Поэтому для получения реального времени используем UTC-геттеры,
// а не локальные (которые применяют смещение часового пояса машины).
const pad2 = (n) => String(n).padStart(2, '0');

// Строка без суффикса 'Z' - фронтенд распарсит её как локальное время браузера,
// что совпадает с реальным значением из БД (без двойного смещения зоны)
const dbDateToLocalString = (date) => {
    if (!date) return '';
    const d = new Date(date);
    return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}T${pad2(d.getUTCHours())}:${pad2(d.getUTCMinutes())}:${pad2(d.getUTCSeconds())}`;
}

// Случайная длительность в минутах в заданном диапазоне (для оценки реального времени оператора,
// т.к. фактическая длительность в БД отражает работу оборудования, а не время оператора)
const randomDuration = (min, max) => {
    return Math.floor(Math.random() * (max - min + 1)) + min;
};

// Читаемая строка для описаний
const dbDateToReadable = (date) => {
    if (!date) return '';
    const d = new Date(date);
    return `${pad2(d.getUTCDate())}.${pad2(d.getUTCMonth() + 1)}.${d.getUTCFullYear()} ${pad2(d.getUTCHours())}:${pad2(d.getUTCMinutes())}:${pad2(d.getUTCSeconds())}`;
};

const getShiftFromDbDate = (date) => {
    const hours = new Date(date).getUTCHours();
    return hours >= 8 && hours < 20 ? 'день' : 'ночь';
};

// Парсит наивную локальную строку "YYYY-MM-DDTHH:mm[:ss]" (без 'Z') в Date,
// у которого UTC-компоненты совпадают с указанными в строке цифрами.
// Это нужно, чтобы при отправке параметра в SQL Server (который хранит
// datetime без часового пояса) сравнение происходило с реальным значением
// из БД независимо от системного часового пояса самого Node.js процесса.
function parseNaiveLocalDate(str) {
    const m = String(str).match(/^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})(?::(\d{2}))?/);
    if (!m) return new Date(NaN);
    const [, y, mo, d, h, mi, s] = m;
    return new Date(Date.UTC(+y, +mo - 1, +d, +h, +mi, s ? +s : 0));
}

app.use(cors({
    origin: ['http://localhost:5173', 'http://localhost:5174', 'http://localhost:3000', 'http://172.30.1.249:5173', 'http://172.30.1.249:3000', 'http://10.250.96.88:5173', 'http://10.250.96.88:3000'],
    credentials: true
}));
app.use(express.json());

// Rate limiting middleware
const limiter = rateLimit({
    windowMs: 15 * 60 * 1000, // 15 минут
    max: 100, // максимум 100 запросов за окно
    message: { error: 'Слишком много запросов, попробуйте позже' }
});

app.use('/api/', limiter);

// Проект без авторизации: доступ открыт всем в локальной сети.
// authenticateToken/requireAdmin/validateCSRF оставлены как no-op,
// чтобы не переписывать сигнатуры существующих роутов.
function authenticateToken(req, res, next) {
    next();
}

function requireAdmin(req, res, next) {
    next();
}

function validateCSRF(req, res, next) {
    next();
}

// Карта оборудования штабелеров -> номер (1..10), как в index.html
const EQUIPMENT_TO_STACKER = {
    'SC0101': 1, 'SC0201': 2, 'SC0301': 3, 'SC0401': 4, 'SC0501': 5,
    'SC0601': 6, 'SC0701': 7, 'SC0801': 8, 'SC0901': 9, 'SC1001': 10
};

function parseRange(req, res) {
    const { from, to } = req.query;
    if (!from || !to) {
        res.status(400).json({ error: 'Требуются параметры from и to (ISO дата-время)' });
        return null;
    }
    const fromDate = parseNaiveLocalDate(from);
    const toDate = parseNaiveLocalDate(to);
    if (isNaN(fromDate) || isNaN(toDate)) {
        res.status(400).json({ error: 'Некорректный формат даты' });
        return null;
    }
    return { fromDate, toDate };
}

app.get('/api/health', (req, res) => {
    res.json({ ok: true, time: new Date().toISOString() });
});

// ===== Отчёт по ошибкам штабелеров (WCS_1386.eqpt_fault_record) =====
app.get('/api/fault-report', authenticateToken, async (req, res) => {
    const range = parseRange(req, res);
    if (!range) return;
    try {
        const codes = Object.keys(EQUIPMENT_TO_STACKER);
        const inClause = codes.map((c, i) => `@code${i}`).join(',');
        const params = { from: range.fromDate, to: range.toDate };
        codes.forEach((c, i) => { params['code' + i] = c; });

        const result = await queryWCS(
            `SELECT eqpt_code, begin_time, finish_time
             FROM eqpt_fault_record
             WHERE eqpt_code IN (${inClause})
               AND begin_time >= @from AND begin_time <= @to
             ORDER BY begin_time ASC`,
            params
        );

        const rows = result.recordset.map(r => ({
            stacker: EQUIPMENT_TO_STACKER[r.eqpt_code],
            eqptCode: r.eqpt_code,
            beginTime: dbDateToLocalString(r.begin_time),
            finishTime: dbDateToLocalString(r.finish_time)
        }));

        res.json({ rows });
    } catch (err) {
        console.error('Ошибка /api/fault-report:', err);
        res.status(500).json({ error: err.message });
    }
});

// ===== Отгрузка (WMS_1386.out_delivering_bill + wcs_job_his) =====
app.get('/api/shipment', authenticateToken, async (req, res) => {
    const range = parseRange(req, res);
    if (!range) return;
    try {
        const result = await queryWMS(
            `SELECT b.bill_no, b.start_time, b.end_time,
                    COUNT(DISTINCT j.container_code) AS qty
             FROM out_delivering_bill b
             LEFT JOIN wcs_job_his j
                    ON j.bill_no = b.bill_no
                   AND j.bill_class = 'Delivering'
                   AND j.type = 'MoveOut'
             WHERE b.is_deleted = 0
               AND b.start_time >= @from AND b.start_time <= @to
             GROUP BY b.bill_no, b.start_time, b.end_time
             ORDER BY b.start_time ASC`,
            { from: range.fromDate, to: range.toDate }
        );
        const rows = result.recordset.map(r => ({
            bill_no: r.bill_no,
            start_time: dbDateToLocalString(r.start_time),
            end_time: dbDateToLocalString(r.end_time),
            qty: r.qty
        }));
        res.json({ rows });
    } catch (err) {
        console.error('Ошибка /api/shipment:', err);
        res.status(500).json({ error: err.message });
    }
});

// ===== Приемка (WMS_1386.in_receiving_bill + in_receiving_bill_detail + in_receiving_record) =====
app.get('/api/receipt', authenticateToken, async (req, res) => {
    const range = parseRange(req, res);
    if (!range) return;
    try {
        const result = await queryWMS(
            `SELECT b.bill_no, b.start_time, b.end_time,
                    COUNT(DISTINCT r.container_code) AS qty
             FROM in_receiving_bill b
             LEFT JOIN in_receiving_bill_detail d ON d.receiving_bill_id = b.id
             LEFT JOIN in_receiving_record r ON r.receiving_bill_detail_id = d.id AND r.status = 'Completed'
             WHERE b.is_deleted = 0
               AND b.start_time >= @from AND b.start_time <= @to
             GROUP BY b.bill_no, b.start_time, b.end_time
             ORDER BY b.start_time ASC`,
            { from: range.fromDate, to: range.toDate }
        );
        const rows = result.recordset.map(r => ({
            bill_no: r.bill_no,
            start_time: dbDateToLocalString(r.start_time),
            end_time: dbDateToLocalString(r.end_time),
            qty: r.qty
        }));
        res.json({ rows });
    } catch (err) {
        console.error('Ошибка /api/receipt:', err);
        res.status(500).json({ error: err.message });
    }
});

// ===== Паллеты D1/D2 (WMS_1386.inv_container) =====
app.get('/api/pallets', authenticateToken, async (req, res) => {
    const range = parseRange(req, res);
    if (!range) return;
    try {
        // Получаем все паллеты с их временами обновления
        const result = await queryWMS(
            `SELECT loc_code, update_time
             FROM inv_container
             WHERE (loc_code LIKE '%D1' OR loc_code LIKE '%D2')`,
            {}
        );

        const pallets = {};
        for (const row of result.recordset) {
            const loc = row.loc_code;
            const type = loc.slice(-2);
            const id = loc.slice(0, -2);
            const updateTime = new Date(row.update_time);

            if (!pallets[id]) {
                pallets[id] = { d1: false, d2: false, d1Time: null, d2Time: null };
            }

            // Проверяем, находится ли время обновления в заданном диапазоне
            const inRange = updateTime >= new Date(range.fromDate) && updateTime <= new Date(range.toDate);

            if (type === 'D1') {
                if (inRange) pallets[id].d1 = true;
                pallets[id].d1Time = updateTime;
            }
            if (type === 'D2') {
                if (inRange) pallets[id].d2 = true;
                pallets[id].d2Time = updateTime;
            }
        }

        // Учитываем только те паллеты, у которых хотя бы одна сторона попала в диапазон
        const filteredPallets = {};
        for (const id in pallets) {
            const p = pallets[id];
            if (p.d1 || p.d2) {
                filteredPallets[id] = p;
            }
        }

        const paired = [];
        const unpaired = [];
        for (const id in filteredPallets) {
            const p = filteredPallets[id];
            if (p.d1 && p.d2) paired.push({ id });
            else if (p.d1) unpaired.push({ id, missing: 'D2' });
            else if (p.d2) unpaired.push({ id, missing: 'D1' });
        }

        res.json({
            paired,
            unpaired,
            pairedRows: paired.length * 2,
            unpairedRows: unpaired.length,
            totalRows: paired.length * 2 + unpaired.length
        });
    } catch (err) {
        console.error('Ошибка /api/pallets:', err);
        res.status(500).json({ error: err.message });
    }
});

// ===== Логи операторов (operator_logs) =====
// Локальное хранилище в JSON файле
const OPERATOR_LOGS_FILE = path.join(__dirname, 'operator_logs.json');

// Инициализация файла логов
const initOperatorLogsFile = () => {
    if (!fs.existsSync(OPERATOR_LOGS_FILE)) {
        fs.writeFileSync(OPERATOR_LOGS_FILE, JSON.stringify([], null, 2));
        console.log('[Operator Logs] Файл создан');
    } else {
        console.log('[Operator Logs] Файл уже существует');
    }
};

// Чтение логов
const readOperatorLogs = () => {
    try {
        const data = fs.readFileSync(OPERATOR_LOGS_FILE, 'utf8');
        return JSON.parse(data);
    } catch (err) {
        console.error('[Operator Logs] Ошибка чтения:', err);
        return [];
    }
};

// Запись логов
const writeOperatorLogs = (logs) => {
    try {
        fs.writeFileSync(OPERATOR_LOGS_FILE, JSON.stringify(logs, null, 2));
    } catch (err) {
        console.error('[Operator Logs] Ошибка записи:', err);
    }
};

// Инициализация при запуске
initOperatorLogsFile();

// Загрузка исторических данных из БД в логи
const loadHistoricalData = async () => {
    try {
        const logs = readOperatorLogs();
        
        // Загружаем ошибки штабелеров за последние 30 дней
        const thirtyDaysAgo = new Date();
        thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
        
        console.log('[Operator Logs] Загрузка ошибок штабелеров...');
        const faultResult = await queryWCS(
            `SELECT eqpt_code, begin_time, finish_time
             FROM eqpt_fault_record
             WHERE begin_time >= @from
             ORDER BY begin_time DESC`,
            { from: thirtyDaysAgo }
        );
        console.log(`[Operator Logs] Получено ${faultResult.recordset.length} ошибок из БД`);
        
        // Определяем тип оборудования и его название в родительном падеже (для "Ошибка ...")
        const getEquipmentGenitive = (code) => {
            if (code === 'A02PLC') return 'шатла';
            if (code.startsWith('SC') && code.length === 6) {
                const num = parseInt(code.substring(2, 4));
                if (num >= 1 && num <= 10) {
                    return `штабелера ${num}`;
                }
            }
            if (code.startsWith('A0')) return 'конвейера';
            return code;
        };
        
        // Группируем ошибки по типу оборудования и времени завершения (в течение 1 минуты)
        const groupedErrors = {};
        faultResult.recordset.forEach(r => {
            const genitive = getEquipmentGenitive(r.eqpt_code);
            const finishTime = new Date(r.finish_time);
            const beginTime = new Date(r.begin_time);
            // Округляем до минуты для группировки
            const minuteKey = Math.floor(finishTime.getTime() / 60000);
            const key = `${minuteKey}_${genitive}`;
            
            if (!groupedErrors[key]) {
                groupedErrors[key] = {
                    time: finishTime,
                    beginTime: beginTime,
                    genitive: genitive,
                    codes: new Set()
                };
            }
            if (beginTime < groupedErrors[key].beginTime) {
                groupedErrors[key].beginTime = beginTime;
            }
            groupedErrors[key].codes.add(r.eqpt_code);
        });

        // Добавляем сгруппированные ошибки в логи
        Object.values(groupedErrors).forEach(group => {
            const codesList = Array.from(group.codes).sort().join(', ');
            const durationMinutes = group.genitive === 'конвейера' ? randomDuration(1, 3) : randomDuration(3, 5);
            const log = {
                id: logs.length > 0 ? Math.max(...logs.map(l => l.id)) + 1 : 1,
                operator_name: 'System',
                action: `Ошибка ${group.genitive}: ${codesList}`,
                description: `Ошибка сброшена`,
                timestamp: dbDateToLocalString(group.time),
                shift: getShiftFromDbDate(group.time),
                total_minutes: durationMinutes,
                duration_version: 2,
                created_at: new Date().toISOString()
            };
            logs.push(log);
        });
        
        // Загружаем отгрузки за последние 30 дней
        console.log('[Operator Logs] Загрузка отгрузок...');
        const shipmentResult = await queryWMS(
            `SELECT bill_no, start_time, end_time
             FROM out_delivering_bill
             WHERE is_deleted = 0 AND start_time >= @from
             ORDER BY start_time DESC`,
            { from: thirtyDaysAgo }
        );
        console.log(`[Operator Logs] Получено ${shipmentResult.recordset.length} отгрузок из БД`);
        
        // Добавляем отгрузки в логи
        shipmentResult.recordset.forEach(r => {
            const durationMinutes = randomDuration(3, 5);
            const log = {
                id: logs.length > 0 ? Math.max(...logs.map(l => l.id)) + 1 : 1,
                operator_name: 'System',
                action: `Отгрузка заказа ${r.bill_no}`,
                description: `Проверка накладной прошла, партия соответствует системе, распределение паллетов (вручную/автоматически), запуск отгрузки в ${dbDateToReadable(r.start_time)}, отгрузка завершена в ${dbDateToReadable(r.end_time)}`,
                timestamp: dbDateToLocalString(r.start_time),
                shift: getShiftFromDbDate(r.start_time),
                total_minutes: durationMinutes,
                duration_version: 2,
                created_at: new Date().toISOString()
            };
            logs.push(log);
        });
        
        // Загружаем отмены паллет при приемке за последние 30 дней
        // (приемка логируется только когда оператор отменяет паллету)
        console.log('[Operator Logs] Загрузка отмен паллет при приемке...');
        const receiptResult = await queryWMS(
            `SELECT r.container_code, r.update_time, b.bill_no
             FROM in_receiving_record r
             JOIN in_receiving_bill_detail d ON d.id = r.receiving_bill_detail_id
             JOIN in_receiving_bill b ON b.id = d.receiving_bill_id
             WHERE r.status = 'Canceled' AND r.update_time >= @from
             ORDER BY r.update_time DESC`,
            { from: thirtyDaysAgo }
        );
        console.log(`[Operator Logs] Получено ${receiptResult.recordset.length} отмен паллет из БД`);
        
        // Добавляем отмены паллет в логи
        receiptResult.recordset.forEach(r => {
            const log = {
                id: logs.length > 0 ? Math.max(...logs.map(l => l.id)) + 1 : 1,
                operator_name: 'System',
                action: `Приемка ${r.bill_no}`,
                description: `Оператор отменил паллету ${r.container_code}`,
                total_minutes: randomDuration(3, 5),
                duration_version: 2,
                timestamp: dbDateToLocalString(r.update_time),
                shift: getShiftFromDbDate(r.update_time),
                created_at: new Date().toISOString()
            };
            logs.push(log);
        });
        
        writeOperatorLogs(logs);
        console.log(`[Operator Logs] Загружено ${faultResult.recordset.length} ошибок, ${shipmentResult.recordset.length} отгрузок, ${receiptResult.recordset.length} отмен паллет`);
    } catch (err) {
        console.error('[Operator Logs] Ошибка загрузки исторических данных:', err);
    }
};

// Загрузка только новых данных из БД (без дублирования уже существующих записей)
const loadNewData = async () => {
    try {
        const logs = readOperatorLogs();

        // Определяем самую позднюю дату среди уже загруженных System-записей,
        // чтобы подгружать только новые данные из БД
        const systemLogs = logs.filter(l => l.operator_name === 'System');
        let sinceDate = new Date();
        sinceDate.setDate(sinceDate.getDate() - 30);
        if (systemLogs.length > 0) {
            const latestTimestamp = systemLogs.reduce((max, l) => {
                const t = new Date(l.timestamp);
                return t > max ? t : max;
            }, sinceDate);
            sinceDate = latestTimestamp;
        }

        const existingKeys = new Set(
            logs.map(l => `${l.action}|${l.timestamp}`)
        );

        let addedCount = 0;

        // Ошибки штабелеров
        const faultResult = await queryWCS(
            `SELECT eqpt_code, begin_time, finish_time
             FROM eqpt_fault_record
             WHERE begin_time >= @from
             ORDER BY begin_time DESC`,
            { from: sinceDate }
        );

        const getEquipmentGenitive = (code) => {
            if (code === 'A02PLC') return 'шатла';
            if (code.startsWith('SC') && code.length === 6) {
                const num = parseInt(code.substring(2, 4));
                if (num >= 1 && num <= 10) {
                    return `штабелера ${num}`;
                }
            }
            if (code.startsWith('A0')) return 'конвейера';
            return code;
        };

        const groupedErrors = {};
        faultResult.recordset.forEach(r => {
            const genitive = getEquipmentGenitive(r.eqpt_code);
            const finishTime = new Date(r.finish_time);
            const beginTime = new Date(r.begin_time);
            const minuteKey = Math.floor(finishTime.getTime() / 60000);
            const key = `${minuteKey}_${genitive}`;

            if (!groupedErrors[key]) {
                groupedErrors[key] = {
                    time: finishTime,
                    beginTime: beginTime,
                    genitive: genitive,
                    codes: new Set()
                };
            }
            if (beginTime < groupedErrors[key].beginTime) {
                groupedErrors[key].beginTime = beginTime;
            }
            groupedErrors[key].codes.add(r.eqpt_code);
        });

        Object.values(groupedErrors).forEach(group => {
            const codesList = Array.from(group.codes).sort().join(', ');
            const action = `Ошибка ${group.genitive}: ${codesList}`;
            const timestamp = dbDateToLocalString(group.time);
            const key = `${action}|${timestamp}`;
            if (existingKeys.has(key)) return;
            existingKeys.add(key);

            const durationMinutes = group.genitive === 'конвейера' ? randomDuration(1, 3) : randomDuration(3, 5);
            const log = {
                id: logs.length > 0 ? Math.max(...logs.map(l => l.id)) + 1 : 1,
                operator_name: 'System',
                action,
                description: `Ошибка сброшена`,
                timestamp,
                shift: getShiftFromDbDate(group.time),
                total_minutes: durationMinutes,
                duration_version: 2,
                created_at: new Date().toISOString()
            };
            logs.push(log);
            addedCount++;
        });

        // Отгрузки
        const shipmentResult = await queryWMS(
            `SELECT bill_no, start_time, end_time
             FROM out_delivering_bill
             WHERE is_deleted = 0 AND start_time >= @from
             ORDER BY start_time DESC`,
            { from: sinceDate }
        );

        shipmentResult.recordset.forEach(r => {
            const action = `Отгрузка заказа ${r.bill_no}`;
            const timestamp = dbDateToLocalString(r.start_time);
            const key = `${action}|${timestamp}`;
            if (existingKeys.has(key)) return;
            existingKeys.add(key);

            const durationMinutes = randomDuration(3, 5);
            const log = {
                id: logs.length > 0 ? Math.max(...logs.map(l => l.id)) + 1 : 1,
                operator_name: 'System',
                action,
                description: `Проверка накладной прошла, партия соответствует системе, распределение паллетов (вручную/автоматически), запуск отгрузки в ${dbDateToReadable(r.start_time)}, отгрузка завершена в ${dbDateToReadable(r.end_time)}`,
                timestamp,
                shift: getShiftFromDbDate(r.start_time),
                total_minutes: durationMinutes,
                duration_version: 2,
                created_at: new Date().toISOString()
            };
            logs.push(log);
            addedCount++;
        });

        // Отмены паллет при приемке
        const receiptResult = await queryWMS(
            `SELECT r.container_code, r.update_time, b.bill_no
             FROM in_receiving_record r
             JOIN in_receiving_bill_detail d ON d.id = r.receiving_bill_detail_id
             JOIN in_receiving_bill b ON b.id = d.receiving_bill_id
             WHERE r.status = 'Canceled' AND r.update_time >= @from
             ORDER BY r.update_time DESC`,
            { from: sinceDate }
        );

        receiptResult.recordset.forEach(r => {
            const action = `Приемка ${r.bill_no}`;
            const timestamp = dbDateToLocalString(r.update_time);
            const key = `${action}|${timestamp}`;
            if (existingKeys.has(key)) return;
            existingKeys.add(key);

            const log = {
                id: logs.length > 0 ? Math.max(...logs.map(l => l.id)) + 1 : 1,
                operator_name: 'System',
                action,
                description: `Оператор отменил паллету ${r.container_code}`,
                total_minutes: randomDuration(3, 5),
                duration_version: 2,
                timestamp,
                shift: getShiftFromDbDate(r.update_time),
                created_at: new Date().toISOString()
            };
            logs.push(log);
            addedCount++;
        });

        if (addedCount > 0) {
            writeOperatorLogs(logs);
            console.log(`[Operator Logs] Добавлено ${addedCount} новых записей`);
        }
    } catch (err) {
        console.error('[Operator Logs] Ошибка загрузки новых данных:', err);
    }
};

// Рутинные задачи начала и конца дневной смены (8:00-8:30 и 19:30-20:00).
// За каждую смену все рутинные задачи выполняет ОДИН оператор (не System и не вперемешку).
// ОТКЛЮЧЕНО: пользователи удалены
// const OPERATOR_NAMES = USERS.filter(u => u.role === 'operator').map(u => u.username);

// const ROUTINE_START_ACTIONS = [
//     { action: 'Перезапуск компьютера', description: 'Перезапуск компьютера в начале смены', min: 5, max: 7 },
//     { action: 'Запуск HMI', description: 'Запуск HMI в начале смены', min: 3, max: 6 },
//     { action: 'Вход в аккаунты WCS, WMS, WareCore Reports', description: 'Авторизация в системах WCS, WMS, WareCore Reports', min: 2, max: 3 },
//     { action: 'Проверка коробочного остатка', description: 'Спросить есть ли коробочный остаток в группе', min: 1, max: 1 },
// ];

// const ROUTINE_END_ACTIONS = [
//     { action: 'Подготовка отчета', description: 'Подготовка отчета в конце смены', min: 20, max: 25 },
//     { action: 'Отправление отчета', description: 'Отправление отчета в конце смены', min: 3, max: 4 },
// ];

// const generateRoutineLogs = () => {
//     if (OPERATOR_NAMES.length === 0) return;
//     const logs = readOperatorLogs();
//     const existingRoutineKeys = new Set(
//         logs.filter(l => l.routine).map(l => `${l.action}|${l.timestamp.slice(0, 10)}|${l.shift}`)
//     );
//     let addedCount = 0;
//     const today = new Date();

//     for (let i = 0; i < 29; i++) {
//         const day = new Date(today);
//         day.setDate(day.getDate() - i);
//         const dateStr = `${day.getFullYear()}-${pad2(day.getMonth() + 1)}-${pad2(day.getDate())}`;

//         // Один случайный оператор на всю дневную смену
//         const dayShiftOperator = OPERATOR_NAMES[Math.floor(Math.random() * OPERATOR_NAMES.length)];

//         // Начало дневной смены: 8:00, задачи последовательно друг за другом
//         let cursor = new Date(day);
//         cursor.setHours(8, 0, 0, 0);
//         ROUTINE_START_ACTIONS.forEach(item => {
//             const key = `${item.action}|${dateStr}|день`;
//             const duration = randomDuration(item.min, item.max);
//             if (!existingRoutineKeys.has(key)) {
//                 const log = {
//                     id: logs.length > 0 ? Math.max(...logs.map(l => l.id)) + 1 : 1,
//                     operator_name: dayShiftOperator,
//                     action: item.action,
//                     description: item.description,
//                     timestamp: `${dateStr}T${pad2(cursor.getHours())}:${pad2(cursor.getMinutes())}:00`,
//                     shift: 'день',
//                     total_minutes: duration,
//                     routine: true,
//                     created_at: new Date().toISOString()
//                 };
//                 logs.push(log);
//                 existingRoutineKeys.add(key);
//                 addedCount++;
//             }
//             cursor = new Date(cursor.getTime() + duration * 60000);
//         });

//         // Конец дневной смены: задачи заканчиваются ровно к 20:00
//         const endDurations = ROUTINE_END_ACTIONS.map(item => randomDuration(item.min, item.max));
//         const totalEndDuration = endDurations.reduce((a, b) => a + b, 0);
//         let endCursor = new Date(day);
//         endCursor.setHours(20, 0, 0, 0);
//         endCursor = new Date(endCursor.getTime() - totalEndDuration * 60000);
//         ROUTINE_END_ACTIONS.forEach((item, idx) => {
//             const key = `${item.action}|${dateStr}|день`;
//             const duration = endDurations[idx];
//             if (!existingRoutineKeys.has(key)) {
//                 const log = {
//                     id: logs.length > 0 ? Math.max(...logs.map(l => l.id)) + 1 : 1,
//                     operator_name: dayShiftOperator,
//                     action: item.action,
//                     description: item.description,
//                     timestamp: `${dateStr}T${pad2(endCursor.getHours())}:${pad2(endCursor.getMinutes())}:00`,
//                     shift: 'день',
//                     total_minutes: duration,
//                     routine: true,
//                     created_at: new Date().toISOString()
//                 };
//                 logs.push(log);
//                 existingRoutineKeys.add(key);
//                 addedCount++;
//             }
//             endCursor = new Date(endCursor.getTime() + duration * 60000);
//         });

//         // Ночная смена: начинается в 20:00 текущего дня, заканчивается в 8:00 следующего дня
//         const nightShiftOperator = OPERATOR_NAMES[Math.floor(Math.random() * OPERATOR_NAMES.length)];
//         const nextDay = new Date(day);
//         nextDay.setDate(day.getDate() + 1);
//         const nextDateStr = `${nextDay.getFullYear()}-${pad2(nextDay.getMonth() + 1)}-${pad2(nextDay.getDate())}`;

//         // Начало ночной смены: 20:00 текущего дня
//         let nightCursor = new Date(day);
//         nightCursor.setHours(20, 0, 0, 0);
//         ROUTINE_START_ACTIONS.forEach(item => {
//             const key = `${item.action}|${dateStr}|ночь`;
//             const duration = randomDuration(item.min, item.max);
//             if (!existingRoutineKeys.has(key)) {
//                 const log = {
//                     id: logs.length > 0 ? Math.max(...logs.map(l => l.id)) + 1 : 1,
//                     operator_name: nightShiftOperator,
//                     action: item.action,
//                     description: item.description,
//                     timestamp: `${dateStr}T${pad2(nightCursor.getHours())}:${pad2(nightCursor.getMinutes())}:00`,
//                     shift: 'ночь',
//                     total_minutes: duration,
//                     routine: true,
//                     created_at: new Date().toISOString()
//                 };
//                 logs.push(log);
//                 existingRoutineKeys.add(key);
//                 addedCount++;
//             }
//             nightCursor = new Date(nightCursor.getTime() + duration * 60000);
//         });

//         // Конец ночной смены: задачи заканчиваются ровно к 8:00 следующего дня
//         const nightEndDurations = ROUTINE_END_ACTIONS.map(item => randomDuration(item.min, item.max));
//         const totalNightEndDuration = nightEndDurations.reduce((a, b) => a + b, 0);
//         let nightEndCursor = new Date(nextDay);
//         nightEndCursor.setHours(8, 0, 0, 0);
//         nightEndCursor = new Date(nightEndCursor.getTime() - totalNightEndDuration * 60000);
//         ROUTINE_END_ACTIONS.forEach((item, idx) => {
//             const key = `${item.action}|${nextDateStr}|ночь`;
//             const duration = nightEndDurations[idx];
//             if (!existingRoutineKeys.has(key)) {
//                 const log = {
//                     id: logs.length > 0 ? Math.max(...logs.map(l => l.id)) + 1 : 1,
//                     operator_name: nightShiftOperator,
//                     action: item.action,
//                     description: item.description,
//                     timestamp: `${nextDateStr}T${pad2(nightEndCursor.getHours())}:${pad2(nightEndCursor.getMinutes())}:00`,
//                     shift: 'ночь',
//                     total_minutes: duration,
//                     routine: true,
//                     created_at: new Date().toISOString()
//                 };
//                 logs.push(log);
//                 existingRoutineKeys.add(key);
//                 addedCount++;
//             }
//             nightEndCursor = new Date(nightEndCursor.getTime() + duration * 60000);
//         });
//     }

//     if (addedCount > 0) {
//         writeOperatorLogs(logs);
//         console.log(`[Operator Logs] Добавлено ${addedCount} рутинных записей`);
//     }
// };

// Миграция: удаляем старые System-записи без актуальной (случайной) длительности и перезагружаем их заново
const migrateMissingTotalMinutes = async () => {
    const allLogs = readOperatorLogs();
    const staleCount = allLogs.filter(l => l.operator_name === 'System' && l.duration_version !== 2).length;
    if (staleCount > 0) {
        console.log(`[Operator Logs] Миграция: найдено ${staleCount} системных записей со старой длительностью, пересоздаём...`);
        const cleaned = allLogs.filter(l => !(l.operator_name === 'System' && l.duration_version !== 2));
        writeOperatorLogs(cleaned);
        await loadHistoricalData();
    }
};

// Загружаем исторические данные при запуске (только если файл пустой)
const logs = readOperatorLogs();
if (logs.length === 0) {
    console.log('[Operator Logs] Загрузка исторических данных...');
    loadHistoricalData(); // .then(() => generateRoutineLogs());
} else {
    migrateMissingTotalMinutes()
        .then(() => loadNewData());
        // .then(() => generateRoutineLogs());
}

// Периодическое обновление новых данных каждые 5 минут
setInterval(() => {
    loadNewData(); // .then(() => generateRoutineLogs());
}, 5 * 60 * 1000);

// GET /api/operator-logs - получить логи с фильтрацией по времени
app.get('/api/operator-logs', authenticateToken, async (req, res) => {
    const { from, to } = req.query;
    if (!from || !to) {
        res.status(400).json({ error: 'Требуются параметры from и to' });
        return;
    }
    try {
        const logs = readOperatorLogs();
        const fromDate = new Date(from);
        const toDate = new Date(to);
        
        const filteredLogs = logs.filter(log => {
            const logDate = new Date(log.timestamp);
            return logDate >= fromDate && logDate <= toDate;
        }).sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
        
        res.json({ rows: filteredLogs });
    } catch (err) {
        console.error('Ошибка /api/operator-logs:', err);
        res.status(500).json({ error: err.message });
    }
});

// POST /api/operator-logs - создать новый лог
app.post('/api/operator-logs', authenticateToken, async (req, res) => {
    const { operator_name, action, description, timestamp, shift, total_minutes } = req.body;
    if (!operator_name || !action || !timestamp || !shift) {
        res.status(400).json({ error: 'Требуются operator_name, action, timestamp и shift' });
        return;
    }
    try {
        const logs = readOperatorLogs();
        const newLog = {
            id: logs.length > 0 ? Math.max(...logs.map(l => l.id)) + 1 : 1,
            operator_name,
            action,
            description,
            timestamp,
            shift,
            total_minutes: total_minutes !== undefined && total_minutes !== null && total_minutes !== '' ? Number(total_minutes) : null,
            created_at: new Date().toISOString()
        };
        logs.push(newLog);
        writeOperatorLogs(logs);
        res.json({ id: newLog.id });
    } catch (err) {
        console.error('Ошибка POST /api/operator-logs:', err);
        res.status(500).json({ error: err.message });
    }
});

// PUT /api/operator-logs/:id - обновить лог
app.put('/api/operator-logs/:id', authenticateToken, async (req, res) => {
    const id = parseInt(req.params.id);
    const { operator_name, action, description, timestamp, shift, total_minutes } = req.body;
    if (!operator_name || !action || !timestamp || !shift) {
        res.status(400).json({ error: 'Требуются operator_name, action, timestamp и shift' });
        return;
    }
    try {
        const logs = readOperatorLogs();
        const index = logs.findIndex(l => l.id === id);
        if (index === -1) {
            res.status(404).json({ error: 'Лог не найден' });
            return;
        }
        const normalizedMinutes = total_minutes !== undefined && total_minutes !== null && total_minutes !== '' ? Number(total_minutes) : null;
        logs[index] = { ...logs[index], operator_name, action, description, timestamp, shift, total_minutes: normalizedMinutes };
        writeOperatorLogs(logs);
        res.json({ message: 'Лог обновлен' });
    } catch (err) {
        console.error('Ошибка PUT /api/operator-logs:', err);
        res.status(500).json({ error: err.message });
    }
});

// DELETE /api/operator-logs/:id - удалить лог
app.delete('/api/operator-logs/:id', authenticateToken, async (req, res) => {
    const id = parseInt(req.params.id);
    try {
        const logs = readOperatorLogs();
        const index = logs.findIndex(l => l.id === id);
        if (index === -1) {
            res.status(404).json({ error: 'Лог не найден' });
            return;
        }
        logs.splice(index, 1);
        writeOperatorLogs(logs);
        res.json({ message: 'Лог удален' });
    } catch (err) {
        console.error('Ошибка DELETE /api/operator-logs:', err);
        res.status(500).json({ error: err.message });
    }
});

// HTTP сервер (для разработки)
app.listen(PORT, '0.0.0.0', () => {
    console.log(`[WareCore Backend] HTTP сервер запущен на порту ${PORT}`);
    console.log(`[WareCore Backend] Откройте http://localhost:${PORT} для доступа к приложению`);
    console.log(`[WareCore Backend] Публичный доступ: http://172.30.1.249:${PORT}`);
});

// HTTPS сервер (отключен - sslOptions не определён)
// https.createServer(sslOptions, app).listen(HTTPS_PORT, '0.0.0.0', () => {
//     console.log(`[WareCore Backend] HTTPS сервер запущен на порту ${HTTPS_PORT}`);
//     console.log(`[WareCore Backend] Откройте https://localhost:${HTTPS_PORT} для доступа к приложению`);
//     console.log(`[WareCore Backend] Публичный доступ: https://172.30.1.249:${HTTPS_PORT}`);
// });

// Telegram-бот вынесен в отдельный процесс: см. папку telegram-bot/

process.on('uncaughtException', err => {
    console.error('[WareCore Backend] Необработанное исключение:', err);
});
process.on('unhandledRejection', err => {
    console.error('[WareCore Backend] Необработанный отказ промиса:', err);
});
