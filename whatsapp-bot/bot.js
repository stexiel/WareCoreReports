require('dotenv').config({ path: __dirname + '/.env' });
const { Client, LocalAuth } = require('whatsapp-web.js');
const cron = require('node-cron');
const qrcode = require('qrcode-terminal');
const { queryWMS, queryWCS } = require('./db');

process.on('uncaughtException', (err) => {
    console.error('[WhatsApp Bot] uncaughtException:', err && err.message ? err.message : err);
});
process.on('unhandledRejection', (err) => {
    console.error('[WhatsApp Bot] unhandledRejection:', err && err.message ? err.message : err);
});

const DB_ERROR_TEXT = '⚠️ *Обратитесь к оператору WCS*\n\nСвязь с базой данных оборвана.';
const DOC_LIST_LIMIT = 60;
const DETAIL_CHUNK_SIZE = 100; // строк на одно сообщение (защита от слишком длинных сообщений)

// Список разрешённых номеров (chatId, формат "79991234567@c.us"). Пусто = разрешены все.
const ALLOWED_CHATS = (process.env.ALLOWED_CHATS || '')
    .split(',')
    .map(s => s.trim())
    .filter(Boolean);

// Группы для рассылки коробочного остатка (chatId формат "123456789@g.us")
const ALLOWED_GROUPS = (process.env.ALLOWED_GROUPS || '')
    .split(',')
    .map(s => s.trim())
    .filter(Boolean);

// Группа для мгновенных уведомлений об ошибках штабелеров/шаттлов
const ALERT_GROUP = (process.env.ALERT_GROUP || '').trim();

// Карта оборудования -> название для уведомления об ошибке
const EQUIPMENT_TO_LABEL = {
    'A01PLC': 'Шаттл A01',
    'A02PLC': 'Шаттл A02',
    'SC0101': 'Штабелер 1', 'SC0201': 'Штабелер 2', 'SC0301': 'Штабелер 3',
    'SC0401': 'Штабелер 4', 'SC0501': 'Штабелер 5', 'SC0601': 'Штабелер 6',
    'SC0701': 'Штабелер 7', 'SC0801': 'Штабелер 8', 'SC0901': 'Штабелер 9',
    'SC1001': 'Штабелер 10'
};

// Перевод кодов ошибок (fault_code -> описание на русском). Если код не найден - показываем как есть.
const FAULT_CODE_TRANSLATIONS = {
    'DH01': 'Выход за край - левая сторона платформы',
    'DH02': 'Выход за край - правая сторона платформы',
    'DS01': 'Превышена высота - слева',
    'DS02': 'Превышена высота - справа',
    'DS03': 'Превышена ширина - слева спереди',
    'DS04': 'Превышена ширина - слева сзади',
    'DS05': 'Превышена ширина - справа спереди',
    'DS06': 'Превышена ширина - справа сзади',
    'ES02': 'Сработал аварийный стоп 2', 'ES03': 'Сработал аварийный стоп 3',
    'ES04': 'Сработал аварийный стоп 4', 'ES05': 'Сработал аварийный стоп 5',
    'ES06': 'Сработал аварийный стоп 6', 'ES07': 'Сработал аварийный стоп 7',
    'ES08': 'Сработал аварийный стоп 8', 'ES09': 'Сработал аварийный стоп 9',
    'ES10': 'Сработал аварийный стоп 10', 'ES11': 'Сработал аварийный стоп 11',
    'ES12': 'Сработал аварийный стоп 12', 'ES13': 'Сработал аварийный стоп 13',
    'ES14': 'Сработал аварийный стоп 14', 'ES15': 'Сработал аварийный стоп 15',
    'ES16': 'Сработал аварийный стоп 16', 'ES17': 'Сработал аварийный стоп 17',
    'ES18': 'Сработал аварийный стоп 18', 'ES19': 'Сработал аварийный стоп 19',
    'ES20': 'Сработал аварийный стоп 20', 'ES21': 'Сработал аварийный стоп 21',
    'ES22': 'Сработал аварийный стоп 22', 'ES23': 'Сработал аварийный стоп 23',
    'ES24': 'Сработал аварийный стоп 24', 'ES25': 'Сработал аварийный стоп 25',
    'ES26': 'Сработал аварийный стоп 26', 'ES27': 'Сработал аварийный стоп 27',
    'ES28': 'Сработал аварийный стоп 28',
    'M001': 'Пауза движения',
    'M002': 'Не в автоматическом режиме (блокировка)',
    'M003': 'Не закрыта дверь шкафа управления',
    'M004': 'Вилы не в среднем положении при движении',
    'M005': 'Открыта дверь ограждения',
    'M010': 'Авария движения (резерв)',
    'MF01': 'Авария преобразователя вил',
    'MF04': 'Ошибка энкодера вил в среднем положении',
    'MF09': 'Координата вил вне допустимого диапазона',
    'MF13': 'Микроподъём не в верхнем/нижнем положении при движении вил',
    'MF16': 'Вилы не выровнены по горизонтали снаружи',
    'MFF1': 'Ошибка состояния интерфейса при движении вил',
    'MFF11': 'Авария вил (резерв)', 'MFF12': 'Авария вил (резерв)',
    'MFF3': 'Авария вил (резерв)', 'MFF4': 'Авария вил (резерв)', 'MFF5': 'Авария вил (резерв)',
    'ML01': 'Неисправность преобразователя подъёма',
    'ML02': 'Неисправность датчика расстояния подъёма',
    'ML03': 'Достигнут верхний предел подъёма',
    'ML12': 'Авария тормоза подъёма',
    'ML14': 'Достигнут нижний предел подъёма',
    'ML15': 'Ошибка резервного положения подъёма',
    'MML1': 'Превышен предел микроподъёма',
    'MML5': 'Нет выравнивания по горизонтали (микроподъём)',
    'MR02': 'Сработала защита фаз CPL2', 'MR03': 'Сработала защита фаз CPL3',
    'MR04': 'Сработала защита фаз CPL4', 'MR05': 'Сработала защита фаз CPL5',
    'MR06': 'Сработала защита фаз CPL6', 'MR07': 'Сработала защита фаз CPL7',
    'MR08': 'Сработала защита фаз CPL8', 'MR09': 'Сработала защита фаз CPL9',
    'MT01': 'Авария преобразователя привода движения',
    'MT02': 'Неисправность датчика расстояния (движение)',
    'MT03': 'Достигнут передний предел хода',
    'MT05': 'Координата движения вне диапазона',
    'MT07': 'Превышение хода',
    'MT09': 'Аномальная скорость движения',
    'MT13': 'Достигнут задний предел хода',
    'P001': 'Не подано основное питание',
    'P002': 'Состояние аварийного стопа (питание)',
    'P010': 'Сработала защита фаз (питание)',
    'P011': 'Сработала защита ослабления троса',
    'P012': 'Сработал ограничитель скорости',
    'P013': 'Защита ослабления троса ограничителя скорости',
    'S003': 'Сбой связи SC и MainA',
    'TC07': 'Неверная команда загрузки',
    'TD01': 'Место назначения разгрузки занято',
    'TD03': 'Нет сигнала подтверждения на интерфейсе передачи (разгрузка)',
    'TD04': 'Сигнал подтверждения не вернулся в исходное состояние после разгрузки',
    'TD05': 'Тайм-аут разгрузки',
    'TD06': 'Логическая занятость места разгрузки',
    'TI01': 'Вилы не в среднем положении (инициализация)',
    'TI04': 'Авария инициализации (резерв)',
    'TM01': 'Тайм-аут основного процесса',
    'TM02': 'Занятость места при отсутствии задачи',
    'TP01': 'Взятие пустого места (загрузка)',
    'TP02': 'Нет сигнала подтверждения перед загрузкой',
    'TP03': 'Сигнал подтверждения не вернулся в исходное состояние после загрузки',
    'TP04': 'Тайм-аут загрузки',
    'TP05': 'Авария загрузки (резерв)', 'TP06': 'Авария загрузки (резерв)',
    'TW01': 'Тайм-аут канала команд WCS'
};

function translateFaultCode(faultCode, faultDesc) {
    if (faultCode && FAULT_CODE_TRANSLATIONS[faultCode]) {
        return `${FAULT_CODE_TRANSLATIONS[faultCode]} (${faultCode})`;
    }
    if (faultCode) return faultCode;
    return faultDesc || 'неизвестная ошибка';
}

const client = new Client({
    authStrategy: new LocalAuth({ dataPath: __dirname + '/.wwebjs_auth' }),
    puppeteer: {
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox'],
        executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'
    }
});

// ===== Работа с данными приемки (те же запросы, что и в telegram-bot) =====

async function fetchReceiptDocuments() {
    const result = await queryWMS(
        `SELECT TOP 200
                bill_no AS docNumber,
                start_time AS startTime,
                end_time AS endTime
         FROM in_receiving_bill
         WHERE is_deleted = 0
           AND start_time >= DATEADD(day, -2, GETDATE())
         ORDER BY start_time ASC`
    );
    return result.recordset.map(r => ({
        docNumber: r.docNumber,
        completed: !!r.endTime
    }));
}

async function fetchReceiptDetail(docNumber) {
    const result = await queryWMS(
        `SELECT r.container_code AS containerNumber,
                r.qty AS quantity,
                m.code AS article,
                l.lot_produce AS batchNumber,
                s.putaway_time,
                s.pick_time,
                CASE
                    WHEN s.putaway_time IS NOT NULL THEN 'Выставлено'
                    ELSE 'Не выставлено'
                END AS palletState
         FROM in_receiving_bill b
         JOIN in_receiving_bill_detail d ON d.receiving_bill_id = b.id
         JOIN in_receiving_record r ON r.receiving_bill_detail_id = d.id AND r.status = 'Completed'
         LEFT JOIN inv_lot l ON l.id = d.lot_id
         LEFT JOIN bm_material m ON m.id = l.material_id
         LEFT JOIN inv_stock s ON s.container_code = r.container_code AND s.lot_id = d.lot_id
         WHERE b.bill_no = @docNumber
         ORDER BY r.container_code`,
        { docNumber }
    );
    return result.recordset.map(r => ({
        containerNumber: r.containerNumber,
        quantity: r.quantity,
        article: r.article,
        batchNumber: r.batchNumber,
        palletState: r.palletState
    }));
}

// ===== Состояние по каждому чату (номеру WhatsApp) =====

const sessions = new Map(); // chatId -> { docs: [], detailsCache: Map }

function getSession(chatId) {
    if (!sessions.has(chatId)) {
        sessions.set(chatId, { docs: [], detailsCache: new Map() });
    }
    return sessions.get(chatId);
}

function buildListText(docs) {
    if (docs.length === 0) {
        return '📋 *Документы приемки*\n\nНет данных за последние 2 дня.';
    }
    const shown = Math.min(docs.length, DOC_LIST_LIMIT);
    const truncatedNote = docs.length > DOC_LIST_LIMIT ? ` (показано ${shown} из ${docs.length})` : '';
    let text = `📋 *Документы приемки* — всего ${docs.length}${truncatedNote}\n\n`;
    docs.slice(0, DOC_LIST_LIMIT).forEach((doc) => {
        const icon = doc.completed ? '✅' : '❌';
        text += `${icon} ${doc.docNumber}: ${doc.completed ? 'Выставлено' : 'Не выставлено'}\n`;
    });
    text += `\nНапишите "список" или "0", чтобы обновить.`;
    return text;
}

function buildDetailText(doc, details) {
    let text = `📦 *${doc.docNumber}* — ${doc.completed ? '✅ Выставлено' : '❌ Не выставлено'}\n`;
    details.forEach(item => {
        const icon = item.palletState === 'Выставлено' ? '✅' : '❌';
        text += `${icon} ${item.containerNumber}: ${item.palletState}\n`;
    });
    text += `\nОтветьте *0*, чтобы вернуться к списку документов.`;
    return text;
}

// ===== Ежедневная рассылка по группам =====

const GROUP_DAILY_TEXT = 'Ассалаумағалейкум, коробочный остаток барма?';

async function sendDailyGroupMessages() {
    for (const groupId of ALLOWED_GROUPS) {
        try {
            await client.sendMessage(groupId, GROUP_DAILY_TEXT);
            console.log(`[WhatsApp Bot] Сообщение в группу ${groupId}`);
        } catch (err) {
            console.error(`[WhatsApp Bot] Ошибка отправки в группу ${groupId}:`, err.message);
        }
    }
}

function scheduleGroupMessages() {
    if (ALLOWED_GROUPS.length === 0) {
        console.log('[WhatsApp Bot] ALLOWED_GROUPS пуст — рассылка по группам отключена');
        return;
    }
    // 08:00 и 20:00 каждый день
    cron.schedule('0 8 * * *', sendDailyGroupMessages, { timezone: 'Asia/Almaty' });
    cron.schedule('0 20 * * *', sendDailyGroupMessages, { timezone: 'Asia/Almaty' });
    console.log('[WhatsApp Bot] Рассылка по группам запланирована на 08:00 и 20:00');
}

// ===== Мгновенные уведомления об ошибках штабелеров/шаттлов =====

const FAULT_POLL_INTERVAL_MS = 5000;
let lastFaultCheckTime = null;

async function pollEquipmentFaults() {
    if (!ALERT_GROUP) return;
    try {
        if (!lastFaultCheckTime) {
            // При первом запуске просто фиксируем текущий момент, чтобы не слать историю
            lastFaultCheckTime = new Date();
            return;
        }
        const codes = Object.keys(EQUIPMENT_TO_LABEL);
        const inClause = codes.map((c, i) => `@code${i}`).join(',');
        const params = { from: lastFaultCheckTime };
        codes.forEach((c, i) => { params['code' + i] = c; });

        const result = await queryWCS(
            `SELECT eqpt_code, begin_time, fault_code, fault_desc
             FROM eqpt_fault_record
             WHERE eqpt_code IN (${inClause})
               AND begin_time > @from
             ORDER BY begin_time ASC`,
            params
        );

        if (result.recordset.length > 0) {
            for (const row of result.recordset) {
                const label = EQUIPMENT_TO_LABEL[row.eqpt_code] || row.eqpt_code;
                const time = new Date(row.begin_time).toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
                const errorDesc = translateFaultCode(row.fault_code, row.fault_desc);
                await client.sendMessage(ALERT_GROUP, `⚠️ ${label} в ошибке: ${errorDesc} (${time})`).catch(err => {
                    console.error('[WhatsApp Bot] Ошибка отправки уведомления об ошибке:', err.message);
                });
            }
            const maxTime = result.recordset.reduce((max, r) => {
                const t = new Date(r.begin_time);
                return t > max ? t : max;
            }, lastFaultCheckTime);
            lastFaultCheckTime = maxTime;
        }
    } catch (err) {
        console.error('[WhatsApp Bot] Ошибка опроса ошибок оборудования:', err.message);
    }
}

function startFaultPolling() {
    if (!ALERT_GROUP) {
        console.log('[WhatsApp Bot] ALERT_GROUP пуст — мгновенные уведомления об ошибках отключены');
        return;
    }
    setInterval(pollEquipmentFaults, FAULT_POLL_INTERVAL_MS);
    console.log(`[WhatsApp Bot] Мгновенные уведомления об ошибках штабелеров/шаттлов включены (опрос каждые ${FAULT_POLL_INTERVAL_MS / 1000}с)`);
}

function chunkLines(text, chunkSize) {
    const lines = text.split('\n');
    const chunks = [];
    for (let i = 0; i < lines.length; i += chunkSize) {
        chunks.push(lines.slice(i, i + chunkSize).join('\n'));
    }
    return chunks;
}

async function sendChunked(chatId, text) {
    if (text.length <= 3500) {
        await client.sendMessage(chatId, text);
        return;
    }
    const chunks = chunkLines(text, DETAIL_CHUNK_SIZE);
    for (const chunk of chunks) {
        await client.sendMessage(chatId, chunk);
    }
}

async function sendDocumentsList(chatId) {
    await client.sendMessage(chatId, '⏳ Загрузка документов приемки...');
    const session = getSession(chatId);
    try {
        session.docs = await fetchReceiptDocuments();
        session.detailsCache.clear();
        await sendChunked(chatId, buildListText(session.docs));
    } catch (err) {
        console.error('[WhatsApp Bot] Ошибка загрузки документов:', err.message);
        await client.sendMessage(chatId, DB_ERROR_TEXT);
    }
}

async function sendDocumentDetail(chatId, idx) {
    const session = getSession(chatId);
    const doc = session.docs[idx];
    if (!doc) {
        await client.sendMessage(chatId, '⚠️ Список устарел или неверный номер. Напишите "список", чтобы обновить.');
        return;
    }

    await client.sendMessage(chatId, '⏳ Загрузка деталей документа...');
    try {
        let details = session.detailsCache.get(doc.docNumber);
        if (!details) {
            details = await fetchReceiptDetail(doc.docNumber);
            session.detailsCache.set(doc.docNumber, details);
        }
        await sendChunked(chatId, buildDetailText(doc, details));
    } catch (err) {
        console.error('[WhatsApp Bot] Ошибка загрузки деталей:', err.message);
        await client.sendMessage(chatId, DB_ERROR_TEXT);
    }
}

// ===== Обработка сообщений =====

const LIST_TRIGGERS = /^(старт|start|прием|приёмка|приемка|меню|список|list|0)$/i;
const REFRESH_TRIGGERS = /^(обновить|update|refresh)$/i;

client.on('message', async (msg) => {
    if (msg.from === 'status@broadcast') return;

    const chatId = msg.from;
    const isGroup = chatId.endsWith('@g.us');
    const body = (msg.body || '').trim();

    // ==== Группы: бот не отвечает на команды, кроме служебной "id" для определения chatId группы ====
    if (isGroup) {
        if (/^id$/i.test(body)) {
            await client.sendMessage(chatId, `ID этой группы: ${chatId}`).catch(() => {});
        }
        return;
    }

    // ==== Личные чаты: работа с приёмкой ====
    if (ALLOWED_CHATS.length > 0 && !ALLOWED_CHATS.includes(chatId)) return;

    try {
        if (LIST_TRIGGERS.test(body) || REFRESH_TRIGGERS.test(body)) {
            await sendDocumentsList(chatId);
            return;
        }

        // Неизвестное сообщение - игнорируем, не отвечаем
        return;
    } catch (err) {
        console.error('[WhatsApp Bot] Ошибка обработки сообщения:', err.message);
        await client.sendMessage(chatId, DB_ERROR_TEXT).catch(() => {});
    }
});

client.on('qr', (qr) => {
    console.log('[WhatsApp Bot] Отсканируйте QR-код через WhatsApp (Настройки → Связанные устройства):');
    qrcode.generate(qr, { small: true });
});

client.on('ready', () => {
    console.log('[WhatsApp Bot] Запущен и готов к работе (документы приемки за 2 дня)');
    scheduleGroupMessages();
    startFaultPolling();
});

client.on('auth_failure', (msg) => {
    console.error('[WhatsApp Bot] Ошибка авторизации:', msg);
});

client.on('disconnected', (reason) => {
    console.error('[WhatsApp Bot] Отключён:', reason);
});

client.initialize();
