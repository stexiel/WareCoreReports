require('dotenv').config({ path: __dirname + '/.env' });
const { TelegramBot } = require('node-telegram-bot-api');
const { queryWMS } = require('./db');

const TELEGRAM_TOKEN = process.env.TELEGRAM_TOKEN;

if (!TELEGRAM_TOKEN) {
    console.error('[Telegram Bot] Не задан TELEGRAM_TOKEN в telegram-bot/.env');
    process.exit(1);
}

process.on('uncaughtException', (err) => {
    console.error('[Telegram Bot] uncaughtException:', err && err.message ? err.message : err);
});
process.on('unhandledRejection', (err) => {
    console.error('[Telegram Bot] unhandledRejection:', err && err.message ? err.message : err);
});

const DB_ERROR_TEXT = '⚠️ *Обратитесь к оператору WCS*\n\nСвязь с базой данных оборвана.';

const bot = new TelegramBot(TELEGRAM_TOKEN, { polling: true });

function setupCommands() {
    bot.setMyCommands([
        { command: 'start', description: 'Запустить бота' },
        { command: 'receipt', description: 'Показать документы приемки' }
    ]).catch(err => console.error('[Telegram Bot] Ошибка setMyCommands:', err.message));
}

setupCommands();
setInterval(setupCommands, 60 * 1000);

// ===== Работа с данными приемки =====


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

// ===== Кэш в памяти для коротких callback_data =====

let docsCache = []; // [{ docNumber, completed }]
let currentPeriod = 'twoDays';
const detailsCache = new Map(); // docNumber -> details[]
const DOC_LIST_LIMIT = 60;

const FILTER_CODES = { a: 'Все', i: 'Импорт', g: 'Получен', v: 'Выставлено' };

function pad(str, len) {
    str = String(str ?? '');
    if (str.length >= len) return str.slice(0, len);
    return str + ' '.repeat(len - str.length);
}

function buildListKeyboard(docs, period) {
    const rows = docs.slice(0, DOC_LIST_LIMIT).map((doc, idx) => ([{
        text: `${doc.completed ? '✅' : '❌'} ${doc.docNumber}: ${doc.completed ? 'Выставлено' : 'Не выставлено'}`,
        callback_data: `d:${idx}`
    }]));
    rows.push([{ text: '🔄 Обновить', callback_data: 'refresh' }]);
    return { inline_keyboard: rows };
}

// Уменьшено с 40 до 15: с артикулом/партией/кол-вом каждая паллета занимает
// несколько строк, и на 40 штук на странице сообщение превышало лимит Telegram
// в 4096 символов (ошибка MESSAGE_TOO_LONG).
const PAGE_SIZE = 15;

function buildDetailKeyboard(idx, activeCode, page, totalPages) {
    const keyboard = [];

    if (totalPages > 1) {
        const pageRow = [];
        if (page > 0) {
            pageRow.push({ text: '<', callback_data: `f:${idx}::${page - 1}` });
        }
        pageRow.push({ text: `${page + 1}/${totalPages}`, callback_data: 'noop' });
        if (page < totalPages - 1) {
            pageRow.push({ text: '>', callback_data: `f:${idx}::${page + 1}` });
        }
        keyboard.push(pageRow);
    }

    keyboard.push([{ text: '⬅ Список документов', callback_data: 'list' }]);
    return { inline_keyboard: keyboard };
}

function renderDetailText(doc, details, filterCode, page) {
    const filtered = details;
    const SEP = '───────────────';

    let text = `📄 Документ: ${doc.docNumber}\n`;
    text += doc.completed ? '✅ Выставлено полностью' : '❌ Не выставлено полностью';
    text += `\n${SEP}`;

    const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
    const safePage = Math.min(page, totalPages - 1);
    const rows = filtered.slice(safePage * PAGE_SIZE, safePage * PAGE_SIZE + PAGE_SIZE);

    rows.forEach((item, i) => {
        const num = safePage * PAGE_SIZE + i + 1;
        const icon = item.palletState === 'Выставлено' ? '✅' : '❌';
        text += `\n${num}. 📦 ${item.containerNumber}`;
        text += `\n   Статус: ${icon} ${item.palletState}`;
        text += `\n   Артикул: ${item.article || '—'}`;
        text += `\n   Партия: ${item.batchNumber || '—'}`;
        text += `\n   Кол-во коробок: ${item.quantity ?? '—'}`;
        text += `\n${SEP}`;
    });

    if (totalPages > 1) {
        text += `\n📍 Страница ${safePage + 1} из ${totalPages}`;
    }

    return { text, totalPages, page: safePage };
}

// ===== Команды =====

bot.onText(/^\/start/, async (msg) => {
    const chatId = msg.chat.id;
    await bot.sendMessage(
        chatId,
        '👋 *WareCore Reports Bot*\n\nНажмите кнопку ниже, чтобы получить список документов приемки.',
        {
            parse_mode: 'Markdown',
            reply_markup: { inline_keyboard: [[{ text: '📋 Приемка', callback_data: 'receipt' }]] }
        }
    );
});

bot.onText(/^\/receipt/, async (msg) => {
    await sendDocumentsList(msg.chat.id, null, currentPeriod);
});

function buildListText(period) {
    if (docsCache.length === 0) {
        return '📋 *Документы приемки*\n\nНет данных.';
    }
    const shown = Math.min(docsCache.length, DOC_LIST_LIMIT);
    const truncatedNote = docsCache.length > DOC_LIST_LIMIT ? ` (показано ${shown} из ${docsCache.length})` : '';
    return `📋 *Документы приемки* — всего ${docsCache.length}${truncatedNote}\n\nВыберите документ:`;
}

async function sendDocumentsList(chatId, messageId, period) {
    currentPeriod = period || currentPeriod;
    const loadingText = '⏳ Загрузка документов приемки...';
    let sentMsg;
    if (messageId) {
        await bot.editMessageText(loadingText, { chat_id: chatId, message_id: messageId }).catch(() => {});
    } else {
        sentMsg = await bot.sendMessage(chatId, loadingText);
    }

    try {
        docsCache = await fetchReceiptDocuments();
        detailsCache.clear();

        const text = buildListText(currentPeriod);
        const options = { parse_mode: 'Markdown', reply_markup: buildListKeyboard(docsCache, currentPeriod) };

        if (messageId) {
            await bot.editMessageText(text, { chat_id: chatId, message_id: messageId, ...options });
        } else {
            await bot.editMessageText(text, { chat_id: chatId, message_id: sentMsg.message_id, ...options });
        }
    } catch (err) {
        console.error('[Telegram Bot] Ошибка загрузки документов:', err.message);
        const options = { parse_mode: 'Markdown', reply_markup: { inline_keyboard: [[{ text: '🔄 Повторить', callback_data: 'refresh' }]] } };
        if (messageId) {
            await bot.editMessageText(DB_ERROR_TEXT, { chat_id: chatId, message_id: messageId, ...options }).catch(() => {});
        } else {
            await bot.editMessageText(DB_ERROR_TEXT, { chat_id: chatId, message_id: sentMsg.message_id, ...options }).catch(() => {});
        }
    }
}

async function sendDocumentDetail(chatId, messageId, idx, filterCode, page = 0) {
    const doc = docsCache[idx];
    if (!doc) {
        await bot.editMessageText('⚠️ Список устарел. Нажмите «Приемка» ещё раз.', {
            chat_id: chatId,
            message_id: messageId,
            reply_markup: { inline_keyboard: [[{ text: '📋 Приемка', callback_data: 'receipt' }]] }
        }).catch(() => {});
        return;
    }

    await bot.editMessageText('⏳ Загрузка деталей документа...', { chat_id: chatId, message_id: messageId }).catch(() => {});

    try {
        let details = detailsCache.get(doc.docNumber);
        if (!details) {
            details = await fetchReceiptDetail(doc.docNumber);
            detailsCache.set(doc.docNumber, details);
        }

        const { text, totalPages, page: safePage } = renderDetailText(doc, details, filterCode, page);
        await bot.editMessageText(text, {
            chat_id: chatId,
            message_id: messageId,
            reply_markup: buildDetailKeyboard(idx, filterCode, safePage ?? 0, totalPages)
        });
    } catch (err) {
        console.error('[Telegram Bot] Ошибка загрузки деталей:', err.message);
        const isTelegramError = err.message && err.message.startsWith('ETELEGRAM');
        const errorText = isTelegramError
            ? '⚠️ Не удалось отобразить сообщение (слишком много данных). Попробуйте другую страницу.'
            : DB_ERROR_TEXT;
        await bot.editMessageText(errorText, {
            chat_id: chatId,
            message_id: messageId,
            parse_mode: isTelegramError ? undefined : 'Markdown',
            reply_markup: { inline_keyboard: [[{ text: '⬅ Список документов', callback_data: 'list' }]] }
        }).catch(() => {});
    }
}

// ===== Обработка нажатий инлайн-кнопок =====

bot.on('callback_query', async (query) => {
    const chatId = query.message.chat.id;
    const messageId = query.message.message_id;
    const data = query.data || '';

    try {
        if (data === 'receipt' || data === 'refresh') {
            await bot.answerCallbackQuery(query.id);
            await sendDocumentsList(chatId, messageId, currentPeriod);
            return;
        }

        if (data === 'list') {
            await bot.answerCallbackQuery(query.id);
            await bot.editMessageText(buildListText(currentPeriod), {
                chat_id: chatId,
                message_id: messageId,
                parse_mode: 'Markdown',
                reply_markup: buildListKeyboard(docsCache, currentPeriod)
            });
            return;
        }

        if (data === 'noop') {
            await bot.answerCallbackQuery(query.id);
            return;
        }

        if (data.startsWith('d:')) {
            const idx = parseInt(data.slice(2), 10);
            await bot.answerCallbackQuery(query.id);
            await sendDocumentDetail(chatId, messageId, idx, 'a', 0);
            return;
        }

        if (data.startsWith('f:')) {
            const [, idxStr, code, pageStr] = data.split(':');
            const idx = parseInt(idxStr, 10);
            const page = parseInt(pageStr, 10) || 0;
            await bot.answerCallbackQuery(query.id);
            await sendDocumentDetail(chatId, messageId, idx, code, page);
            return;
        }

        await bot.answerCallbackQuery(query.id);
    } catch (err) {
        console.error('[Telegram Bot] Ошибка обработки callback_query:', err.message);
        try { await bot.answerCallbackQuery(query.id, { text: 'Ошибка, попробуйте ещё раз' }); } catch (e) {}
    }
});

bot.on('polling_error', (err) => {
    console.error('[Telegram Bot] Polling error:', err.message);
});

console.log('[Telegram Bot] Запущен (документы приемки, без отображения времени)');
