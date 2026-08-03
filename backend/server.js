require('dotenv').config();
const express = require('express');
const cors = require('cors');
const path = require('path');
const bcrypt = require('bcrypt');
const jwt = require('jsonwebtoken');
const rateLimit = require('express-rate-limit');
const crypto = require('crypto');
const https = require('https');
const fs = require('fs');
const { sql, queryWMS, queryWCS } = require('./db');

const app = express();
const PORT = parseInt(process.env.PORT || '3000', 10);
const HTTPS_PORT = parseInt(process.env.HTTPS_PORT || '3443', 10);
const JWT_SECRET = process.env.JWT_SECRET || 'your-secret-key-change-in-production';
const CSRF_SECRET = process.env.CSRF_SECRET || 'csrf-secret-change-in-production';

// SSL сертификаты
const sslOptions = {
  key: fs.readFileSync(path.join(__dirname, 'certs', 'key.pem')),
  cert: fs.readFileSync(path.join(__dirname, 'certs', 'cert.pem'))
};

// Хранилище CSRF токенов
const csrfTokens = new Map();

// Генерация CSRF токена
function generateCSRFToken() {
    return crypto.randomBytes(32).toString('hex');
}

app.use(cors({
    origin: ['http://localhost:5173', 'http://localhost:5174', 'http://localhost:3000', 'https://localhost:5173', 'https://localhost:5174', 'https://localhost:3000', 'http://92.38.29.140:5173', 'https://92.38.29.140:5173', 'http://92.38.29.140:3000', 'https://92.38.29.140:3000', 'http://92.38.29.140:3443', 'https://92.38.29.140:3443'],
    credentials: true
}));
app.use(express.json());

// Middleware для генерации CSRF токена
app.use((req, res, next) => {
    if (req.method === 'GET' && req.path === '/api/csrf-token') {
        const token = generateCSRFToken();
        csrfTokens.set(token, Date.now());
        res.json({ csrfToken: token });
        return;
    }
    next();
});

// Middleware для проверки CSRF токена
function validateCSRF(req, res, next) {
    if (req.method === 'GET' || req.method === 'HEAD' || req.method === 'OPTIONS') {
        return next();
    }

    const token = req.headers['x-csrf-token'];
    if (!token) {
        return res.status(403).json({ error: 'CSRF токен отсутствует' });
    }

    if (!csrfTokens.has(token)) {
        return res.status(403).json({ error: 'Недействительный CSRF токен' });
    }

    // Удаляем токен после использования (one-time use)
    csrfTokens.delete(token);
    next();
}

// Rate limiting middleware
const limiter = rateLimit({
    windowMs: 15 * 60 * 1000, // 15 минут
    max: 100, // максимум 100 запросов за окно
    message: { error: 'Слишком много запросов, попробуйте позже' }
});

const authLimiter = rateLimit({
    windowMs: 15 * 60 * 1000, // 15 минут
    max: 5, // максимум 5 попыток входа за окно
    message: { error: 'Слишком много попыток входа, попробуйте позже' }
});

app.use('/api/', limiter);
app.use('/api/login', authLimiter);

// Раздаём статику фронтенда
app.use(express.static(path.join(__dirname, '../frontend')));

// ===== АВТОРИЗАЦИЯ =====
// Хранилище пользователей с ролями
// Роли: admin - полный доступ, operator - все кроме админки, warehouseman - только приемка
const USERS = [
    { id: 1, username: 'stexiel', password: 'ASER2007', role: 'admin', hashed: false }
];

// Middleware для проверки токена
function authenticateToken(req, res, next) {
    const authHeader = req.headers['authorization'];
    const token = authHeader && authHeader.split(' ')[1];

    if (!token) {
        return res.status(401).json({ error: 'Требуется авторизация' });
    }

    try {
        const decoded = jwt.verify(token, JWT_SECRET);
        req.user = decoded;
        next();
    } catch (err) {
        return res.status(403).json({ error: 'Недействительный токен' });
    }
}

// Middleware для проверки роли admin
function requireAdmin(req, res, next) {
    if (req.user.role !== 'admin') {
        return res.status(403).json({ error: 'Требуются права администратора' });
    }
    next();
}

// POST /api/login - вход в систему
app.post('/api/login', validateCSRF, async (req, res) => {
    const { username, password } = req.body;

    if (!username || !password) {
        return res.status(400).json({ error: 'Требуются логин и пароль' });
    }

    const user = USERS.find(u => u.username === username);

    if (!user) {
        return res.status(401).json({ error: 'Неверный логин или пароль' });
    }

    // Проверяем пароль (с хешированием или без для совместимости)
    let passwordMatch = false;
    if (user.hashed) {
        passwordMatch = await bcrypt.compare(password, user.password);
    } else {
        passwordMatch = user.password === password;
        // Если пароль не хеширован, хешируем его для будущего использования
        if (passwordMatch) {
            user.password = await bcrypt.hash(password, 10);
            user.hashed = true;
        }
    }

    if (!passwordMatch) {
        return res.status(401).json({ error: 'Неверный логин или пароль' });
    }

    const token = jwt.sign(
        { 
            username: user.username, 
            role: user.role,
            id: user.id
        },
        JWT_SECRET,
        { expiresIn: '12h' }
    );

    res.json({ token, username: user.username, role: user.role, id: user.id });
});

// POST /api/logout - выход из системы
app.post('/api/logout', authenticateToken, (req, res) => {
    // JWT токены stateless, но на клиенте нужно удалить токен
    res.json({ message: 'Выход выполнен' });
});

// GET /api/me - информация о текущем пользователе
app.get('/api/me', authenticateToken, (req, res) => {
    res.json({ username: req.user.username, role: req.user.role, id: req.user.id });
});

// ===== УПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯМИ (только для admin) =====

// GET /api/users - получить всех пользователей
app.get('/api/users', authenticateToken, requireAdmin, (req, res) => {
    const users = USERS.map(u => ({
        id: u.id,
        username: u.username,
        role: u.role,
        password: u.password // для админа показываем пароль
    }));
    res.json(users);
});

// POST /api/users - создать пользователя
app.post('/api/users', authenticateToken, requireAdmin, validateCSRF, async (req, res) => {
    const { username, password, role } = req.body;

    if (!username || !password || !role) {
        return res.status(400).json({ error: 'Требуются username, password и role' });
    }

    if (!['admin', 'operator', 'warehouseman'].includes(role)) {
        return res.status(400).json({ error: 'Некорректная роль. Допустимые: admin, operator, warehouseman' });
    }

    if (USERS.find(u => u.username === username)) {
        return res.status(400).json({ error: 'Пользователь с таким именем уже существует' });
    }

    const newId = USERS.length > 0 ? Math.max(...USERS.map(u => u.id)) + 1 : 1;
    const hashedPassword = await bcrypt.hash(password, 10);
    const newUser = { id: newId, username, password: hashedPassword, role, hashed: true };
    USERS.push(newUser);

    res.json({ id: newId, username, role });
});

// PUT /api/users/:id - обновить пользователя
app.put('/api/users/:id', authenticateToken, requireAdmin, validateCSRF, async (req, res) => {
    const id = parseInt(req.params.id);
    const { username, password, role } = req.body;

    const userIndex = USERS.findIndex(u => u.id === id);
    if (userIndex === -1) {
        return res.status(404).json({ error: 'Пользователь не найден' });
    }

    // Нельзя изменить роль самого админа
    if (USERS[userIndex].username === 'stexiel' && role && role !== 'admin') {
        return res.status(400).json({ error: 'Нельзя изменить роль главного администратора' });
    }

    if (username) USERS[userIndex].username = username;
    if (password) {
        USERS[userIndex].password = await bcrypt.hash(password, 10);
        USERS[userIndex].hashed = true;
    }
    if (role) {
        if (!['admin', 'operator', 'warehouseman'].includes(role)) {
            return res.status(400).json({ error: 'Некорректная роль' });
        }
        USERS[userIndex].role = role;
    }

    res.json({ id: USERS[userIndex].id, username: USERS[userIndex].username, role: USERS[userIndex].role });
});

// DELETE /api/users/:id - удалить пользователя
app.delete('/api/users/:id', authenticateToken, requireAdmin, validateCSRF, (req, res) => {
    const id = parseInt(req.params.id);

    const userIndex = USERS.findIndex(u => u.id === id);
    if (userIndex === -1) {
        return res.status(404).json({ error: 'Пользователь не найден' });
    }

    // Нельзя удалить главного админа
    if (USERS[userIndex].username === 'stexiel') {
        return res.status(400).json({ error: 'Нельзя удалить главного администратора' });
    }

    USERS.splice(userIndex, 1);

    res.json({ message: 'Пользователь удален' });
});

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
    const fromDate = new Date(from);
    const toDate = new Date(to);
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
            beginTime: r.begin_time,
            finishTime: r.finish_time
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
        res.json({ rows: result.recordset });
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
        res.json({ rows: result.recordset });
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
        const result = await queryWMS(
            `SELECT loc_code, update_time
             FROM inv_container
             WHERE (loc_code LIKE '%D1' OR loc_code LIKE '%D2')
               AND update_time >= @from AND update_time <= @to`,
            { from: range.fromDate, to: range.toDate }
        );

        const pallets = {};
        for (const row of result.recordset) {
            const loc = row.loc_code;
            const type = loc.slice(-2);
            const id = loc.slice(0, -2);
            if (!pallets[id]) pallets[id] = { d1: false, d2: false };
            if (type === 'D1') pallets[id].d1 = true;
            if (type === 'D2') pallets[id].d2 = true;
        }

        const paired = [];
        const unpaired = [];
        for (const id in pallets) {
            const p = pallets[id];
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

// ===== Документы приемки (WMS_1386.in_receiving_bill) =====
app.get('/api/receipt-documents', authenticateToken, async (req, res) => {
    try {
        // Получаем последние 50 документов приемки
        const result = await queryWMS(
            `SELECT TOP 50
                    bill_no AS docNumber,
                    start_time AS startTime,
                    end_time AS endTime
             FROM in_receiving_bill
             WHERE is_deleted = 0
             ORDER BY start_time DESC`,
            {}
        );

        const docs = result.recordset.map(r => ({
            docNumber: r.docNumber,
            startTime: r.startTime ? r.startTime.toISOString().replace('T', ' ').slice(0, 19) : '',
            endTime: r.endTime ? r.endTime.toISOString().replace('T', ' ').slice(0, 19) : ''
        }));

        res.json(docs);
    } catch (err) {
        console.error('Ошибка /api/receipt-documents:', err);
        res.status(500).json({ error: err.message });
    }
});

// ===== Детали документа приемки (WMS_1386.in_receiving_bill + in_receiving_bill_detail + in_receiving_record) =====
app.get('/api/receipt-detail', authenticateToken, async (req, res) => {
    const { docNumber } = req.query;
    if (!docNumber) {
        res.status(400).json({ error: 'Требуется параметр docNumber' });
        return;
    }
    try {
        const result = await queryWMS(
            `SELECT r.container_code AS containerNumber,
                    r.qty AS quantity,
                    '' AS externalDocNumber,
                    m.code AS article,
                    l.lot_produce AS batchNumber,
                    s.putaway_time,
                    s.pick_time,
                    CASE
                        WHEN s.container_code IS NULL THEN 'Импорт'
                        WHEN s.pick_time IS NOT NULL THEN 'Выставлено'
                        WHEN s.putaway_time IS NOT NULL THEN 'Получен'
                        ELSE 'Импорт'
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

        const details = result.recordset.map(r => ({
            containerNumber: r.containerNumber,
            quantity: r.quantity,
            externalDocNumber: r.externalDocNumber,
            article: r.article,
            batchNumber: r.batchNumber,
            palletState: r.palletState
        }));

        res.json(details);
    } catch (err) {
        console.error('Ошибка /api/receipt-detail:', err);
        res.status(500).json({ error: err.message });
    }
});

// HTTP сервер (для разработки)
app.listen(PORT, '0.0.0.0', () => {
    console.log(`[WareCore Backend] HTTP сервер запущен на порту ${PORT}`);
    console.log(`[WareCore Backend] Откройте http://localhost:${PORT} для доступа к приложению`);
    console.log(`[WareCore Backend] Публичный доступ: http://92.38.29.140:${PORT}`);
});

// HTTPS сервер
https.createServer(sslOptions, app).listen(HTTPS_PORT, '0.0.0.0', () => {
    console.log(`[WareCore Backend] HTTPS сервер запущен на порту ${HTTPS_PORT}`);
    console.log(`[WareCore Backend] Откройте https://localhost:${HTTPS_PORT} для доступа к приложению`);
    console.log(`[WareCore Backend] Публичный доступ: https://92.38.29.140:${HTTPS_PORT}`);
});

process.on('uncaughtException', err => {
    console.error('[WareCore Backend] Необработанное исключение:', err);
});
process.on('unhandledRejection', err => {
    console.error('[WareCore Backend] Необработанный отказ промиса:', err);
});
