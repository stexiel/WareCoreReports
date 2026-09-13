require('dotenv').config({ path: __dirname + '/.env' });
const sql = require('mssql');

// Собственное подключение к БД для Telegram-бота.
// Дублирует backend/db.js специально: бот не зависит от backend/server.js
// и продолжает работать, даже если backend недоступен.
const baseConfig = {
    server: process.env.DB_SERVER,
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    port: parseInt(process.env.DB_PORT || '1433', 10),
    options: {
        encrypt: false,
        trustServerCertificate: true
    },
    pool: {
        max: 5,
        min: 0,
        idleTimeoutMillis: 30000
    }
};

const pools = {};

async function getPool(database) {
    if (pools[database]) {
        try {
            if (pools[database].connected) return pools[database];
        } catch (e) { /* переподключаемся ниже */ }
    }
    const pool = new sql.ConnectionPool({ ...baseConfig, database });
    await pool.connect();
    pool.on('error', err => {
        console.error(`[Bot DB:${database}] Ошибка пула соединений:`, err.message);
    });
    pools[database] = pool;
    return pool;
}

async function queryWMS(query, params = {}) {
    const pool = await getPool('WMS_1386');
    const request = pool.request();
    Object.entries(params).forEach(([key, val]) => request.input(key, val));
    return request.query(query);
}

module.exports = { sql, queryWMS, getPool };
