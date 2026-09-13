require('dotenv').config({ path: __dirname + '/../telegram-bot/.env' });
const sql = require('mssql');

// Использует те же реквизиты БД, что и telegram-bot (telegram-bot/.env),
// но собственный пул соединений - не зависит от backend/server.js и telegram-bot/bot.js.
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
        console.error(`[WA Bot DB:${database}] Ошибка пула соединений:`, err.message);
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

async function queryWCS(query, params = {}) {
    const pool = await getPool('WCS_1386');
    const request = pool.request();
    Object.entries(params).forEach(([key, val]) => request.input(key, val));
    return request.query(query);
}

module.exports = { sql, queryWMS, queryWCS, getPool };
