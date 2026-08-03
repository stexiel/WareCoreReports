const { queryWMS } = require('./db');

async function probePalletStates() {
    try {
        console.log('=== Исследование статусов паллет ===\n');

        // Получаем пример данных из in_receiving_record
        const recordResult = await queryWMS(
            `SELECT TOP 10 
                    r.container_code,
                    r.qty,
                    r.status,
                    s.putaway_time,
                    s.pick_time,
                    s.loc_code
             FROM in_receiving_record r
             LEFT JOIN inv_stock s ON s.container_code = r.container_code
             WHERE r.status = 'Completed'
             ORDER BY r.id DESC`
        );

        console.log('Пример данных из in_receiving_record + inv_stock:');
        console.table(recordResult.recordset);

        // Получаем пример данных из inv_container
        const containerResult = await queryWMS(
            `SELECT TOP 10 
                    container_code,
                    loc_code,
                    update_time
             FROM inv_container
             ORDER BY update_time DESC`
        );

        console.log('\nПример данных из inv_container:');
        console.table(containerResult.recordset);

        // Проверяем связь между loc_code и статусами
        const locAnalysis = await queryWMS(
            `SELECT TOP 20
                    s.container_code,
                    s.loc_code,
                    s.putaway_time,
                    s.pick_time,
                    CASE
                        WHEN s.loc_code IS NULL OR s.loc_code = '' THEN 'NO_LOC'
                        WHEN s.loc_code LIKE '%IN%' OR s.loc_code LIKE '%RECEIVE%' THEN 'IN_LOC'
                        WHEN s.loc_code LIKE '%OUT%' OR s.loc_code LIKE '%SHIP%' THEN 'OUT_LOC'
                        ELSE 'OTHER_LOC'
                    END AS loc_type
             FROM inv_stock s
             WHERE s.container_code IS NOT NULL
             ORDER BY s.update_time DESC`
        );

        console.log('\nАнализ loc_code:');
        console.table(locAnalysis.recordset);

        process.exit(0);
    } catch (err) {
        console.error('Ошибка:', err);
        process.exit(1);
    }
}

probePalletStates();
