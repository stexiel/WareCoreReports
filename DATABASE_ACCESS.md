# Документация по доступу к базе данных SQL Server

## ⚠️ ВАЖНОЕ ПРЕДУПРЕЖДЕНИЕ

**ЭТОТ ДОСТУП ЯВЛЯЕТСЯ ТОЛЬКО ДЛЯ ЧТЕНИЯ (READ-ONLY)**

- ❌ **ЗАПРЕЩЕНО:** Изменять данные в базе данных
- ❌ **ЗАПРЕЩЕНО:** Удалять записи или таблицы
- ❌ **ЗАПРЕЩЕНО:** Исправлять или модифицировать структуру БД
- ❌ **ЗАПРЕЩЕНО:** Выполнять любые команды INSERT, UPDATE, DELETE, DROP, ALTER
- ✅ **РАЗРЕШЕНО:** Только чтение данных (SELECT команды)
- ✅ **РАЗРЕШЕНО:** Просмотр структуры таблиц и данных

**Любые попытки изменения данных могут привести к нарушению работы системы WMS/WCS!**

---

## Информация о подключении

### Сервер
- **IP адрес:** 172.30.1.10
- **Имя сервера:** WMSSERVER1
- **Тип БД:** Microsoft SQL Server
- **Порт:** 1433 (стандартный)

### Учетные данные для подключения
- **Имя пользователя:** `readonly_user`
- **Пароль:** `Read123!Pass@`
- **Тип доступа:** Read-Only (только чтение)

### Доступные базы данных

#### 1. WMS_1386
- **Описание:** Основная база данных WMS (Warehouse Management System)
- **Количество таблиц:** 105
- **Основные таблицы:**
  - `inv_container` - Контейнеры
  - `inv_lot` - Лоты
  - `inv_stock` - Складские остатки
  - `bm_material` - Материалы
  - `bm_customer` - Клиенты
  - `bm_supplier` - Поставщики
  - `wcs_job` - Задания WCS
  - `wcs_request` - Запросы WCS
  - `ur_user` - Пользователи
  - `ur_role` - Роли

#### 2. WCS_1386
- **Описание:** База данных WCS (Warehouse Control System)
- **Количество таблиц:** 73
- **Основные таблицы:**
  - `wcs_job` - Задания WCS
  - `wcs_order` - Заказы WCS
  - `wcs_plc` - PLC контроллеры
  - `wcs_route` - Маршруты
  - `eqpt_rgv` - RGV оборудование
  - `eqpt_conveyor` - Конвейеры
  - `pr_robot` - Роботы
  - `rgv_job` - Задания RGV

#### 3. WMS_Test
- **Описание:** Тестовая база данных WMS
- **Количество таблиц:** 105
- **Структура:** Идентична WMS_1386

---

## Способы подключения

### Способ 1: PowerShell (с компьютера)

```powershell
# Подключение к WMS_1386
Add-Type -AssemblyName System.Data
$conn = New-Object System.Data.SqlClient.SqlConnection("Server=172.30.1.10;Database=WMS_1386;User Id=readonly_user;Password=Read123!Pass@")
$conn.Open()

# Выполнение запроса
$cmd = $conn.CreateCommand()
$cmd.CommandText = "SELECT TOP 10 * FROM inv_stock"
$reader = $cmd.ExecuteReader()
while($reader.Read()) {
    Write-Host $reader[0]
}
$conn.Close()
```

### Способ 2: SQL Server Management Studio (SSMS)

1. Откройте SSMS
2. Нажмите "Connect" → "Database Engine"
3. Введите:
   - Server name: `172.30.1.10`
   - Authentication: `SQL Server Authentication`
   - Login: `readonly_user`
   - Password: `Read123!Pass@`
4. Нажмите "Connect"

### Способ 3: Командная строка (sqlcmd)

```cmd
sqlcmd -S 172.30.1.10 -U readonly_user -P Read123!Pass@ -d WMS_1386 -Q "SELECT TOP 10 * FROM inv_stock"
```

### Способ 4: Python

```python
import pyodbc

conn_str = (
    'DRIVER={ODBC Driver 17 for SQL Server};'
    'SERVER=172.30.1.10;'
    'DATABASE=WMS_1386;'
    'UID=readonly_user;'
    'PWD=Read123!Pass@'
)

conn = pyodbc.connect(conn_str)
cursor = conn.cursor()
cursor.execute("SELECT TOP 10 * FROM inv_stock")
for row in cursor:
    print(row)
conn.close()
```

---

## Примеры безопасных запросов (только чтение)

### Просмотр списка таблиц
```sql
SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'
```

### Просмотр данных таблицы
```sql
SELECT TOP 100 * FROM inv_stock
```

### Фильтрация данных
```sql
SELECT * FROM inv_lot WHERE lot_status = 'ACTIVE'
```

### Подсчет записей
```sql
SELECT COUNT(*) FROM inv_stock
```

---

## Ограничения прав доступа

Пользователь `readonly_user` имеет роль `db_datareader`, которая предоставляет:

✅ **Разрешено:**
- SELECT из всех таблиц
- VIEW DEFINITION (просмотр структуры)

❌ **Запрещено:**
- INSERT (вставка данных)
- UPDATE (обновление данных)
- DELETE (удаление данных)
- EXECUTE (выполнение хранимых процедур)
- DDL операции (CREATE, ALTER, DROP)

---

## Контактная информация

При возникновении проблем с доступом:
- Администратор БД: [указать контакт]
- Системный администратор: [указать контакт]

---

## Дата создания документа

27 июля 2026

---

## История изменений

| Дата | Изменение | Автор |
|------|-----------|-------|
| 27.07.2026 | Создание документа, настройка read-only доступа | System |
