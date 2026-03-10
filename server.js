const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const readline = require('readline');
const Database = require('better-sqlite3');
const XLSX = require('xlsx');
const multer = require('multer');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

// === НАСТРОЙКИ ===
const CONFIG = {
  PORT: process.env.PORT || 3000,
  DATA_FOLDER: process.env.DATA_FOLDER || './data',
  DB_PATH: process.env.DB_PATH || './database.sqlite',
  API_KEY: process.env.ANTHROPIC_API_KEY || '',
  ADMIN_PASSWORD: process.env.ADMIN_PASSWORD || 'admin123',
  SAMPLE_SIZE_FOR_AI: 50,
};

// === MULTER (загрузка файлов) ===
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const folder = path.resolve(CONFIG.DATA_FOLDER);
    if (!fs.existsSync(folder)) fs.mkdirSync(folder, { recursive: true });
    cb(null, folder);
  },
  filename: (req, file, cb) => {
    // Сохраняем оригинальное имя файла
    const originalName = Buffer.from(file.originalname, 'latin1').toString('utf8');
    cb(null, originalName);
  },
});
const upload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    if (/\.(csv|xlsx|xls)$/i.test(file.originalname)) {
      cb(null, true);
    } else {
      cb(new Error('Только CSV и Excel файлы'));
    }
  },
  limits: { fileSize: 200 * 1024 * 1024 }, // 200 МБ
});

// === ИНИЦИАЛИЗАЦИЯ SQLITE ===
console.log('📦 Инициализация SQLite базы данных...');
const db = new Database(CONFIG.DB_PATH);
db.pragma('journal_mode = WAL'); // Быстрее для записи

// Создаём таблицу для метаданных файлов
db.exec(`
  CREATE TABLE IF NOT EXISTS files_meta (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    filename TEXT UNIQUE,
    columns TEXT,
    row_count INTEGER,
    loaded_at DATETIME DEFAULT CURRENT_TIMESTAMP
  )
`);

console.log('✅ SQLite готов\n');

// === ФУНКЦИИ РАБОТЫ С БАЗОЙ ДАННЫХ ===

// Создать таблицу для данных файла
function createDataTable(tableName, columns) {
  const safeTableName = tableName.replace(/[^a-zA-Z0-9_]/g, '_');
  
  // Удаляем старую таблицу если есть
  db.exec(`DROP TABLE IF EXISTS "${safeTableName}"`);
  
  // Создаём колонки (все как TEXT для простоты)
  const columnDefs = columns.map((col, idx) => {
    const safeCol = `col_${idx}`;
    return `"${safeCol}" TEXT`;
  }).join(', ');
  
  db.exec(`CREATE TABLE "${safeTableName}" (id INTEGER PRIMARY KEY AUTOINCREMENT, ${columnDefs})`);
  
  // Создаём индексы для быстрого поиска
  // Индекс на первые 3 колонки (обычно там БИН, дата и т.д.)
  for (let i = 0; i < Math.min(3, columns.length); i++) {
    db.exec(`CREATE INDEX IF NOT EXISTS "idx_${safeTableName}_col_${i}" ON "${safeTableName}"("col_${i}")`);
  }
  
  return safeTableName;
}

// Загрузить CSV в SQLite
async function loadCsvToSqlite(filePath) {
  const fileName = path.basename(filePath);
  const tableName = fileName.replace(/[^a-zA-Z0-9]/g, '_');
  
  console.log(`\n📖 Загрузка в SQLite: ${fileName}`);
  const startTime = Date.now();
  
  const fileStats = fs.statSync(filePath);
  const fileSizeMB = (fileStats.size / 1024 / 1024).toFixed(1);
  console.log(`   Размер файла: ${fileSizeMB} МБ`);
  
  return new Promise((resolve) => {
    let columns = [];
    let rowCount = 0;
    let headerParsed = false;
    let delimiter = null;
    let safeTableName = null;
    let insertStmt = null;
    let batch = [];
    const BATCH_SIZE = 5000;
    
    const rl = readline.createInterface({
      input: fs.createReadStream(filePath, { encoding: 'utf8' }),
      crlfDelay: Infinity,
    });
    
    rl.on('line', (line) => {
      if (!line.trim()) return;
      
      // Определяем разделитель
      if (!delimiter) {
        const semicolons = (line.match(/;/g) || []).length;
        const commas = (line.match(/,/g) || []).length;
        delimiter = semicolons > commas ? ';' : ',';
        console.log(`   Разделитель: "${delimiter}"`);
      }
      
      const values = parseCSVLine(line, delimiter);
      
      // Ищем заголовки
      if (!headerParsed) {
        const nonEmptyCount = values.filter(v => v && v.trim()).length;
        if (nonEmptyCount >= 3 && values.some(v => v && /[а-яА-Яa-zA-Z]{2,}/.test(v))) {
          columns = values.map((v, idx) => {
            const val = String(v || '').trim();
            return val || `Колонка_${idx + 1}`;
          });
          
          // Убираем пустые колонки в конце
          while (columns.length > 0 && !columns[columns.length - 1].trim()) {
            columns.pop();
          }
          
          console.log(`   Колонок: ${columns.length}`);
          console.log(`   Первые: ${columns.slice(0, 5).join(', ')}...`);
          
          // Создаём таблицу
          safeTableName = createDataTable(tableName, columns);
          
          // Подготавливаем INSERT
          const placeholders = columns.map(() => '?').join(', ');
          insertStmt = db.prepare(`INSERT INTO "${safeTableName}" (${columns.map((_, i) => `"col_${i}"`).join(', ')}) VALUES (${placeholders})`);
          
          headerParsed = true;
          return;
        }
        return;
      }
      
      rowCount++;
      
      // Добавляем в batch
      const rowValues = columns.map((_, idx) => values[idx] || '');
      batch.push(rowValues);
      
      // Вставляем batch
      if (batch.length >= BATCH_SIZE) {
        const insertMany = db.transaction((rows) => {
          for (const row of rows) {
            insertStmt.run(...row);
          }
        });
        insertMany(batch);
        batch = [];
        
        process.stdout.write(`\r   Загружено: ${rowCount.toLocaleString()} строк...`);
      }
    });
    
    rl.on('close', () => {
      // Вставляем остаток
      if (batch.length > 0 && insertStmt) {
        const insertMany = db.transaction((rows) => {
          for (const row of rows) {
            insertStmt.run(...row);
          }
        });
        insertMany(batch);
      }
      
      console.log(`\r   Загружено: ${rowCount.toLocaleString()} строк    `);
      
      // Сохраняем метаданные
      const metaStmt = db.prepare(`
        INSERT OR REPLACE INTO files_meta (filename, columns, row_count, loaded_at)
        VALUES (?, ?, ?, datetime('now'))
      `);
      metaStmt.run(fileName, JSON.stringify(columns), rowCount);
      
      const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
      console.log(`   ✅ Готово за ${elapsed} сек`);
      
      resolve({
        filename: fileName,
        tableName: safeTableName,
        columns,
        rowCount,
      });
    });
    
    rl.on('error', (err) => {
      console.error(`   ❌ Ошибка: ${err.message}`);
      resolve(null);
    });
  });
}

// Парсер CSV строки
function parseCSVLine(line, delimiter = ';') {
  const result = [];
  let current = '';
  let inQuotes = false;
  
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    
    if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === delimiter && !inQuotes) {
      result.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }
  
  result.push(current.trim());
  return result;
}

// Загрузить Excel файл в SQLite
function loadExcelToSqlite(filePath) {
  const fileName = path.basename(filePath);
  const tableName = fileName.replace(/[^a-zA-Z0-9]/g, '_');
  
  console.log(`\n📖 Загрузка Excel в SQLite: ${fileName}`);
  const startTime = Date.now();
  
  const fileStats = fs.statSync(filePath);
  const fileSizeMB = (fileStats.size / 1024 / 1024).toFixed(1);
  console.log(`   Размер файла: ${fileSizeMB} МБ`);
  
  try {
    const workbook = XLSX.readFile(filePath, {
      cellDates: true,
      cellNF: false,
      raw: true,
    });
    
    let totalRows = 0;
    let allColumns = [];
    
    workbook.SheetNames.forEach((sheetName, sheetIdx) => {
      console.log(`   📄 Лист ${sheetIdx + 1}: "${sheetName}"`);
      
      const sheet = workbook.Sheets[sheetName];
      if (!sheet['!ref']) {
        console.log(`      ⚠️ Пустой лист`);
        return;
      }
      
      const jsonData = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: '',
        blankrows: false,
      });
      
      if (jsonData.length < 2) {
        console.log(`      ⚠️ Нет данных`);
        return;
      }
      
      // Находим заголовки
      let headerRowIndex = 0;
      for (let i = 0; i < Math.min(10, jsonData.length); i++) {
        const row = jsonData[i];
        if (!row) continue;
        const nonEmpty = row.filter(v => v !== undefined && v !== '').length;
        if (nonEmpty >= 2) {
          headerRowIndex = i;
          break;
        }
      }
      
      const columns = (jsonData[headerRowIndex] || []).map((v, idx) => {
        return String(v || '').trim() || `Колонка_${idx + 1}`;
      });
      
      console.log(`      Колонок: ${columns.length}`);
      
      // Создаём таблицу для этого листа
      const sheetTableName = `${tableName}_${sheetName.replace(/[^a-zA-Z0-9]/g, '_')}`;
      const safeTableName = createDataTable(sheetTableName, columns);
      
      // Подготавливаем INSERT
      const placeholders = columns.map(() => '?').join(', ');
      const insertStmt = db.prepare(`INSERT INTO "${safeTableName}" (${columns.map((_, i) => `"col_${i}"`).join(', ')}) VALUES (${placeholders})`);
      
      // Вставляем данные батчами
      const dataRows = jsonData.slice(headerRowIndex + 1);
      const BATCH_SIZE = 5000;
      let batch = [];
      
      const insertMany = db.transaction((rows) => {
        for (const row of rows) {
          insertStmt.run(...row);
        }
      });
      
      dataRows.forEach((row, idx) => {
        const rowValues = columns.map((_, i) => {
          let val = row[i];
          if (val instanceof Date) {
            val = val.toLocaleDateString('ru-RU');
          }
          return val !== undefined ? String(val) : '';
        });
        batch.push(rowValues);
        
        if (batch.length >= BATCH_SIZE) {
          insertMany(batch);
          batch = [];
          process.stdout.write(`\r      Загружено: ${(idx + 1).toLocaleString()} строк...`);
        }
      });
      
      if (batch.length > 0) {
        insertMany(batch);
      }
      
      console.log(`\r      ✅ ${dataRows.length.toLocaleString()} строк    `);
      
      totalRows += dataRows.length;
      allColumns = columns;
      
      // Сохраняем метаданные
      const metaStmt = db.prepare(`
        INSERT OR REPLACE INTO files_meta (filename, columns, row_count, loaded_at)
        VALUES (?, ?, ?, datetime('now'))
      `);
      metaStmt.run(`${fileName}:${sheetName}`, JSON.stringify(columns), dataRows.length);
    });
    
    const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
    console.log(`   ✅ Готово за ${elapsed} сек (${totalRows.toLocaleString()} строк)`);
    
    return { filename: fileName, columns: allColumns, rowCount: totalRows };
    
  } catch (error) {
    console.error(`   ❌ Ошибка: ${error.message}`);
    return null;
  }
}

// Сканировать папку и загрузить файлы
async function scanAndLoadFiles() {
  const folderPath = path.resolve(CONFIG.DATA_FOLDER);
  
  if (!fs.existsSync(folderPath)) {
    fs.mkdirSync(folderPath, { recursive: true });
    console.log(`📁 Создана папка: ${folderPath}`);
  }

  const files = fs.readdirSync(folderPath).filter(f => 
    /\.(csv|xlsx|xls)$/i.test(f)
  );

  console.log(`\n🔍 Сканирование папки: ${folderPath}`);
  console.log(`📄 Найдено файлов: ${files.length}\n`);

  for (const filename of files) {
    const filePath = path.join(folderPath, filename);
    const ext = path.extname(filename).toLowerCase();
    
    // Проверяем, загружен ли уже файл
    const existing = db.prepare('SELECT * FROM files_meta WHERE filename LIKE ?').get(`${filename}%`);
    
    if (existing) {
      console.log(`⏭️ ${filename} уже загружен (${existing.row_count.toLocaleString()} строк)`);
      continue;
    }
    
    // Загружаем в зависимости от типа
    if (ext === '.csv') {
      await loadCsvToSqlite(filePath);
    } else if (ext === '.xlsx' || ext === '.xls') {
      loadExcelToSqlite(filePath);
    }
  }

  // Показываем статистику
  const allFiles = db.prepare('SELECT * FROM files_meta').all();
  const totalRows = allFiles.reduce((sum, f) => sum + f.row_count, 0);
  
  console.log(`\n════════════════════════════════════════`);
  console.log(`📊 Всего в базе: ${totalRows.toLocaleString()} строк`);
  console.log(`📁 Файлов: ${allFiles.length}`);
  console.log(`════════════════════════════════════════\n`);
}

// === API ENDPOINTS ===

// Статус базы данных
app.get('/api/status', (req, res) => {
  const files = db.prepare('SELECT * FROM files_meta').all();
  const totalRows = files.reduce((sum, f) => sum + f.row_count, 0);
  
  res.json({
    status: 'ok',
    totalFiles: files.length,
    totalRows,
    files: files.map(f => ({
      name: f.filename,
      sheets: ['Sheet1'],
      totalRows: f.row_count,
      columns: JSON.parse(f.columns),
    })),
  });
});

// Получить уникальные значения для фильтров
app.get('/api/filters', (req, res) => {
  console.log('\n📋 Загрузка фильтров...');
  
  const files = db.prepare('SELECT * FROM files_meta').all();
  const filters = {
    companies: new Set(),
    bins: new Set(),
    dates: new Set(),
  };
  
  for (const fileMeta of files) {
    const tableName = fileMeta.filename.replace(/[^a-zA-Z0-9]/g, '_');
    const columns = JSON.parse(fileMeta.columns);
    
    // Ищем индексы нужных колонок
    let companyColIdx = -1;
    let binColIdx = -1;
    let dateColIdx = -1;
    
    columns.forEach((col, idx) => {
      const colLower = col.toLowerCase();
      // Ищем "Наименование субъекта" или похожее
      if (colLower === 'наименование субъекта') {
        companyColIdx = idx;
      }
      if (colLower.includes('наименование') && colLower.includes('субъект') && companyColIdx === -1) {
        companyColIdx = idx;
      }
      // Дата или День
      if (col === 'Дата' || col === 'День' || colLower === 'дата' || colLower === 'день') {
        dateColIdx = idx;
      }
    });
    
    try {
      // Получаем уникальные компании
      if (companyColIdx >= 0) {
        const companies = db.prepare(`
          SELECT DISTINCT "col_${companyColIdx}" as val 
          FROM "${tableName}" 
          WHERE "col_${companyColIdx}" IS NOT NULL AND "col_${companyColIdx}" != ''
          LIMIT 500
        `).all();
        companies.forEach(r => filters.companies.add(r.val));
      }
      
      // Получаем уникальные даты
      if (dateColIdx >= 0) {
        const dates = db.prepare(`
          SELECT DISTINCT "col_${dateColIdx}" as val 
          FROM "${tableName}" 
          WHERE "col_${dateColIdx}" IS NOT NULL AND "col_${dateColIdx}" != ''
          ORDER BY "col_${dateColIdx}"
          LIMIT 100
        `).all();
        dates.forEach(r => filters.dates.add(r.val));
      }
    } catch (e) {
      console.error(`   Ошибка:`, e.message);
    }
  }
  
  const result = {
    companies: [...filters.companies].sort(),
    dates: [...filters.dates].sort(),
  };
  
  console.log(`   Компаний: ${result.companies.length}, Дат: ${result.dates.length}`);
  
  res.json(result);
});

// Поиск с фильтрами
app.get('/api/search-filtered', (req, res) => {
  const { company, dateFrom, dateTo, limit = 1000 } = req.query;
  
  console.log(`\n🔍 Поиск с фильтрами: компания=${company || 'все'}, период=${dateFrom || '?'}-${dateTo || '?'}`);
  
  const files = db.prepare('SELECT * FROM files_meta').all();
  const results = [];
  const aggregatedStats = {};
  
  for (const fileMeta of files) {
    const tableName = fileMeta.filename.replace(/[^a-zA-Z0-9]/g, '_');
    const columns = JSON.parse(fileMeta.columns);
    
    // Ищем индексы колонок
    let companyColIdx = -1;
    let dateColIdx = -1;
    
    columns.forEach((col, idx) => {
      const colLower = col.toLowerCase();
      // Ищем "Наименование субъекта"
      if (colLower === 'наименование субъекта') {
        companyColIdx = idx;
      }
      if (colLower.includes('наименование') && colLower.includes('субъект') && companyColIdx === -1) {
        companyColIdx = idx;
      }
      // Дата или День
      if (col === 'Дата' || col === 'День' || colLower === 'дата' || colLower === 'день') {
        dateColIdx = idx;
      }
    });
    
    // Строим WHERE
    const whereParts = [];
    const params = [];
    
    if (company && companyColIdx >= 0) {
      whereParts.push(`"col_${companyColIdx}" LIKE ?`);
      params.push(`%${company}%`);
    }
    
    // Конвертируем дату из 2026-01-01 в 01.01.2026
    const convertDate = (isoDate) => {
      if (!isoDate) return null;
      const parts = isoDate.split('-');
      if (parts.length === 3) {
        return `${parts[2]}.${parts[1]}.${parts[0]}`;
      }
      return isoDate;
    };
    
    const dateFromFormatted = convertDate(dateFrom);
    const dateToFormatted = convertDate(dateTo);
    
    if (dateFromFormatted && dateColIdx >= 0) {
      whereParts.push(`"col_${dateColIdx}" >= ?`);
      params.push(dateFromFormatted);
    }
    
    if (dateToFormatted && dateColIdx >= 0) {
      whereParts.push(`"col_${dateColIdx}" <= ?`);
      params.push(dateToFormatted);
    }
    
    if (whereParts.length === 0) {
      continue;
    }
    
    try {
      const sql = `SELECT * FROM "${tableName}" WHERE ${whereParts.join(' AND ')} LIMIT ${Number(limit)}`;
      const rows = db.prepare(sql).all(...params);
      
      console.log(`   Найдено в ${fileMeta.filename}: ${rows.length}`);
      
      rows.forEach(row => {
        // Агрегируем числа
        columns.forEach((col, i) => {
          const val = row[`col_${i}`];
          if (!val) return;
          
          const cleanVal = String(val).replace(/\s/g, '').replace(',', '.');
          const num = parseFloat(cleanVal);
          
          if (!isNaN(num)) {
            if (!aggregatedStats[col]) {
              aggregatedStats[col] = { sum: 0, count: 0 };
            }
            aggregatedStats[col].sum += num;
            aggregatedStats[col].count++;
          }
        });
        
        // Сохраняем первые 100 строк
        if (results.length < 100) {
          const data = {};
          columns.forEach((col, i) => {
            data[col] = row[`col_${i}`];
          });
          results.push(data);
        }
      });
      
    } catch (e) {
      console.error(`   Ошибка:`, e.message);
    }
  }
  
  res.json({
    found: results.length,
    results,
    stats: aggregatedStats,
  });
});

// Поиск в базе данных
app.get('/api/search', (req, res) => {
  const { q, file, limit = 100 } = req.query;
  
  if (!q) {
    return res.status(400).json({ error: 'Параметр q обязателен' });
  }

  console.log(`\n🔍 Поиск: "${q}"`);
  
  const files = db.prepare('SELECT * FROM files_meta').all();
  const results = [];
  
  for (const fileMeta of files) {
    if (file && fileMeta.filename !== file) continue;
    
    const tableName = fileMeta.filename.replace(/[^a-zA-Z0-9]/g, '_');
    const columns = JSON.parse(fileMeta.columns);
    
    // Ищем по всем колонкам
    const whereClauses = columns.map((_, i) => `"col_${i}" LIKE ?`).join(' OR ');
    const searchPattern = `%${q}%`;
    const params = columns.map(() => searchPattern);
    
    try {
      const rows = db.prepare(`
        SELECT * FROM "${tableName}" 
        WHERE ${whereClauses} 
        LIMIT ${Number(limit)}
      `).all(...params);
      
      rows.forEach(row => {
        const data = {};
        columns.forEach((col, i) => {
          data[col] = row[`col_${i}`];
        });
        results.push({
          file: fileMeta.filename,
          row: row.id,
          data,
        });
      });
    } catch (e) {
      console.error(`   Ошибка поиска в ${tableName}:`, e.message);
    }
  }

  console.log(`   Найдено: ${results.length}`);
  
  res.json({
    query: q,
    found: results.length,
    results: results.slice(0, Number(limit)),
  });
});

// Агрегация данных
app.get('/api/aggregate', (req, res) => {
  const { file, column, operation = 'sum', groupBy } = req.query;

  if (!file || !column) {
    return res.status(400).json({ error: 'Параметры file и column обязательны' });
  }

  const fileMeta = db.prepare('SELECT * FROM files_meta WHERE filename = ?').get(file);
  if (!fileMeta) {
    return res.status(404).json({ error: 'Файл не найден' });
  }

  const tableName = file.replace(/[^a-zA-Z0-9]/g, '_');
  const columns = JSON.parse(fileMeta.columns);
  const colIndex = columns.indexOf(column);
  
  if (colIndex === -1) {
    return res.status(400).json({ error: `Колонка "${column}" не найдена` });
  }

  const colName = `col_${colIndex}`;
  
  let sql;
  if (groupBy) {
    const groupColIndex = columns.indexOf(groupBy);
    if (groupColIndex === -1) {
      return res.status(400).json({ error: `Колонка группировки "${groupBy}" не найдена` });
    }
    const groupColName = `col_${groupColIndex}`;
    sql = `SELECT "${groupColName}" as group_key, ${operation}(CAST("${colName}" AS REAL)) as result, COUNT(*) as count FROM "${tableName}" GROUP BY "${groupColName}" ORDER BY result DESC LIMIT 100`;
  } else {
    sql = `SELECT ${operation}(CAST("${colName}" AS REAL)) as result, COUNT(*) as count FROM "${tableName}"`;
  }

  try {
    const result = db.prepare(sql).all();
    res.json({ column, operation, groupBy, result });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Запрос к AI
app.post('/api/ask', async (req, res) => {
  const { question, file, apiKey, filters } = req.body;

  if (!question) {
    return res.status(400).json({ error: 'Вопрос обязателен' });
  }

  const key = apiKey || CONFIG.API_KEY;
  
  if (!key) {
    return res.status(500).json({ error: 'API ключ не настроен. Нажмите 🔑 чтобы добавить ключ.' });
  }

  console.log(`\n💬 Вопрос: ${question}`);
  if (filters) {
    console.log(`   Фильтры: компания="${filters.company || 'все'}", период=${filters.dateFrom || '?'}-${filters.dateTo || '?'}`);
  }

  let searchResults = [];
  let aggregatedStats = {};
  
  // Если есть фильтры — ищем по ним
  const company = filters?.company;
  const dateFrom = filters?.dateFrom;
  const dateTo = filters?.dateTo;
  
  if (company || dateFrom || dateTo) {
    console.log(`🔍 Поиск по фильтрам...`);
    
    const files = db.prepare('SELECT * FROM files_meta').all();
    
    for (const fileMeta of files) {
      const tableName = fileMeta.filename.replace(/[^a-zA-Z0-9]/g, '_');
      const columns = JSON.parse(fileMeta.columns);
      
      // Ищем индексы колонок
      let companyColIdx = -1;
      let dateColIdx = -1;
      
      columns.forEach((col, idx) => {
        const colLower = col.toLowerCase();
        if (colLower === 'наименование субъекта') {
          companyColIdx = idx;
        }
        if (colLower.includes('наименование') && colLower.includes('субъект') && companyColIdx === -1) {
          companyColIdx = idx;
        }
        if (col === 'День' || col === 'Дата' || colLower === 'день' || colLower === 'дата') {
          dateColIdx = idx;
        }
      });
      
      // Строим WHERE
      const whereParts = [];
      const params = [];
      
      if (company && companyColIdx >= 0) {
        whereParts.push(`"col_${companyColIdx}" LIKE ?`);
        params.push(`%${company}%`);
      }
      
      // Конвертируем дату из 2026-01-01 в 01.01.2026
      const convertDate = (isoDate) => {
        if (!isoDate) return null;
        const parts = isoDate.split('-');
        if (parts.length === 3) {
          return `${parts[2]}.${parts[1]}.${parts[0]}`;
        }
        return isoDate;
      };
      
      const dateFromFormatted = convertDate(dateFrom);
      const dateToFormatted = convertDate(dateTo);
      
      if (dateFromFormatted && dateColIdx >= 0) {
        whereParts.push(`"col_${dateColIdx}" >= ?`);
        params.push(dateFromFormatted);
      }
      
      if (dateToFormatted && dateColIdx >= 0) {
        whereParts.push(`"col_${dateColIdx}" <= ?`);
        params.push(dateToFormatted);
      }
      
      if (whereParts.length === 0) continue;
      
      try {
        const sql = `SELECT * FROM "${tableName}" WHERE ${whereParts.join(' AND ')} LIMIT 5000`;
        const rows = db.prepare(sql).all(...params);
        
        console.log(`   Найдено: ${rows.length} записей`);
        
        // Агрегируем данные
        rows.forEach(row => {
          columns.forEach((col, i) => {
            const val = row[`col_${i}`];
            if (!val) return;
            
            const cleanVal = String(val).replace(/\s/g, '').replace(',', '.');
            const num = parseFloat(cleanVal);
            
            if (!isNaN(num)) {
              if (!aggregatedStats[col]) {
                aggregatedStats[col] = { sum: 0, count: 0 };
              }
              aggregatedStats[col].sum += num;
              aggregatedStats[col].count++;
            }
          });
          
          // Сохраняем первые 50 для примера
          if (searchResults.length < 50) {
            const data = {};
            columns.forEach((col, i) => {
              data[col] = row[`col_${i}`];
            });
            searchResults.push(data);
          }
        });
        
      } catch (e) {
        console.error(`   Ошибка:`, e.message);
      }
    }
  }

  // Формируем контекст
  let dataContext = '';
  
  if (Object.keys(aggregatedStats).length > 0) {
    dataContext = `🔍 РЕЗУЛЬТАТЫ ПОИСКА: Найдено записей

📊 АГРЕГИРОВАННАЯ СТАТИСТИКА:
${Object.entries(aggregatedStats).map(([col, stats]) => 
  `• ${col}: сумма=${stats.sum.toLocaleString()}, записей=${stats.count}`
).join('\n')}

📝 ПРИМЕРЫ ДАННЫХ (${searchResults.length} записей):`;
  } else {
    const files = db.prepare('SELECT * FROM files_meta').all();
    const totalRows = files.reduce((sum, f) => sum + f.row_count, 0);
    
    dataContext = `БАЗА ДАННЫХ:
${files.map(f => `• ${f.filename}: ${f.row_count.toLocaleString()} строк`).join('\n')}
Всего: ${totalRows.toLocaleString()} строк

⚠️ Фильтр не выбран. Попросите пользователя выбрать компанию в фильтре справа.`;
  }

  // Добавляем информацию о выбранных фильтрах
  let filterInfo = '';
  if (filters?.company) {
    filterInfo = `\n\n🏢 ВЫБРАННАЯ КОМПАНИЯ: ${filters.company}`;
  }

  try {
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': key,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 1500,
        system: `Ты — ИИ-аналитик энергетической компании. Тебя зовут AI.

ГЛАВНОЕ ПРАВИЛО: НЕ ЗАДАВАЙ ВОПРОСОВ! Сразу давай ответ.

ФОРМАТИРОВАНИЕ:
- Используй смайлики по теме: ⚡ для энергии, 📊 для статистики, 🏭 для компаний, 📅 для периода, ✅ для успеха
- НЕ используй символы # и * для заголовков
- НЕ используй markdown таблицы с |---|
- Используй HTML теги: <strong>, <br>, <table>, <tr>, <td>
- Формат чисел: 1 000 000 (с пробелами)
- Отвечай кратко и структурированно

ПРИМЕР ХОРОШЕГО ОТВЕТА:
⚡ <strong>Факт генерации</strong><br>
🏭 Компания: Экибастузская ГРЭС-1<br>
📅 Период: 01.01.2026 – 05.01.2026<br><br>
📊 <strong>Результаты:</strong><br>
• Фактическая генерация: <strong>411 889 834 кВт·ч</strong><br>
• Записей: 120<br>
• Среднее в час: 3 432 415 кВт·ч

ТЕРМИНЫ:
- Факт (Генерация) = фактическая выработка электроэнергии
- Факт (Потребление) = фактическое потребление
- План (Продажа) = плановая генерация
- План (Покупка) = плановое потребление
- Дисбаланс = разница между планом и фактом
${filterInfo}

ДАННЫЕ:
${dataContext}

${searchResults.length > 0 ? JSON.stringify(searchResults.slice(0, 30), null, 2) : ''}`,
        messages: [{ role: 'user', content: question }],
      }),
    });

    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.error?.message || 'API Error');
    }

    const data = await response.json();
    const answer = data.content?.[0]?.text || 'Нет ответа';

    res.json({ answer });
  } catch (error) {
    console.error('AI Error:', error);
    res.status(500).json({ error: error.message });
  }
});

// Перезагрузить файл
app.post('/api/reload', async (req, res) => {
  const { file } = req.body;
  
  if (file) {
    // Удаляем из базы
    db.prepare('DELETE FROM files_meta WHERE filename = ?').run(file);
    const tableName = file.replace(/[^a-zA-Z0-9]/g, '_');
    db.exec(`DROP TABLE IF EXISTS "${tableName}"`);
  } else {
    // Очищаем всё
    db.prepare('DELETE FROM files_meta').run();
    const tables = db.prepare("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'files_meta'").all();
    tables.forEach(t => {
      db.exec(`DROP TABLE IF EXISTS "${t.name}"`);
    });
  }
  
  await scanAndLoadFiles();
  
  res.json({ status: 'ok' });
});

// Пересканировать папку
app.post('/api/rescan', async (req, res) => {
  await scanAndLoadFiles();
  
  const files = db.prepare('SELECT * FROM files_meta').all();
  const totalRows = files.reduce((sum, f) => sum + f.row_count, 0);
  
  res.json({
    status: 'ok',
    totalFiles: files.length,
    totalRows,
  });
});

// Чаты (без изменений)
const HISTORY_FOLDER = path.join(__dirname, 'chats');
if (!fs.existsSync(HISTORY_FOLDER)) {
  fs.mkdirSync(HISTORY_FOLDER, { recursive: true });
}

app.get('/api/chats', (req, res) => {
  try {
    const files = fs.readdirSync(HISTORY_FOLDER)
      .filter(f => f.endsWith('.json'))
      .map(f => {
        const filePath = path.join(HISTORY_FOLDER, f);
        const data = JSON.parse(fs.readFileSync(filePath, 'utf8'));
        const stats = fs.statSync(filePath);
        return {
          id: f.replace('.json', ''),
          title: data.title || 'Без названия',
          messageCount: data.messages?.length || 0,
          updatedAt: stats.mtime,
        };
      })
      .sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt));
    res.json({ chats: files });
  } catch (error) {
    res.json({ chats: [] });
  }
});

app.post('/api/chats', (req, res) => {
  const id = 'chat_' + Date.now();
  const filePath = path.join(HISTORY_FOLDER, `${id}.json`);
  const chat = { id, title: 'Новый чат', messages: [], createdAt: new Date().toISOString() };
  fs.writeFileSync(filePath, JSON.stringify(chat, null, 2), 'utf8');
  res.json(chat);
});

app.get('/api/chats/:id', (req, res) => {
  const filePath = path.join(HISTORY_FOLDER, `${req.params.id}.json`);
  try {
    if (fs.existsSync(filePath)) {
      res.json(JSON.parse(fs.readFileSync(filePath, 'utf8')));
    } else {
      res.status(404).json({ error: 'Чат не найден' });
    }
  } catch (error) {
    res.status(500).json({ error: 'Ошибка чтения' });
  }
});

app.post('/api/chats/:id/messages', (req, res) => {
  const { message } = req.body;
  const filePath = path.join(HISTORY_FOLDER, `${req.params.id}.json`);
  try {
    let chat = { id: req.params.id, title: 'Новый чат', messages: [] };
    if (fs.existsSync(filePath)) {
      chat = JSON.parse(fs.readFileSync(filePath, 'utf8'));
    }
    chat.messages.push({ ...message, timestamp: new Date().toISOString() });
    if (chat.title === 'Новый чат' && message.role === 'user') {
      chat.title = message.content.slice(0, 40) + (message.content.length > 40 ? '...' : '');
    }
    fs.writeFileSync(filePath, JSON.stringify(chat, null, 2), 'utf8');
    res.json({ status: 'ok', title: chat.title });
  } catch (error) {
    res.status(500).json({ error: 'Не удалось сохранить' });
  }
});

app.delete('/api/chats/:id', (req, res) => {
  const filePath = path.join(HISTORY_FOLDER, `${req.params.id}.json`);
  try {
    if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
    res.json({ status: 'ok' });
  } catch (error) {
    res.status(500).json({ error: 'Не удалось удалить' });
  }
});

// Проверка пароля администратора
app.post('/api/admin/login', (req, res) => {
  const { password } = req.body;
  if (password === CONFIG.ADMIN_PASSWORD) {
    res.json({ ok: true });
  } else {
    res.status(401).json({ error: 'Неверный пароль' });
  }
});

// Загрузка файла (только для администратора)
app.post('/api/admin/upload', (req, res) => {
  const adminPassword = req.headers['x-admin-password'];
  if (adminPassword !== CONFIG.ADMIN_PASSWORD) {
    return res.status(401).json({ error: 'Нет доступа' });
  }

  upload.single('file')(req, res, async (err) => {
    if (err) {
      return res.status(400).json({ error: err.message });
    }
    if (!req.file) {
      return res.status(400).json({ error: 'Файл не выбран' });
    }

    console.log(`\n📤 Загружен файл: ${req.file.filename}`);

    // Удаляем старую запись если есть
    db.prepare('DELETE FROM files_meta WHERE filename LIKE ?').run(`${req.file.filename}%`);
    const oldTable = req.file.filename.replace(/[^a-zA-Z0-9]/g, '_');
    try { db.exec(`DROP TABLE IF EXISTS "${oldTable}"`); } catch(e) {}

    // Загружаем в SQLite
    const filePath = req.file.path;
    const ext = path.extname(req.file.filename).toLowerCase();
    let result;
    if (ext === '.csv') {
      result = await loadCsvToSqlite(filePath);
    } else {
      result = loadExcelToSqlite(filePath);
    }

    if (result) {
      res.json({ ok: true, filename: req.file.filename, rows: result.rowCount || result.totalRows });
    } else {
      res.status(500).json({ error: 'Ошибка загрузки файла' });
    }
  });
});

// Удаление файла (только для администратора)
app.delete('/api/admin/files/:filename', (req, res) => {
  const adminPassword = req.headers['x-admin-password'];
  if (adminPassword !== CONFIG.ADMIN_PASSWORD) {
    return res.status(401).json({ error: 'Нет доступа' });
  }

  const filename = decodeURIComponent(req.params.filename);
  db.prepare('DELETE FROM files_meta WHERE filename LIKE ?').run(`${filename}%`);
  const tableName = filename.replace(/[^a-zA-Z0-9]/g, '_');
  try { db.exec(`DROP TABLE IF EXISTS "${tableName}"`); } catch(e) {}

  const filePath = path.join(path.resolve(CONFIG.DATA_FOLDER), filename);
  if (fs.existsSync(filePath)) fs.unlinkSync(filePath);

  console.log(`🗑️ Удалён файл: ${filename}`);
  res.json({ ok: true });
});

// Статический файл
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// === ЗАПУСК ===
async function startServer() {
  console.log(`
╔══════════════════════════════════════════════════════════════╗
║           📊 DATA AGENT SERVER v3.0 (SQLite)                 ║
║                                                              ║
║  Быстрый поиск по миллионам строк!                          ║
╚══════════════════════════════════════════════════════════════╝
`);

  await scanAndLoadFiles();

  app.listen(CONFIG.PORT, () => {
    console.log(`🚀 Сервер запущен: http://localhost:${CONFIG.PORT}`);
    console.log(`📁 Папка данных: ${path.resolve(CONFIG.DATA_FOLDER)}`);
    console.log(`💾 База данных: ${path.resolve(CONFIG.DB_PATH)}`);
    console.log(`\n💡 Положите CSV файлы в папку "${CONFIG.DATA_FOLDER}"\n`);
  });
}

startServer();
