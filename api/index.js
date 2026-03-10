import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import { google } from 'googleapis';

dotenv.config();

const app = express();
app.use(cors());
app.use(express.json());

// ─── Auth ──────────────────────────────────────────────────────────────────────

function getAuth() {
  let credentials;

  if (process.env.GOOGLE_CREDENTIALS_JSON) {
    credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON);
  } else if (process.env.GOOGLE_CLIENT_EMAIL && process.env.GOOGLE_PRIVATE_KEY) {
    credentials = {
      type: 'service_account',
      client_email: process.env.GOOGLE_CLIENT_EMAIL,
      // dotenv v16+ 在雙引號字串中會自動展開 \n，這裡雙重保險
      private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
    };
  } else {
    throw new Error(
      'Google credentials not configured. Set GOOGLE_CREDENTIALS_JSON or GOOGLE_CLIENT_EMAIL + GOOGLE_PRIVATE_KEY in .env'
    );
  }

  return new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}

const auth = getAuth();
const sheets = google.sheets({ version: 'v4', auth });

// ─── Helpers ───────────────────────────────────────────────────────────────────

function getSheetId(sheetName) {
  const key = `SHEET_ID_${sheetName.toUpperCase()}`;
  const id = process.env[key];
  if (!id) throw new Error(`Sheet ID not found for "${sheetName}". Set ${key} in .env`);
  return id;
}

function rowsToObjects(rows) {
  if (!rows || rows.length < 2) return [];
  const [headers, ...data] = rows;
  return data.map(row =>
    Object.fromEntries(headers.map((h, i) => [h, row[i] ?? '']))
  );
}

async function getRange(sheetId, range) {
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: sheetId, range });
  return res.data.values || [];
}

// 欄位索引（0-based）轉 A1 欄位字母，例如 0→A, 25→Z, 26→AA
function colToLetter(index) {
  let letter = '';
  let n = index + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}

// ─── Route Registry ────────────────────────────────────────────────────────────
// 所有功能都在這裡集中定義，新增功能時只需在此加一筆



// 將路徑中的 :sheet 替換為 test（用於 / 入口說明）
function toExamplePath(path) {
  return path.replace(':sheet', 'test');
}

// ─── Routes ────────────────────────────────────────────────────────────────────

// GET / — API 入口說明（只顯示 test 範例路徑）
app.get('/', (req, res) => {
  res.json({
    message: 'Google Sheets API',
    version: '2.0.0',
    endpoints: ROUTES.map(r => ({
      name: r.name,
      method: r.method,
      path: toExamplePath(r.path),
      description: r.description,
    })),
  });
});

// GET /api/health
app.get('/api/health', (req, res) => {
  res.json({ status: 'OK', timestamp: new Date().toISOString() });
});

// GET /api/:sheet/tabsName — 取得所有分頁名稱
app.get('/api/:sheet/tabsName', async (req, res) => {
  try {
    const { sheet } = req.params;
    const sheetId = getSheetId(sheet);
    const response = await sheets.spreadsheets.get({
      spreadsheetId: sheetId,
      fields: 'sheets.properties.title',
    });
    const tabs = response.data.sheets.map(s => s.properties.title);
    res.json({ success: true, sheet, tabCount: tabs.length, tabs });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// ─── HowToUseForAgent 共用文件生成函式 ────────────────────────────────────────
// tabs 為空 → 通用說明模式（用 :tab 佔位）
// tabs 有值 → 鎖定模式（直接給出可用的完整 URL，明確告知 agent 只能動這些分頁）

const API_CONTEXT = {
  what: '這是一個串接 Google Sheets 的 API，讓程式可以透過 HTTP 讀寫試算表資料。',
  currentSheet: '目前串接的 Sheet 為開發測試用途（/api/test/...），供開發期間驗證功能使用。',
  deployedAt: 'https://greenbox-sheets-api.vercel.app',
  devNote:
    '開發完成後，若需要將路徑切換至正式 Sheet，請通知 Glen 協助更新 API 路徑與對應的 Sheet ID。',
};

function buildHowToUse(tabs = [], sheet = 'test') {
  const locked = tabs.length > 0;

  // 鎖定模式：為每個分頁生成一組完整可用的 URL 清單
  if (locked) {
    const operations = {};
    for (const tab of tabs) {
      const encodedTab = encodeURIComponent(tab); // 中文分頁名稱必須編碼才能正確放入 URL
      operations[tab] = {
        getRaw: {
          method: 'GET',
          url: `/api/${sheet}/tabRaw=${encodedTab}`,
          description: '取得此分頁的原始資料（二維陣列，不處理標題），用於了解分頁結構',
        },
        getAllRows: {
          method: 'GET',
          url: `/api/${sheet}/tab=${encodedTab}`,
          description: '取得此分頁的全部資料列（回傳物件陣列，第一行為欄位名稱）',
        },
        getRow: {
          method: 'GET',
          url: `/api/${sheet}/tab=${encodedTab}/row=N`,
          description: '取得第 N 筆資料，將 N 替換為正整數（1 = 第一筆資料，不含標題列）',
        },
        getRows: {
          method: 'GET',
          url: `/api/${sheet}/tab=${encodedTab}/row=X-Y`,
          description: '取得第 X 到第 Y 筆資料（含頭尾），X、Y 替換為正整數且 X <= Y',
        },
        append: {
          method: 'POST',
          url: `/api/${sheet}/tab=${encodedTab}`,
          description: '新增資料到此分頁末尾，支援三種格式',
          body: '單筆: { "values": ["v1","v2",...] } | 多筆陣列: { "values": [["v1","v2"],[...]] } | 多筆物件: { "rows": [{"欄位名":"值",...},...] }',
          warning: '使用 rows（物件格式）時，物件的 key 必須與該分頁的標題列欄位名稱完全一致（大小寫相同）。不符合的 key 會被忽略，該欄位將寫入空白。建議先用 GET 取得資料確認欄位名稱後再進行寫入。',
        },
        updateRow: {
          method: 'PUT',
          url: `/api/${sheet}/tab=${encodedTab}/row=N`,
          description: '覆寫第 N 筆資料，N 替換為正整數',
          body: '{ "values": ["欄位1值", "欄位2值", ...] }',
        },
        deleteRow: {
          method: 'DELETE',
          url: `/api/${sheet}/tab=${encodedTab}/row=N`,
          description: '清空第 N 筆資料的內容（列保留），N 替換為正整數',
        },
        addColumn: {
          method: 'POST',
          url: `/api/${sheet}/tab=${encodedTab}/col`,
          description: '在標題列末尾新增一個欄位，可同時填入各列的值。欄位名稱不可與現有欄位重複',
          body: '{ "name": "新欄位名稱", "values": ["row1值", "row2值", ...] }，values 為選填',
        },
        renameColumn: {
          method: 'PUT',
          url: `/api/${sheet}/tab=${encodedTab}/col`,
          description: '修改欄位名稱。from 必須是現有欄位，to 不可與現有欄位重複',
          body: '{ "from": "舊欄位名稱", "to": "新欄位名稱" }',
        },
      };
    }

    return {
      mode: 'locked',
      context: API_CONTEXT,
      instruction:
        '你只能操作 allowedTabs 中列出的分頁。' +
        'operations 中已提供每個分頁的完整 URL，URL 裡的分頁名稱已 URL 編碼（encodeURIComponent），可直接使用，不可自行修改。' +
        '只有 N、X、Y 需要替換為實際數字。不要對其他分頁進行任何操作。',
      allowedTabs: tabs,
      urlEncoding: '所有 URL 中的分頁名稱（tab= 後的部分）均已使用 encodeURIComponent 編碼。若要自行構造 URL，中文或特殊字元的分頁名稱必須先經過 encodeURIComponent 處理。',
      rowNumbering: 'row 從 1 開始，不含標題列。row=1 是第一筆資料（Google Sheets 第 2 行）。',
      operations,
    };
  }

  // 通用模式：完整說明所有功能，tab 用 :tab 佔位
  return {
    mode: 'generic',
    context: API_CONTEXT,
    urlEncoding: `當分頁名稱包含中文或特殊字元時，必須使用 encodeURIComponent 編碼後再放入 URL。例：分頁「測試」→ encodeURIComponent("測試") = "%E6%B8%AC%E8%A9%A6"，URL 為 /api/${sheet}/tab=%E6%B8%AC%E8%A9%A6。建議先呼叫 /api/${sheet}/tabsName 取得分頁名稱，再自行編碼組合 URL。`,
    rowNumbering: 'row 從 1 開始，不含標題列。row=1 是第一筆資料（Google Sheets 第 2 行）。',
    endpoints: [
      { method: 'GET',    url: `/api/${sheet}/tabsName`,         description: '取得所有分頁名稱（原始名稱，未編碼）' },
      { method: 'GET',    url: `/api/${sheet}/tabRaw=:tab`,      description: '取得分頁原始資料（二維陣列，不處理標題），供 Agent 了解分頁結構' },
      { method: 'GET',    url: `/api/${sheet}/tab=:tab`,         description: '取得分頁全部資料列，:tab 為 encodeURIComponent 編碼後的分頁名稱' },
      { method: 'GET',    url: `/api/${sheet}/tab=:tab/row=N`,   description: '取得第 N 筆資料' },
      { method: 'GET',    url: `/api/${sheet}/tab=:tab/row=X-Y`, description: '取得第 X～Y 筆資料' },
      { method: 'POST',   url: `/api/${sheet}/tab=:tab`,         description: '新增資料（單/多筆）  body: { values:[...或[[...]]] } 或 { rows:[{欄位名:值,...},...] }。注意：rows 格式的 key 必須與分頁標題列欄位名稱完全一致，不符的 key 將被忽略寫入空白。' },
      { method: 'POST',   url: `/api/${sheet}/tab=:tab/col`,     description: '新增欄位（可同時填值）  body: { name: "欄位名稱", values: [...] }' },
      { method: 'PUT',    url: `/api/${sheet}/tab=:tab/col`,     description: '修改欄位名稱  body: { from: "舊名稱", to: "新名稱" }' },
      { method: 'PUT',    url: `/api/${sheet}/tab=:tab/row=N`,   description: '更新第 N 筆資料  body: { values: [...] }' },
      { method: 'DELETE', url: `/api/${sheet}/tab=:tab/row=N`,   description: '清空第 N 筆資料' },
    ],
    tip: `若要鎖定特定分頁，改呼叫 GET /api/${sheet}/HowToUseForAgent/{分頁名稱} 取得專屬說明（URL 中的分頁名稱本身也需要 encodeURIComponent）。`,
  };
}

// GET /api/:sheet/HowToUseForAgent — 通用說明（tab 用佔位符顯示）
app.get('/api/:sheet/HowToUseForAgent', (req, res) => {
  res.json(buildHowToUse([], req.params.sheet));
});

// GET /api/:sheet/HowToUseForAgent/分頁1/分頁2/... — 指定一或多個分頁，範例路徑預填
app.get('/api/:sheet/HowToUseForAgent/*tabs', (req, res) => {
  try {
    const { sheet } = req.params;
    // Express 5 中萬用字元參數可能是字串或陣列，兩種都處理
    const raw = req.params.tabs;
    const tabString = Array.isArray(raw) ? raw.join('/') : (raw ?? '');
    const tabs = tabString.split('/').filter(Boolean);
    res.json(buildHowToUse(tabs, sheet));
  } catch (e) {
    console.error('HowToUseForAgent error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// GET /api/debug/auth — 測試 Google 憑證是否正常
app.get('/api/debug/auth', async (req, res) => {
  try {
    const client = await auth.getClient();
    const token = await client.getAccessToken();
    res.json({
      success: true,
      email: process.env.GOOGLE_CLIENT_EMAIL,
      hasToken: !!token.token,
    });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// GET /api/:sheet/tabRaw=:tab — 取得分頁原始資料（二維陣列，不處理標題）
app.get('/api/:sheet/tabRaw=:tab', async (req, res) => {
  try {
    const { sheet, tab } = req.params;
    const sheetId = getSheetId(sheet);
    const rows = await getRange(sheetId, tab);

    res.json({
      success: true,
      sheet,
      tab,
      rowCount: rows.length,
      colCount: rows.length > 0 ? Math.max(...rows.map(r => r.length)) : 0,
      rows,
    });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// GET /api/:sheet/tab=:tab — 取得分頁全部資料
app.get('/api/:sheet/tab=:tab', async (req, res) => {
  try {
    const { sheet, tab } = req.params;
    const sheetId = getSheetId(sheet);
    const rows = await getRange(sheetId, tab);
    const data = rowsToObjects(rows);

    res.json({
      success: true,
      sheet,
      tab,
      rowCount: data.length,
      data,
    });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// GET /api/:sheet/tab=:tab/row=:startRow-:endRow — 取得指定範圍資料列
// 範圍路由必須在單行路由之前，否則 :row 會吃掉整個 "X-Y" 字串
app.get('/api/:sheet/tab=:tab/row=:startRow-:endRow', async (req, res) => {
  try {
    const { sheet, tab } = req.params;
    const startRow = parseInt(req.params.startRow);
    const endRow = parseInt(req.params.endRow);

    if (isNaN(startRow) || isNaN(endRow) || startRow < 1 || endRow < startRow) {
      return res.status(400).json({ error: 'row 範圍無效，需滿足 X >= 1 且 X <= Y' });
    }

    const sheetId = getSheetId(sheet);
    const sheetStart = startRow + 1; // 加 1 跳過標題列
    const sheetEnd = endRow + 1;

    const [headerRows, dataRows] = await Promise.all([
      getRange(sheetId, `${tab}!1:1`),
      getRange(sheetId, `${tab}!${sheetStart}:${sheetEnd}`),
    ]);

    const headers = headerRows[0];
    if (!headers) {
      return res.status(404).json({ success: false, error: '找不到標題列' });
    }

    const data = dataRows.map(row =>
      Object.fromEntries(headers.map((h, j) => [h, row[j] ?? '']))
    );

    res.json({ success: true, sheet, tab, startRow, endRow, rowCount: data.length, data });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// GET /api/:sheet/tab=:tab/row=:row — 取得指定資料列
// row 從 1 開始，對應到第一筆「資料列」（標題列不計入）
app.get('/api/:sheet/tab=:tab/row=:row', async (req, res) => {
  try {
    const { sheet, tab } = req.params;
    const row = parseInt(req.params.row);

    if (isNaN(row) || row < 1) {
      return res.status(400).json({ error: 'row 必須是大於 0 的整數' });
    }

    const sheetId = getSheetId(sheet);
    const sheetRow = row + 1; // 加 1 跳過標題列（第 1 行）

    const [headerRows, dataRows] = await Promise.all([
      getRange(sheetId, `${tab}!1:1`),
      getRange(sheetId, `${tab}!${sheetRow}:${sheetRow}`),
    ]);

    const headers = headerRows[0];
    const rowData = dataRows[0];

    if (!headers || !rowData) {
      return res.status(404).json({ success: false, error: `找不到第 ${row} 筆資料` });
    }

    const data = Object.fromEntries(headers.map((h, i) => [h, rowData[i] ?? '']));

    res.json({ success: true, sheet, tab, row, data });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// POST /api/:sheet/tab=:tab — 新增資料（單筆或多筆，自動判斷）
// 單筆：{ values: ["欄位1值", "欄位2值", ...] }
// 多筆陣列：{ values: [["v1","v2",...], ["v3","v4",...]] }
// 多筆物件：{ rows: [{ 欄位名: 值, ... }, ...] }（自動對齊標題列）
app.post('/api/:sheet/tab=:tab', async (req, res) => {
  try {
    const { sheet, tab } = req.params;
    const { values, rows } = req.body;
    const sheetId = getSheetId(sheet);
    let writeData;

    if (rows !== undefined) {
      // 物件陣列格式：讀取標題列自動對齊
      if (!Array.isArray(rows) || rows.length === 0) {
        return res.status(400).json({ error: 'rows 需為非空陣列' });
      }
      const headerRows = await getRange(sheetId, `${tab}!1:1`);
      const allHeaders = headerRows[0] ?? [];
      // Google Sheets API 回傳的 row 1 可能從 A 欄開始（含空格），
      // 裁掉前導空字串，確保與 append 寫入的起始欄對齊
      const firstNonEmpty = allHeaders.findIndex(h => h.trim() !== '');
      const headers = firstNonEmpty > 0 ? allHeaders.slice(firstNonEmpty) : allHeaders;
      if (headers.length === 0) {
        return res.status(400).json({ error: '找不到標題列，請先建立標題再使用 rows 格式' });
      }
      writeData = rows.map(row => headers.map(h => row[h] ?? ''));

    } else if (Array.isArray(values) && values.length > 0) {
      if (Array.isArray(values[0])) {
        // 多筆：二維陣列
        writeData = values;
      } else {
        // 單筆：一維陣列
        writeData = [values];
      }
    } else {
      return res.status(400).json({
        error: 'body 需包含：values（一維或二維陣列）或 rows（物件陣列）',
      });
    }

    const result = await sheets.spreadsheets.values.append({
      spreadsheetId: sheetId,
      range: tab,
      valueInputOption: 'USER_ENTERED',
      resource: { values: writeData },
    });

    res.json({
      success: true,
      sheet,
      tab,
      appendedRows: writeData.length,
      result: result.data,
    });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// POST /api/:sheet/tab=:tab/col — 在標題列末尾新增欄位，可同時填入各列的值
// body: { name: "欄位名稱", values: ["row1值", "row2值", ...] }  values 為選填
app.post('/api/:sheet/tab=:tab/col', async (req, res) => {
  try {
    const { sheet, tab } = req.params;
    const { name, values } = req.body;

    if (!name || typeof name !== 'string' || !name.trim()) {
      return res.status(400).json({ error: 'body 需包含 name: "欄位名稱"' });
    }
    if (values !== undefined && !Array.isArray(values)) {
      return res.status(400).json({ error: 'values 若提供需為陣列' });
    }

    const sheetId = getSheetId(sheet);

    // 讀取現有標題列，確認新欄位名稱不重複
    const headerRows = await getRange(sheetId, `${tab}!1:1`);
    const headers = headerRows[0] ?? [];

    if (headers.includes(name.trim())) {
      return res.status(400).json({ error: `欄位「${name}」已存在` });
    }

    const nextCol = colToLetter(headers.length);

    // 組合要寫入的資料：第一格是標題，後續是各列的值（row 2 開始）
    const writeData = [[name.trim()]];
    if (values && values.length > 0) {
      values.forEach(v => writeData.push([v ?? '']));
    }

    const result = await sheets.spreadsheets.values.update({
      spreadsheetId: sheetId,
      range: `${tab}!${nextCol}1`,
      valueInputOption: 'USER_ENTERED',
      resource: { values: writeData },
    });

    res.json({
      success: true,
      sheet,
      tab,
      addedColumn: name.trim(),
      position: nextCol,
      filledRows: values ? values.length : 0,
      result: result.data,
    });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// PUT /api/:sheet/tab=:tab/col — 修改欄位名稱
// body: { from: "舊欄位名稱", to: "新欄位名稱" }
app.put('/api/:sheet/tab=:tab/col', async (req, res) => {
  try {
    const { sheet, tab } = req.params;
    const { from, to } = req.body;

    if (!from || typeof from !== 'string' || !from.trim()) {
      return res.status(400).json({ error: 'body 需包含 from: "舊欄位名稱"' });
    }
    if (!to || typeof to !== 'string' || !to.trim()) {
      return res.status(400).json({ error: 'body 需包含 to: "新欄位名稱"' });
    }

    const sheetId = getSheetId(sheet);
    const headerRows = await getRange(sheetId, `${tab}!1:1`);
    const headers = headerRows[0] ?? [];

    const colIndex = headers.indexOf(from.trim());
    if (colIndex === -1) {
      return res.status(404).json({ error: `找不到欄位「${from}」` });
    }
    if (headers.includes(to.trim())) {
      return res.status(400).json({ error: `欄位「${to}」已存在` });
    }

    const colLetter = colToLetter(colIndex);
    await sheets.spreadsheets.values.update({
      spreadsheetId: sheetId,
      range: `${tab}!${colLetter}1`,
      valueInputOption: 'USER_ENTERED',
      resource: { values: [[to.trim()]] },
    });

    res.json({ success: true, sheet, tab, from: from.trim(), to: to.trim(), column: colLetter });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// PUT /api/:sheet/tab=:tab/row=:row — 更新指定資料列
// body: { values: ["欄位1", "欄位2", ...] }
app.put('/api/:sheet/tab=:tab/row=:row', async (req, res) => {
  try {
    const { sheet, tab } = req.params;
    const row = parseInt(req.params.row);
    const { values } = req.body;

    if (isNaN(row) || row < 1) {
      return res.status(400).json({ error: 'row 必須是大於 0 的整數' });
    }
    if (!values || !Array.isArray(values)) {
      return res.status(400).json({ error: 'body 需包含 values: [...]' });
    }

    const sheetId = getSheetId(sheet);
    const sheetRow = row + 1;
    const result = await sheets.spreadsheets.values.update({
      spreadsheetId: sheetId,
      range: `${tab}!${sheetRow}:${sheetRow}`,
      valueInputOption: 'USER_ENTERED',
      resource: { values: [values] },
    });

    res.json({ success: true, sheet, tab, row, result: result.data });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// DELETE /api/:sheet/tab=:tab/row=:row — 清空指定資料列
app.delete('/api/:sheet/tab=:tab/row=:row', async (req, res) => {
  try {
    const { sheet, tab } = req.params;
    const row = parseInt(req.params.row);

    if (isNaN(row) || row < 1) {
      return res.status(400).json({ error: 'row 必須是大於 0 的整數' });
    }

    const sheetId = getSheetId(sheet);
    const sheetRow = row + 1;
    const result = await sheets.spreadsheets.values.clear({
      spreadsheetId: sheetId,
      range: `${tab}!${sheetRow}:${sheetRow}`,
    });

    res.json({ success: true, sheet, tab, row, result: result.data });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// ─── 404 & Error handlers ──────────────────────────────────────────────────────

app.use((req, res) => {
  res.status(404).json({ error: 'Endpoint not found' });
});

app.use((err, req, res, _next) => {
  console.error(err);
  res.status(500).json({ error: 'Internal server error' });
});

// ─── Start (local dev only) ────────────────────────────────────────────────────

if (!process.env.VERCEL) {
  const PORT = process.env.PORT || 3000;
  app.listen(PORT, () => console.log(`Server running at http://localhost:${PORT}`));
}

export default app;
const ROUTES = [
  {
    method: 'GET',
    path: '/api/health',
    name: 'health',
    description: 'API運作狀況的健康檢查',
  },
  {
    method: 'GET',
    path: '/api/:sheet/HowToUseForAgent',
    name: 'howToUse',
    description: '回傳完整 API 使用說明（供 AI Agent 理解本 API 的所有功能與用法）',
  },
  {
    method: 'GET',
    path: '/api/:sheet/HowToUseForAgent/:tab',
    name: 'howToUseForTab',
    description: '回傳完整 API 使用說明，並將操作範圍鎖定於指定分頁。Agent 只能對 allowedTabs 內的分頁進行操作。支援多分頁：/HowToUseForAgent/分頁1/分頁2/...',
  },
  {
    method: 'GET',
    path: '/api/:sheet/tabsName',
    name: 'getTabs',
    description: '取得測試用 Sheet 的所有分頁名稱（原始名稱，未編碼）',
  },
  {
    method: 'GET',
    path: '/api/:sheet/tabRaw=:tab',
    name: 'getTabRaw',
    description: '取得指定分頁的原始資料（二維陣列，不處理標題），供 Agent 了解分頁結構',
  },
  {
    method: 'GET',
    path: '/api/:sheet/tab=:tab',
    name: 'getTab',
    description: '取得指定分頁的全部資料列（物件格式，第一行自動作為欄位名稱）',
  },
  {
    method: 'GET',
    path: '/api/:sheet/tab=:tab/row=:row',
    name: 'getRow',
    description: '取得指定分頁第 N 筆資料（row 從 1 開始，不含標題列）',
  },
  {
    method: 'GET',
    path: '/api/:sheet/tab=:tab/row=:startRow-:endRow',
    name: 'getRows',
    description: '取得指定分頁第 X～Y 筆資料（含頭尾，回傳陣列）',
  },
  {
    method: 'POST',
    path: '/api/:sheet/tab=:tab',
    name: 'append',
    description: '新增資料（單筆或多筆）。單筆: { values:[...] }，多筆陣列: { values:[[...],[...]] }，多筆物件: { rows:[{欄位:值},...] }',
  },
  {
    method: 'POST',
    path: '/api/:sheet/tab=:tab/col',
    name: 'addColumn',
    description: '在指定分頁的標題列末尾新增一個欄位  body: { name: "欄位名稱" }',
  },
  {
    method: 'PUT',
    path: '/api/:sheet/tab=:tab/col',
    name: 'renameColumn',
    description: '修改欄位名稱  body: { from: "舊名稱", to: "新名稱" }',
  },
  {
    method: 'PUT',
    path: '/api/:sheet/tab=:tab/row=:row',
    name: 'updateRow',
    description: '覆寫指定分頁第 N 筆資料  body: { values: [...] }',
  },
  {
    method: 'DELETE',
    path: '/api/:sheet/tab=:tab/row=:row',
    name: 'deleteRow',
    description: '清空指定分頁第 N 筆資料（列保留、內容清除）',
  },
];