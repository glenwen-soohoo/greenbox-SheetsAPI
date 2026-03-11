import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import { google } from 'googleapis';
import { buildHowToUse, buildSheetNote, ROUTES } from './docs.js';

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

// ─── Routes ────────────────────────────────────────────────────────────────────

// GET / — API 入口導覽
app.get('/', (req, res) => {
  res.json({
    message: 'Google Sheets API',
    version: '2.0.0',
    usage: '透過 /api/:sheet 指定你要操作的 Sheet，進入後可查看該 Sheet 的可用方法。',
    example: 'https://greenbox-sheets-api.vercel.app/api/glen',
    notSureWhichSheet: {
      hint: '如果不確定要連到哪個 Sheet，可以先進入測試區：',
      url: 'https://greenbox-sheets-api.vercel.app/api/test',
      warning: '⚠️ 這是測試用的 Sheet，資料僅供開發驗證。若需操作正式資料，請使用對應的正確路徑。',
    },
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

// GET /api/:sheet/HowToUseForAgent — 通用說明（tab 用佔位符顯示）
app.get('/api/:sheet/HowToUseForAgent', (req, res) => {
  res.json(buildHowToUse(req.params.sheet));
});

// GET /api/:sheet/HowToUseForAgent/分頁1/分頁2/... — 指定一或多個分頁，範例路徑預填
app.get('/api/:sheet/HowToUseForAgent/*tabs', (req, res) => {
  try {
    const { sheet } = req.params;
    // Express 5 中萬用字元參數可能是字串或陣列，兩種都處理
    const raw = req.params.tabs;
    const tabString = Array.isArray(raw) ? raw.join('/') : (raw ?? '');
    const tabs = tabString.split('/').filter(Boolean);
    res.json(buildHowToUse(sheet, tabs));
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

// POST /api/:sheet/tab=:tab/col=:col — 在標題列末尾新增欄位，可同時填入各列的值
// body: { values: ["row1值", "row2值", ...] }  values 為選填
app.post('/api/:sheet/tab=:tab/col=:col', async (req, res) => {
  try {
    const { sheet, tab, col } = req.params;
    const { values } = req.body;

    if (values !== undefined && !Array.isArray(values)) {
      return res.status(400).json({ error: 'values 若提供需為陣列' });
    }

    const sheetId = getSheetId(sheet);

    // 讀取現有標題列，確認新欄位名稱不重複
    const headerRows = await getRange(sheetId, `${tab}!1:1`);
    const headers = headerRows[0] ?? [];

    if (headers.includes(col)) {
      return res.status(400).json({ error: `欄位「${col}」已存在` });
    }

    const nextCol = colToLetter(headers.length);

    // 組合要寫入的資料：第一格是標題，後續是各列的值（row 2 開始）
    const writeData = [[col]];
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
      addedColumn: col,
      position: nextCol,
      filledRows: values ? values.length : 0,
      result: result.data,
    });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// PUT /api/:sheet/tab=:tab/col=:col/to=:newCol — 修改欄位名稱
app.put('/api/:sheet/tab=:tab/col=:col/to=:newCol', async (req, res) => {
  try {
    const { sheet, tab, col, newCol } = req.params;

    const sheetId = getSheetId(sheet);
    const headerRows = await getRange(sheetId, `${tab}!1:1`);
    const headers = headerRows[0] ?? [];

    const colIndex = headers.indexOf(col);
    if (colIndex === -1) {
      return res.status(404).json({ error: `找不到欄位「${col}」` });
    }
    if (headers.includes(newCol)) {
      return res.status(400).json({ error: `欄位「${newCol}」已存在` });
    }

    const colLetter = colToLetter(colIndex);
    await sheets.spreadsheets.values.update({
      spreadsheetId: sheetId,
      range: `${tab}!${colLetter}1`,
      valueInputOption: 'USER_ENTERED',
      resource: { values: [[newCol]] },
    });

    res.json({ success: true, sheet, tab, from: col, to: newCol, column: colLetter });
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

// POST /api/:sheet/createTab=:tab — 建立新分頁
app.post('/api/:sheet/createTab=:tab', async (req, res) => {
  try {
    const { sheet, tab } = req.params;
    const sheetId = getSheetId(sheet);
    const result = await sheets.spreadsheets.batchUpdate({
      spreadsheetId: sheetId,
      resource: {
        requests: [{ addSheet: { properties: { title: tab } } }],
      },
    });

    const newSheet = result.data.replies[0].addSheet.properties;
    res.json({ success: true, sheet, tab: newSheet.title, sheetId: newSheet.sheetId });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// PUT /api/:sheet/renameTab=:tab/to=:newTab — 改分頁名稱
app.put('/api/:sheet/renameTab=:tab/to=:newTab', async (req, res) => {
  try {
    const { sheet, tab, newTab } = req.params;
    const to = newTab;

    const sheetId = getSheetId(sheet);

    // 先取得目標分頁的 sheetId（數字 ID）
    const spreadsheet = await sheets.spreadsheets.get({
      spreadsheetId: sheetId,
      fields: 'sheets.properties',
    });
    const found = spreadsheet.data.sheets.find(s => s.properties.title === tab);
    if (!found) {
      return res.status(404).json({ success: false, error: `找不到分頁「${tab}」` });
    }

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: sheetId,
      resource: {
        requests: [{
          updateSheetProperties: {
            properties: { sheetId: found.properties.sheetId, title: to },
            fields: 'title',
          },
        }],
      },
    });

    res.json({ success: true, sheet, from: tab, to });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// PUT /api/:sheet/moveTab=:tab/toIndex=:index — 移動分頁到指定位置（0 = 最前）
app.put('/api/:sheet/moveTab=:tab/toIndex=:index', async (req, res) => {
  try {
    const { sheet, tab } = req.params;
    const index = parseInt(req.params.index);

    if (isNaN(index) || index < 0) {
      return res.status(400).json({ error: 'toIndex 必須是大於等於 0 的整數' });
    }

    const sheetId = getSheetId(sheet);
    const spreadsheet = await sheets.spreadsheets.get({
      spreadsheetId: sheetId,
      fields: 'sheets.properties',
    });
    const found = spreadsheet.data.sheets.find(s => s.properties.title === tab);
    if (!found) {
      return res.status(404).json({ success: false, error: `找不到分頁「${tab}」` });
    }

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: sheetId,
      resource: {
        requests: [{
          updateSheetProperties: {
            properties: { sheetId: found.properties.sheetId, index },
            fields: 'index',
          },
        }],
      },
    });

    res.json({ success: true, sheet, tab, toIndex: index });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// GET /api/:sheet — 與入口相同的 API 說明，但顯示指定 sheet 名稱
app.get('/api/:sheet', (req, res) => {
  const { sheet } = req.params;
  res.json({
    sheetNote: buildSheetNote(sheet),
    message: 'Google Sheets API',
    version: '2.0.0',
    sheet,
    endpoints: ROUTES.map(r => ({
      name: r.name,
      method: r.method,
      path: r.path.replace(':sheet', sheet),
      description: r.description,
    })),
  });
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