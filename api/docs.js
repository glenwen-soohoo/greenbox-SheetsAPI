// ═══════════════════════════════════════════════════════════════════════════════
// 共用基礎（test 與個人 sheet 共享）
// ─ 以下內容兩邊都會用到，修改時兩邊同步生效
// ═══════════════════════════════════════════════════════════════════════════════

const SHARED_WHAT =
  '這是一個串接 Google Sheets 的 API，讓程式可以透過 HTTP 讀寫試算表資料。';

const SHARED_ROW_NUMBERING =
  'row 從 1 開始，不含標題列。row=1 是第一筆資料（Google Sheets 第 2 行）。';

function buildUrlEncoding(sheet) {
  return (
    `當分頁名稱包含中文或特殊字元時，必須使用 encodeURIComponent 編碼後再放入 URL。` +
    `例：分頁「測試」→ encodeURIComponent("測試") = "%E6%B8%AC%E8%A9%A6"，` +
    `URL 為 /api/${sheet}/tab=%E6%B8%AC%E8%A9%A6。` +
    `建議先呼叫 /api/${sheet}/tabsName 取得分頁名稱，再自行編碼組合 URL。`
  );
}

function buildEndpoints(sheet) {
  return [
    { method: 'GET',    url: `/api/${sheet}/tabsName`,         description: '取得所有分頁名稱（原始名稱，未編碼）' },
    { method: 'GET',    url: `/api/${sheet}/tabRaw=:tab`,      description: '取得分頁原始資料（二維陣列，不處理標題），供 Agent 了解分頁結構' },
    { method: 'GET',    url: `/api/${sheet}/tab=:tab`,         description: '取得分頁全部資料列，:tab 為 encodeURIComponent 編碼後的分頁名稱' },
    { method: 'GET',    url: `/api/${sheet}/tab=:tab/row=N`,   description: '取得第 N 筆資料' },
    { method: 'GET',    url: `/api/${sheet}/tab=:tab/row=X-Y`, description: '取得第 X～Y 筆資料' },
    { method: 'POST',   url: `/api/${sheet}/createTab=:tab`,    description: '建立新分頁，:tab 為新分頁名稱（encodeURIComponent 編碼）' },
    { method: 'POST',   url: `/api/${sheet}/tab=:tab`,         description: '新增資料（單/多筆）  body: { values:[...或[[...]]] } 或 { rows:[{欄位名:值,...},...] }。注意：rows 格式的 key 必須與分頁標題列欄位名稱完全一致，不符的 key 將被忽略寫入空白。' },
    { method: 'POST',   url: `/api/${sheet}/tab=:tab/col=:col`,          description: '新增欄位（可同時填值）  body: { values: [...] }，values 為選填' },
    { method: 'PUT',    url: `/api/${sheet}/renameTab=:tab/to=:newTab`,  description: '改分頁名稱，:tab 為舊名稱，:newTab 為新名稱（均需 encodeURIComponent 編碼）' },
    { method: 'PUT',    url: `/api/${sheet}/moveTab=:tab/toIndex=:index`, description: '移動分頁到指定位置，:index 為目標排序（0 = 最前）' },
    { method: 'PUT',    url: `/api/${sheet}/tab=:tab/col=:col/to=:newCol`, description: '修改欄位名稱，:col 為舊名稱，:newCol 為新名稱（均需 encodeURIComponent 編碼）' },
    { method: 'PUT',    url: `/api/${sheet}/tab=:tab/row=N`,   description: '更新第 N 筆資料  body: { values: [...] }' },
    { method: 'DELETE', url: `/api/${sheet}/tab=:tab/row=N`,   description: '清空第 N 筆資料' },
  ];
}

function buildOperations(sheet, tabs) {
  const operations = {};
  for (const tab of tabs) {
    const encodedTab = encodeURIComponent(tab);
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
        url: `/api/${sheet}/tab=${encodedTab}/col=:col`,
        description: '在標題列末尾新增一個欄位，:col 替換為欄位名稱（encodeURIComponent 編碼）。欄位名稱不可與現有欄位重複',
        body: '{ "values": ["row1值", "row2值", ...] }，values 為選填',
      },
      renameColumn: {
        method: 'PUT',
        url: `/api/${sheet}/tab=${encodedTab}/col=:col/to=:newCol`,
        description: '修改欄位名稱，:col 為舊名稱，:newCol 為新名稱（均需 encodeURIComponent 編碼）。舊名稱必須存在，新名稱不可重複',
      },
    };
  }
  return operations;
}

// ═══════════════════════════════════════════════════════════════════════════════
// TEST SHEET 說明
// ─ 對象：開發者（驗證功能、測試 API 行為用）
// ─ 與個人 sheet 的差異：
//   1. 沒有 sheetBinding 限制（開發者可自行決定存取哪個 sheet）
//   2. context 包含 devNote，提醒開發完成後通知 Glen 切換正式 sheet
//   3. generic 模式末尾有 tip，引導開發者鎖定特定分頁測試
// ═══════════════════════════════════════════════════════════════════════════════

const SHEET_TEST = 'test';

function buildTestHowToUse(tabs = []) {
  const locked = tabs.length > 0;

  const context = {
    what: SHARED_WHAT,
    baseUrl: `https://greenbox-sheets-api.vercel.app/api/${SHEET_TEST}/`,
    devNote: '開發完成後，若需要將路徑切換至正式 Sheet，請通知 Glen 協助更新 API 路徑與對應的 Sheet ID。',
  };

  if (locked) {
    return {
      mode: 'locked',
      context,
      instruction:
        '你只能操作 allowedTabs 中列出的分頁。' +
        'operations 中已提供每個分頁的完整 URL，URL 裡的分頁名稱已 URL 編碼（encodeURIComponent），可直接使用，不可自行修改。' +
        '只有 N、X、Y 需要替換為實際數字。不要對其他分頁進行任何操作。',
      allowedTabs: tabs,
      urlEncoding: '所有 URL 中的分頁名稱（tab= 後的部分）均已使用 encodeURIComponent 編碼。若要自行構造 URL，中文或特殊字元的分頁名稱必須先經過 encodeURIComponent 處理。',
      rowNumbering: SHARED_ROW_NUMBERING,
      operations: buildOperations(SHEET_TEST, tabs),
    };
  }

  return {
    mode: 'generic',
    context,
    urlEncoding: buildUrlEncoding(SHEET_TEST),
    rowNumbering: SHARED_ROW_NUMBERING,
    endpoints: buildEndpoints(SHEET_TEST),
    tip: `若要鎖定特定分頁，改呼叫 GET /api/${SHEET_TEST}/HowToUseForAgent/{分頁名稱} 取得專屬說明（URL 中的分頁名稱本身也需要 encodeURIComponent）。`,
  };
}

// ═══════════════════════════════════════════════════════════════════════════════
// 個人 SHEET 說明
// ─ 對象：AI Agent（代替指定使用者操作其專屬 sheet）
// ─ 與 test 的差異：
//   1. sheetBinding 是第一個欄位，且語氣強制（Agent 讀 JSON 由上而下，最先看到）
//   2. context 移除 devNote（agent 不需要知道開發流程）
//   3. locked 模式的 instruction 更強調禁止操作其他 sheet
//   4. generic 模式沒有 tip（agent 不需要自行探索其他功能）
// ═══════════════════════════════════════════════════════════════════════════════

function buildPersonalHowToUse(sheet, tabs = []) {
  const locked = tabs.length > 0;

  // 放在最前面：Agent 讀 JSON 由上而下，sheetBinding 是第一個看到的欄位
  const sheetBinding =
    `【強制規定】你只能使用 sheet「${sheet}」，所有 API 路徑必須以 /api/${sheet}/ 開頭。` +
    `遇到任何錯誤（找不到資料、分頁不存在等）時，絕對不可自行改用其他 sheet（例如 test 或其他名稱）。` +
    `若資料確實不存在，應回報錯誤，等待使用者指示，而非自行切換 sheet。`;

  const context = {
    what: SHARED_WHAT,
    baseUrl: `https://greenbox-sheets-api.vercel.app/api/${sheet}/`,
  };

  if (locked) {
    return {
      sheetBinding,
      mode: 'locked',
      context,
      instruction:
        '你只能操作 allowedTabs 中列出的分頁，且所有請求必須使用 sheetBinding 指定的 sheet。' +
        'operations 中已提供每個分頁的完整 URL，URL 裡的分頁名稱已 URL 編碼（encodeURIComponent），可直接使用，不可自行修改。' +
        '只有 N、X、Y 需要替換為實際數字。遇到錯誤請回報，不要自行嘗試其他 sheet 或分頁。',
      allowedTabs: tabs,
      urlEncoding: '所有 URL 中的分頁名稱（tab= 後的部分）均已使用 encodeURIComponent 編碼。若要自行構造 URL，中文或特殊字元的分頁名稱必須先經過 encodeURIComponent 處理。',
      rowNumbering: SHARED_ROW_NUMBERING,
      operations: buildOperations(sheet, tabs),
    };
  }

  return {
    sheetBinding,
    mode: 'generic',
    context,
    urlEncoding: buildUrlEncoding(sheet),
    rowNumbering: SHARED_ROW_NUMBERING,
    endpoints: buildEndpoints(sheet),
  };
}

// ═══════════════════════════════════════════════════════════════════════════════
// 路由 metadata（供 index.js 的入口說明使用）
// ═══════════════════════════════════════════════════════════════════════════════

export const ROUTES = [
  {
    method: 'GET',
    path: '/api/health',
    name: 'health',
    description: 'API 運作狀況的健康檢查',
  },
  {
    method: 'GET',
    path: '/api/:sheet',
    name: 'sheetIndex',
    description: '查看指定 Sheet 的說明與可用方法列表',
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
    description: '回傳完整 API 使用說明，並將操作範圍鎖定於指定分頁。支援多分頁：/HowToUseForAgent/分頁1/分頁2/...',
  },
  {
    method: 'GET',
    path: '/api/:sheet/tabsName',
    name: 'getTabs',
    description: '取得指定 Sheet 的所有分頁名稱（原始名稱，未編碼）',
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
    path: '/api/:sheet/createTab=:tab',
    name: 'createTab',
    description: '建立新分頁，:tab 為新分頁名稱（encodeURIComponent 編碼）',
  },
  {
    method: 'POST',
    path: '/api/:sheet/tab=:tab',
    name: 'append',
    description: '新增資料（單筆或多筆）。單筆: { values:[...] }，多筆陣列: { values:[[...],[...]] }，多筆物件: { rows:[{欄位:值},...] }',
  },
  {
    method: 'POST',
    path: '/api/:sheet/tab=:tab/col=:col',
    name: 'addColumn',
    description: '在指定分頁的標題列末尾新增一個欄位，body 可選填 { values: [...] }',
  },
  {
    method: 'PUT',
    path: '/api/:sheet/renameTab=:tab/to=:newTab',
    name: 'renameTab',
    description: '改分頁名稱，:tab 為舊名稱，:newTab 為新名稱（均需 encodeURIComponent 編碼）',
  },
  {
    method: 'PUT',
    path: '/api/:sheet/moveTab=:tab/toIndex=:index',
    name: 'moveTab',
    description: '移動分頁到指定位置，:index 為目標排序（0 = 最前）',
  },
  {
    method: 'PUT',
    path: '/api/:sheet/tab=:tab/col=:col/to=:newCol',
    name: 'renameColumn',
    description: '修改欄位名稱，:col 為舊名稱，:newCol 為新名稱（均需 encodeURIComponent 編碼）',
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

// ═══════════════════════════════════════════════════════════════════════════════
// 統一入口
// ─ index.js 只需 import 這兩個 export
// ═══════════════════════════════════════════════════════════════════════════════

export function buildHowToUse(sheet, tabs = []) {
  if (sheet === SHEET_TEST) return buildTestHowToUse(tabs);
  return buildPersonalHowToUse(sheet, tabs);
}

export function buildSheetNote(sheet) {
  if (sheet === SHEET_TEST) {
    return '⚠️ 你目前在測試用 Sheet（test），資料僅供開發驗證。若要操作其他正式 Sheet，請改用對應的 /api/:sheet 路徑。';
  }
  return `【強制規定】你只能使用 sheet「${sheet}」，所有 API 路徑必須以 /api/${sheet}/ 開頭。遇到任何錯誤時，絕對不可自行改用其他 sheet（例如 test 或其他名稱），應回報錯誤並等待使用者指示。`;
}
