// ============================================================
// 進階練習：智慧會議記錄管理系統
// 主題：辦公室自動化 AI — 會議管理
// 對應：Session 1（基本介面、巨集、觸發器）
// ============================================================
// 情境說明：
//   公司每週有多場會議，需要：
//   1. 自動化會議記錄的建立與格式化
//   2. 用巨集快速套用標準格式
//   3. 觸發器在會議前自動建立記錄模板
//   4. 智慧判斷會議類型並套用對應模板
// ============================================================

// ============================================================
// 全域設定
// ============================================================
var 會議設定 = {
  公司名稱: "ABC 科技股份有限公司",
  預設會議室: ["大會議室 A", "小會議室 B", "視訊會議室 C", "主管辦公室"],
  會議類型: {
    "週會": { 顏色: "#1a73e8", 圖示: "📋", 時長: 60 },
    "專案會議": { 顏色: "#34a853", 圖示: "🚀", 時長: 90 },
    "客戶會議": { 顏色: "#ea4335", 圖示: "🤝", 時長: 60 },
    "部門會議": { 顏色: "#fbbc04", 圖示: "🏢", 時長: 45 },
    "臨時會議": { 顏色: "#9c27b0", 圖示: "⚡", 時長: 30 }
  }
};

// ============================================================
// 第一部分：自動建立會議記錄模板
// ============================================================

/**
 * 智慧建立會議記錄
 * 說明：根據會議類型自動產生對應的會議記錄模板
 */
function 建立會議記錄() {
  try {
    var ui = SpreadsheetApp.getUi();

    // 用對話框收集會議資訊
    var 類型回應 = ui.prompt(
      "📋 建立會議記錄",
      "請輸入會議類型（週會/專案會議/客戶會議/部門會議/臨時會議）：",
      ui.ButtonSet.OK_CANCEL
    );
    if (類型回應.getSelectedButton() !== ui.Button.OK) return;
    var 會議類型 = 類型回應.getResponseText().trim();

    var 主題回應 = ui.prompt("📋 會議主題", "請輸入會議主題：", ui.ButtonSet.OK_CANCEL);
    if (主題回應.getSelectedButton() !== ui.Button.OK) return;
    var 會議主題 = 主題回應.getResponseText().trim();

    var 主持人回應 = ui.prompt("📋 主持人", "請輸入主持人姓名：", ui.ButtonSet.OK_CANCEL);
    if (主持人回應.getSelectedButton() !== ui.Button.OK) return;
    var 主持人 = 主持人回應.getResponseText().trim();

    // 取得會議類型設定（若輸入不存在則用預設）
    var 類型設定 = 會議設定.會議類型[會議類型] || 會議設定.會議類型["臨時會議"];

    // 建立工作表
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 今天 = new Date();
    var 日期字串 = Utilities.formatDate(今天, "Asia/Taipei", "yyyyMMdd_HHmm");
    var 表名 = 類型設定.圖示 + " " + 會議類型 + "_" + 日期字串;

    var sheet = ss.insertSheet(表名, 0); // 插入在最前面

    // ===== 套用會議模板 =====
    _套用會議模板(sheet, {
      類型: 會議類型,
      主題: 會議主題,
      主持人: 主持人,
      日期: 今天,
      設定: 類型設定
    });

    Logger.log("✅ 會議記錄已建立：" + 表名);
    ui.alert("✅ 會議記錄「" + 會議主題 + "」已建立！\n\n工作表：" + 表名);

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ 錯誤：" + 錯誤.message);
  }
}

/**
 * 套用會議模板（內部函數）
 * @param {Sheet} sheet - 目標工作表
 * @param {Object} 資訊 - 會議資訊物件
 */
function _套用會議模板(sheet, 資訊) {
  var 顏色 = 資訊.設定.顏色;
  var 日期格式 = Utilities.formatDate(資訊.日期, "Asia/Taipei", "yyyy/MM/dd (E) HH:mm");

  // --- 標題區 ---
  sheet.getRange("A1:F1").merge();
  sheet.getRange("A1").setValue(會議設定.公司名稱 + " — 會議記錄");
  sheet.getRange("A1").setFontSize(16).setFontWeight("bold")
    .setHorizontalAlignment("center").setBackground(顏色).setFontColor("#fff");
  sheet.setRowHeight(1, 45);

  // --- 會議資訊區 ---
  var 資訊區 = [
    ["會議類型", 資訊.設定.圖示 + " " + 資訊.類型, "會議日期", 日期格式],
    ["會議主題", 資訊.主題, "預計時長", 資訊.設定.時長 + " 分鐘"],
    ["主 持 人", 資訊.主持人, "記 錄 者", "（請填寫）"],
    ["會 議 室", "（請選擇）", "出席人數", "（請填寫）"]
  ];

  sheet.getRange(3, 1, 4, 4).setValues(資訊區);

  // 標籤欄樣式
  for (var i = 3; i <= 6; i++) {
    sheet.getRange(i, 1).setBackground("#e8eaf6").setFontWeight("bold");
    sheet.getRange(i, 3).setBackground("#e8eaf6").setFontWeight("bold");
  }

  // --- 出席人員 ---
  sheet.getRange("A8:F8").merge();
  sheet.getRange("A8").setValue("📋 出席人員");
  sheet.getRange("A8").setFontSize(13).setFontWeight("bold").setBackground("#f5f5f5");

  sheet.getRange("A9:F9").setValues([["姓名", "部門", "職稱", "出席", "簽到時間", "備註"]]);
  sheet.getRange("A9:F9").setBackground(顏色).setFontColor("#fff").setFontWeight("bold")
    .setHorizontalAlignment("center");

  // 預留 10 列出席人員
  for (var j = 10; j <= 19; j++) {
    sheet.getRange(j, 4).setValue("☐");
    sheet.getRange(j, 4).setHorizontalAlignment("center");
  }

  // --- 議程 ---
  var 議程起始 = 21;
  sheet.getRange("A" + 議程起始 + ":F" + 議程起始).merge();
  sheet.getRange("A" + 議程起始).setValue("📝 會議議程");
  sheet.getRange("A" + 議程起始).setFontSize(13).setFontWeight("bold").setBackground("#f5f5f5");

  var 議程標題列 = 議程起始 + 1;
  sheet.getRange(議程標題列, 1, 1, 6)
    .setValues([["序號", "議題", "報告人", "預計時間", "決議", "備註"]]);
  sheet.getRange(議程標題列, 1, 1, 6)
    .setBackground(顏色).setFontColor("#fff").setFontWeight("bold").setHorizontalAlignment("center");

  // 根據會議類型預設不同議程
  var 預設議程 = _取得預設議程(資訊.類型);
  for (var k = 0; k < 預設議程.length; k++) {
    sheet.getRange(議程標題列 + 1 + k, 1).setValue(k + 1);
    sheet.getRange(議程標題列 + 1 + k, 2).setValue(預設議程[k]);
    sheet.getRange(議程標題列 + 1 + k, 1).setHorizontalAlignment("center");
  }

  // --- 決議事項 ---
  var 決議起始 = 議程標題列 + 預設議程.length + 3;
  sheet.getRange("A" + 決議起始 + ":F" + 決議起始).merge();
  sheet.getRange("A" + 決議起始).setValue("✅ 決議事項與待辦");
  sheet.getRange("A" + 決議起始).setFontSize(13).setFontWeight("bold").setBackground("#f5f5f5");

  var 決議標題 = 決議起始 + 1;
  sheet.getRange(決議標題, 1, 1, 6)
    .setValues([["編號", "決議內容", "負責人", "截止日期", "優先級", "狀態"]]);
  sheet.getRange(決議標題, 1, 1, 6)
    .setBackground("#ff6f00").setFontColor("#fff").setFontWeight("bold").setHorizontalAlignment("center");

  // 預留 5 列
  for (var m = 1; m <= 5; m++) {
    sheet.getRange(決議標題 + m, 1).setValue(m).setHorizontalAlignment("center");
    sheet.getRange(決議標題 + m, 6).setValue("⬜ 待處理");
  }

  // --- 欄寬設定 ---
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 150);

  // --- 整體框線 ---
  sheet.getDataRange().setBorder(true, true, true, true, true, true,
    "#bdbdbd", SpreadsheetApp.BorderStyle.SOLID);

  sheet.setFrozenRows(1);
}

/**
 * 根據會議類型回傳預設議程
 * @param {string} 類型 - 會議類型名稱
 * @returns {Array<string>} 預設議程列表
 */
function _取得預設議程(類型) {
  var 議程對照 = {
    "週會": [
      "上週工作進度回報",
      "本週工作計畫",
      "遇到的問題與障礙",
      "跨部門協作事項",
      "臨時動議"
    ],
    "專案會議": [
      "專案進度報告",
      "里程碑檢視",
      "風險與問題追蹤",
      "資源需求討論",
      "下階段計畫",
      "臨時動議"
    ],
    "客戶會議": [
      "客戶需求確認",
      "方案展示與說明",
      "報價與合約討論",
      "時程規劃",
      "後續追蹤事項"
    ],
    "部門會議": [
      "部門績效報告",
      "人力資源與招募",
      "預算執行進度",
      "教育訓練規劃",
      "臨時動議"
    ],
    "臨時會議": [
      "議題說明",
      "討論與決議",
      "後續追蹤"
    ]
  };

  return 議程對照[類型] || 議程對照["臨時會議"];
}

// ============================================================
// 第二部分：巨集 — 快速格式化
// ============================================================

/**
 * 巨集：一鍵標記完成
 * 說明：選取一個決議事項列，快速將狀態改為「已完成」
 */
function 巨集_標記完成() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var 選取範圍 = sheet.getActiveRange();
  var 列 = 選取範圍.getRow();

  // 更新狀態欄（F 欄）
  sheet.getRange(列, 6).setValue("✅ 已完成");
  sheet.getRange(列, 6).setFontColor("#1b5e20").setBackground("#c8e6c9");

  // 加入完成時間備註
  sheet.getRange(列, 6).setNote(
    "完成時間：" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm")
  );

  Logger.log("✅ 已標記第 " + 列 + " 列為已完成");
}

/**
 * 巨集：標記出席
 * 說明：選取出席人員列，快速標記已出席並記錄簽到時間
 */
function 巨集_標記出席() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var 列 = sheet.getActiveRange().getRow();

  sheet.getRange(列, 4).setValue("✅");
  sheet.getRange(列, 5).setValue(
    Utilities.formatDate(new Date(), "Asia/Taipei", "HH:mm")
  );

  Logger.log("✅ 第 " + 列 + " 列已簽到");
}

/**
 * 巨集：插入分隔線
 * 說明：在當前位置插入一條視覺分隔線
 */
function 巨集_插入分隔線() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var 列 = sheet.getActiveRange().getRow();

  sheet.getRange(列, 1, 1, 6).merge();
  sheet.getRange(列, 1).setValue("─────────────────────────────");
  sheet.getRange(列, 1).setHorizontalAlignment("center")
    .setFontColor("#bdbdbd").setBackground("#fafafa");
}

// ============================================================
// 第三部分：觸發器 — 自動化管理
// ============================================================

/**
 * onOpen：建立自訂選單
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🤖 智慧會議管理")
    .addItem("📋 建立會議記錄", "建立會議記錄")
    .addItem("📊 會議統計報表", "產生會議統計")
    .addSeparator()
    .addSubMenu(ui.createMenu("⚡ 快速操作")
      .addItem("✅ 標記完成", "巨集_標記完成")
      .addItem("📝 標記出席", "巨集_標記出席")
      .addItem("➖ 插入分隔線", "巨集_插入分隔線"))
    .addSeparator()
    .addItem("⏰ 設定每週自動提醒", "設定每週觸發器")
    .addItem("🗑️ 清除觸發器", "清除所有觸發器")
    .addToUi();
}

/**
 * 設定每週一早上 8 點的會議提醒觸發器
 */
function 設定每週觸發器() {
  // 清除舊的
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "每週會議提醒") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("每週會議提醒")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();

  SpreadsheetApp.getUi().alert("✅ 每週一 8:00 會議提醒已設定！");
}

/**
 * 每週自動提醒：盤點未完成的決議事項
 */
function 每週會議提醒() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var 未完成項目 = [];

    // 掃描所有會議工作表
    sheets.forEach(function(sheet) {
      var 名稱 = sheet.getName();
      if (名稱.indexOf("會議") === -1 && 名稱.indexOf("📋") === -1 &&
          名稱.indexOf("🚀") === -1 && 名稱.indexOf("🤝") === -1) return;

      var 資料 = sheet.getDataRange().getValues();
      for (var i = 0; i < 資料.length; i++) {
        var 狀態 = String(資料[i][5]);
        if (狀態.indexOf("待處理") >= 0 || 狀態.indexOf("進行中") >= 0) {
          未完成項目.push({
            會議: 名稱,
            內容: 資料[i][1],
            負責人: 資料[i][2],
            截止日: 資料[i][3]
          });
        }
      }
    });

    // 記錄到「待辦追蹤」
    var 追蹤表 = ss.getSheetByName("待辦追蹤");
    if (!追蹤表) {
      追蹤表 = ss.insertSheet("待辦追蹤");
      追蹤表.getRange("A1:E1").setValues([["檢查日期", "來源會議", "待辦內容", "負責人", "截止日"]]);
      追蹤表.getRange("A1:E1").setBackground("#ff6f00").setFontColor("#fff").setFontWeight("bold");
    }

    var 今天 = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd");
    未完成項目.forEach(function(item) {
      var 新列 = 追蹤表.getLastRow() + 1;
      追蹤表.getRange(新列, 1, 1, 5).setValues([
        [今天, item.會議, item.內容, item.負責人, item.截止日]
      ]);
    });

    Logger.log("✅ 每週提醒完成，共 " + 未完成項目.length + " 項未完成");

  } catch (錯誤) {
    Logger.log("❌ 每週提醒錯誤：" + 錯誤.message);
  }
}

/**
 * 清除所有觸發器
 */
function 清除所有觸發器() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) { ScriptApp.deleteTrigger(t); });
  SpreadsheetApp.getUi().alert("✅ 已清除 " + triggers.length + " 個觸發器");
}

// ============================================================
// 第四部分：會議統計報表
// ============================================================

/**
 * 產生會議彙總統計
 */
function 產生會議統計() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var 統計 = { 總會議數: 0, 類型統計: {}, 決議統計: { 已完成: 0, 待處理: 0, 進行中: 0 } };

    sheets.forEach(function(sheet) {
      var 名稱 = sheet.getName();
      // 用圖示辨識會議工作表
      for (var 類型 in 會議設定.會議類型) {
        var 圖示 = 會議設定.會議類型[類型].圖示;
        if (名稱.indexOf(圖示) >= 0 || 名稱.indexOf(類型) >= 0) {
          統計.總會議數++;
          統計.類型統計[類型] = (統計.類型統計[類型] || 0) + 1;

          // 統計決議狀態
          var 資料 = sheet.getDataRange().getValues();
          資料.forEach(function(row) {
            var 狀態 = String(row[5]);
            if (狀態.indexOf("已完成") >= 0) 統計.決議統計.已完成++;
            else if (狀態.indexOf("進行中") >= 0) 統計.決議統計.進行中++;
            else if (狀態.indexOf("待處理") >= 0) 統計.決議統計.待處理++;
          });
          break;
        }
      }
    });

    // 顯示統計
    var 報告 = "📊 會議統計報表\n\n";
    報告 += "📌 總會議數：" + 統計.總會議數 + " 場\n\n";
    報告 += "📋 類型分佈：\n";
    for (var t in 統計.類型統計) {
      報告 += "  " + 會議設定.會議類型[t].圖示 + " " + t + "：" + 統計.類型統計[t] + " 場\n";
    }
    報告 += "\n✅ 決議追蹤：\n";
    報告 += "  已完成：" + 統計.決議統計.已完成 + "\n";
    報告 += "  進行中：" + 統計.決議統計.進行中 + "\n";
    報告 += "  待處理：" + 統計.決議統計.待處理 + "\n";

    SpreadsheetApp.getUi().alert(報告);

  } catch (錯誤) {
    Logger.log("❌ 統計錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 初始化範例資料
// ============================================================

/**
 * 建立範例會議資料（含預填出席人員）
 */
function 初始化範例資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("會議排程");
  if (!sheet) sheet = ss.insertSheet("會議排程"); else sheet.clear();

  var 標題 = [["會議名稱", "類型", "日期", "時間", "主持人", "會議室", "狀態"]];
  var 資料 = [
    ["Q2 業務策略會議", "專案會議", "2026/05/05", "09:00", "林經理", "大會議室 A", "已排程"],
    ["行銷部週會", "週會", "2026/05/05", "14:00", "王主管", "小會議室 B", "已排程"],
    ["客戶 A 需求訪談", "客戶會議", "2026/05/06", "10:00", "張業務", "視訊會議室 C", "已排程"],
    ["HR 部門月會", "部門會議", "2026/05/07", "15:00", "陳人資", "小會議室 B", "已排程"],
    ["系統當機緊急會議", "臨時會議", "2026/05/07", "16:30", "劉工程師", "主管辦公室", "已召開"],
    ["全公司月會", "部門會議", "2026/05/08", "09:00", "總經理", "大會議室 A", "已排程"],
    ["產品上線前審查", "專案會議", "2026/05/09", "10:00", "吳 PM", "大會議室 A", "已排程"],
    ["供應商合約討論", "客戶會議", "2026/05/09", "14:00", "林採購", "小會議室 B", "已排程"]
  ];

  sheet.getRange(1, 1, 1, 7).setValues(標題);
  sheet.getRange(2, 1, 資料.length, 7).setValues(資料);

  // 格式化
  sheet.getRange("A1:G1").setBackground("#1a237e").setFontColor("#fff").setFontWeight("bold")
    .setHorizontalAlignment("center");
  sheet.setFrozenRows(1);
  for (var c = 1; c <= 7; c++) sheet.autoResizeColumn(c);

  // 框線
  sheet.getDataRange().setBorder(true, true, true, true, true, true,
    "#bdbdbd", SpreadsheetApp.BorderStyle.SOLID);

  SpreadsheetApp.getUi().alert("✅ 會議排程範例已建立！\n\n請使用選單「🤖 智慧會議管理」開始操作。");
}
