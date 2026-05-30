// ============================================================
// 進階練習：一鍵生成客戶報價單
// 對應：Session 7（格式設定、文字/數字/日期格式）
// ============================================================

/**
 * 一鍵生成專業報價單
 */
function 生成報價單() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 客戶表 = ss.getSheetByName("報價資料");
    if (!客戶表) { SpreadsheetApp.getUi().alert("❌ 請先初始化"); return; }

    var ui = SpreadsheetApp.getUi();
    var 客戶回應 = ui.prompt("📝 報價單", "請輸入客戶名稱：", ui.ButtonSet.OK_CANCEL);
    if (客戶回應.getSelectedButton() !== ui.Button.OK) return;
    var 客戶 = 客戶回應.getResponseText().trim() || "範例客戶";

    var 編號 = "QT-" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd") + "-" +
               String(Math.floor(Math.random() * 100)).padStart(3, "0");

    var sheet = ss.insertSheet(編號, 0);

    // ===== 報價單格式 =====

    // 公司 Logo 區
    sheet.getRange("A1:F1").merge();
    sheet.getRange("A1").setValue("ABC 科技股份有限公司")
      .setFontSize(20).setFontWeight("bold").setFontColor("#1a237e");
    sheet.setRowHeight(1, 50);

    sheet.getRange("A2:F2").merge();
    sheet.getRange("A2").setValue("台北市信義區信義路五段 7 號 ｜ Tel: 02-2345-6789 ｜ www.abc-tech.com")
      .setFontSize(9).setFontColor("#666");

    // 報價單標題
    sheet.getRange("A4:F4").merge();
    sheet.getRange("A4").setValue("📋 報 價 單")
      .setFontSize(24).setFontWeight("bold").setHorizontalAlignment("center")
      .setBackground("#1a237e").setFontColor("#fff");
    sheet.setRowHeight(4, 45);

    // 客戶資訊
    var 今天 = new Date();
    var 有效日 = new Date(今天); 有效日.setDate(今天.getDate() + 30);

    sheet.getRange(6, 1, 4, 6).setValues([
      ["報價編號", 編號, "", "客戶名稱", 客戶, ""],
      ["報價日期", Utilities.formatDate(今天, "Asia/Taipei", "yyyy/MM/dd"), "", "聯絡人", "", ""],
      ["有效期限", Utilities.formatDate(有效日, "Asia/Taipei", "yyyy/MM/dd"), "", "聯絡電話", "", ""],
      ["業務人員", "Elena Wang", "", "傳真", "", ""]
    ]);

    for (var r = 6; r <= 9; r++) {
      sheet.getRange(r, 1).setFontWeight("bold").setBackground("#e8eaf6");
      sheet.getRange(r, 4).setFontWeight("bold").setBackground("#e8eaf6");
    }

    // 品項表格
    sheet.getRange("A11:F11").setValues([["項次", "品名/規格", "單位", "數量", "單價", "金額"]]);
    sheet.getRange("A11:F11").setBackground("#283593").setFontColor("#fff")
      .setFontWeight("bold").setHorizontalAlignment("center");
    sheet.setRowHeight(11, 35);

    // 讀取報價品項
    var 品項資料 = 客戶表.getDataRange().getValues();
    var 小計 = 0;

    for (var i = 1; i < 品項資料.length; i++) {
      var 列 = 11 + i;
      var 金額 = 品項資料[i][3] * 品項資料[i][4]; // 數量 × 單價
      小計 += 金額;

      sheet.getRange(列, 1, 1, 6).setValues([[
        i, 品項資料[i][0], 品項資料[i][1], 品項資料[i][3], 品項資料[i][4], 金額
      ]]);
      sheet.getRange(列, 1).setHorizontalAlignment("center");
      sheet.getRange(列, 4, 1, 3).setNumberFormat("#,##0");
      sheet.getRange(列, 4, 1, 3).setHorizontalAlignment("right");

      // 斑馬紋
      if (i % 2 === 0) sheet.getRange(列, 1, 1, 6).setBackground("#f5f5f5");
    }

    // 合計區
    var 稅金 = Math.round(小計 * 0.05);
    var 總計 = 小計 + 稅金;
    var 合計起始 = 12 + 品項資料.length - 1;

    var 合計資料 = [
      ["", "", "", "", "小　計", 小計],
      ["", "", "", "", "稅金 (5%)", 稅金],
      ["", "", "", "", "總　計", 總計]
    ];

    sheet.getRange(合計起始, 1, 3, 6).setValues(合計資料);
    sheet.getRange(合計起始, 5, 3, 1).setFontWeight("bold").setHorizontalAlignment("right");
    sheet.getRange(合計起始, 6, 3, 1).setNumberFormat("NT$#,##0").setFontWeight("bold");
    sheet.getRange(合計起始 + 2, 5, 1, 2).setBackground("#e8eaf6")
      .setFontSize(14).setBorder(true, true, true, true, false, false,
        "#1a237e", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // 備註
    var 備註列 = 合計起始 + 5;
    sheet.getRange(備註列, 1, 1, 6).merge();
    sheet.getRange(備註列, 1).setValue("📌 備註：").setFontWeight("bold");
    sheet.getRange(備註列 + 1, 1, 1, 6).merge();
    sheet.getRange(備註列 + 1, 1).setValue(
      "1. 報價有效期限 30 天\n2. 付款條件：月結 30 天\n3. 交貨期：訂購後 7~14 個工作天\n4. 以上報價含安裝及教育訓練"
    ).setWrap(true).setFontColor("#555");

    // 欄寬
    sheet.setColumnWidth(1, 50);
    sheet.setColumnWidth(2, 250);
    sheet.setColumnWidth(3, 60);
    sheet.setColumnWidth(4, 80);
    sheet.setColumnWidth(5, 100);
    sheet.setColumnWidth(6, 120);

    // 整體框線
    sheet.getRange(11, 1, 品項資料.length + 3, 6)
      .setBorder(true, true, true, true, true, true, "#bdbdbd", SpreadsheetApp.BorderStyle.SOLID);

    sheet.setFrozenRows(0);

    SpreadsheetApp.getUi().alert("✅ 報價單 " + 編號 + " 已生成！\n總金額：NT$ " + 總計.toLocaleString());

  } catch (錯誤) { Logger.log("❌ " + 錯誤.message); SpreadsheetApp.getUi().alert("❌ " + 錯誤.message); }
}

function 初始化報價資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("報價資料");
  if (!sheet) sheet = ss.insertSheet("報價資料"); else sheet.clear();

  var 標題 = [["品名/規格", "單位", "說明", "數量", "單價"]];
  var 資料 = [
    ["AI 智慧會議系統（基本版）", "套", "含語音辨識、自動會議紀要", 1, 180000],
    ["智慧文件管理模組", "授權", "OCR + AI 分類，10 人授權", 10, 12000],
    ["自動化報表工具", "授權", "每日/週/月報表自動產生", 5, 8500],
    ["教育訓練", "小時", "現場教育訓練（含教材）", 16, 3000],
    ["系統維護（年約）", "年", "含系統更新與技術支援", 1, 50000],
    ["客製化開發", "人天", "依需求客製功能開發", 10, 8000]
  ];

  sheet.getRange(1, 1, 1, 5).setValues(標題);
  sheet.getRange(2, 1, 資料.length, 5).setValues(資料);
  sheet.getRange("A1:E1").setBackground("#283593").setFontColor("#fff").setFontWeight("bold");
  sheet.getRange("E2:E7").setNumberFormat("#,##0");
  sheet.setFrozenRows(1);
  for (var c = 1; c <= 5; c++) sheet.autoResizeColumn(c);

  SpreadsheetApp.getUi().alert("✅ 報價資料已建立！請執行「生成報價單」。");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🤖 報價單系統")
    .addItem("📦 初始化報價資料", "初始化報價資料")
    .addItem("📋 生成報價單", "生成報價單")
    .addToUi();
}
