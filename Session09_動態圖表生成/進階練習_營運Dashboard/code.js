// ============================================================
// 進階練習：自動化營運 Dashboard
// 對應：Session 9（柱狀圖、折線圖、圓餅圖）
// ============================================================

/**
 * 一鍵建立完整營運儀表板（多種圖表組合）
 */
function 建立營運Dashboard() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 資料表 = ss.getSheetByName("營運資料");
    if (!資料表) { SpreadsheetApp.getUi().alert("❌ 請先初始化"); return; }

    var sheet = ss.getSheetByName("Dashboard");
    if (sheet) { sheet.getCharts().forEach(function(c) { sheet.removeChart(c); }); sheet.clear(); }
    else sheet = ss.insertSheet("Dashboard", 0);

    // 標題
    sheet.getRange("A1:L1").merge();
    sheet.getRange("A1").setValue("📊 2026 年營運 Dashboard")
      .setFontSize(22).setFontWeight("bold").setHorizontalAlignment("center")
      .setBackground("#0d47a1").setFontColor("#fff");
    sheet.setRowHeight(1, 55);

    // 讀取資料
    var 資料 = 資料表.getDataRange().getValues();

    // --- 圖表 1：月度營收折線圖 ---
    var 月度範圍 = 資料表.getRange("A1:D13");
    var 折線圖 = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(月度範圍)
      .setPosition(3, 1, 0, 0)
      .setOption("title", "📈 月度營收 vs 成本 vs 利潤趨勢")
      .setOption("titleTextStyle", { fontSize: 14, bold: true })
      .setOption("width", 600).setOption("height", 350)
      .setOption("curveType", "function")
      .setOption("pointSize", 5).setOption("lineWidth", 3)
      .setOption("colors", ["#1a73e8", "#ea4335", "#34a853"])
      .setOption("legend", { position: "bottom" })
      .setOption("vAxis", { format: "#,##0" })
      .build();
    sheet.insertChart(折線圖);

    // --- 圖表 2：部門營收柱狀圖 ---
    var 部門表 = ss.getSheetByName("部門績效");
    if (部門表) {
      var 部門範圍 = 部門表.getRange("A1:C7");
      var 柱狀圖 = sheet.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(部門範圍)
        .setPosition(3, 7, 0, 0)
        .setOption("title", "📊 部門目標 vs 實績")
        .setOption("titleTextStyle", { fontSize: 14, bold: true })
        .setOption("width", 500).setOption("height", 350)
        .setOption("colors", ["#90caf9", "#1565c0"])
        .setOption("vAxis", { format: "#,##0" })
        .build();
      sheet.insertChart(柱狀圖);
    }

    // --- 圖表 3：客戶產業分佈圓餅圖 ---
    var 產業表 = ss.getSheetByName("客戶產業");
    if (產業表) {
      var 產業範圍 = 產業表.getRange("A1:B6");
      var 圓餅圖 = sheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(產業範圍)
        .setPosition(23, 1, 0, 0)
        .setOption("title", "🥧 客戶產業分佈")
        .setOption("titleTextStyle", { fontSize: 14, bold: true })
        .setOption("width", 500).setOption("height", 350)
        .setOption("pieHole", 0.4)
        .setOption("colors", ["#1a73e8", "#34a853", "#fbbc04", "#ea4335", "#9c27b0"])
        .setOption("pieSliceText", "percentage")
        .build();
      sheet.insertChart(圓餅圖);
    }

    // --- 圖表 4：月度客戶數量趨勢 ---
    var 客戶趨勢 = ss.getSheetByName("客戶趨勢");
    if (客戶趨勢) {
      var 趨勢範圍 = 客戶趨勢.getRange("A1:C13");
      var 面積圖 = sheet.newChart()
        .setChartType(Charts.ChartType.AREA)
        .addRange(趨勢範圍)
        .setPosition(23, 7, 0, 0)
        .setOption("title", "📈 客戶數量趨勢")
        .setOption("titleTextStyle", { fontSize: 14, bold: true })
        .setOption("width", 500).setOption("height", 350)
        .setOption("colors", ["#1a73e8", "#34a853"])
        .setOption("isStacked", true)
        .build();
      sheet.insertChart(面積圖);
    }

    SpreadsheetApp.getUi().alert("✅ 營運 Dashboard 已建立！共 4 個圖表。");

  } catch (錯誤) { Logger.log("❌ " + 錯誤.message); SpreadsheetApp.getUi().alert("❌ " + 錯誤.message); }
}

function 初始化營運資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 營運資料（月度）
  var s1 = ss.getSheetByName("營運資料");
  if (!s1) s1 = ss.insertSheet("營運資料"); else s1.clear();
  s1.getRange(1, 1, 1, 4).setValues([["月份", "營收", "成本", "利潤"]]);
  var 月度 = [];
  for (var m = 1; m <= 12; m++) {
    var 營收 = 3000000 + m * 200000 + Math.floor(Math.random() * 500000);
    var 成本 = Math.floor(營收 * (0.55 + Math.random() * 0.1));
    月度.push([m + "月", 營收, 成本, 營收 - 成本]);
  }
  s1.getRange(2, 1, 12, 4).setValues(月度);
  s1.getRange("A1:D1").setBackground("#1565c0").setFontColor("#fff").setFontWeight("bold");
  s1.getRange("B2:D13").setNumberFormat("#,##0");

  // 部門績效
  var s2 = ss.getSheetByName("部門績效");
  if (!s2) s2 = ss.insertSheet("部門績效"); else s2.clear();
  s2.getRange(1, 1, 1, 3).setValues([["部門", "目標", "實績"]]);
  s2.getRange(2, 1, 6, 3).setValues([
    ["業務部", 5000000, 4800000], ["行銷部", 2000000, 2100000],
    ["研發部", 3000000, 2900000], ["客服部", 1000000, 950000],
    ["人資部", 800000, 780000], ["財務部", 600000, 620000]
  ]);
  s2.getRange("A1:C1").setBackground("#2e7d32").setFontColor("#fff").setFontWeight("bold");

  // 客戶產業
  var s3 = ss.getSheetByName("客戶產業");
  if (!s3) s3 = ss.insertSheet("客戶產業"); else s3.clear();
  s3.getRange(1, 1, 1, 2).setValues([["產業", "客戶數"]]);
  s3.getRange(2, 1, 5, 2).setValues([
    ["科技業", 45], ["製造業", 30], ["服務業", 25], ["金融業", 18], ["其他", 12]
  ]);
  s3.getRange("A1:B1").setBackground("#6a1b9a").setFontColor("#fff").setFontWeight("bold");

  // 客戶趨勢
  var s4 = ss.getSheetByName("客戶趨勢");
  if (!s4) s4 = ss.insertSheet("客戶趨勢"); else s4.clear();
  s4.getRange(1, 1, 1, 3).setValues([["月份", "新客戶", "續約客戶"]]);
  var 趨勢 = [];
  for (var n = 1; n <= 12; n++) {
    趨勢.push([n + "月", Math.floor(Math.random() * 10) + 5, Math.floor(Math.random() * 15) + 20]);
  }
  s4.getRange(2, 1, 12, 3).setValues(趨勢);
  s4.getRange("A1:C1").setBackground("#e65100").setFontColor("#fff").setFontWeight("bold");

  SpreadsheetApp.getUi().alert("✅ 營運資料已建立！（4 個資料表）\n請執行「建立營運 Dashboard」。");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🤖 營運 Dashboard")
    .addItem("📦 初始化營運資料", "初始化營運資料")
    .addItem("📊 建立營運 Dashboard", "建立營運Dashboard")
    .addToUi();
}
