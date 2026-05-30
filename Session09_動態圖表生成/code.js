// ============================================================
// Session 9：動態圖表生成
// 日期：115/05/30（六）09:00~12:00
// 講師：林冠廷
// ============================================================
// 本課程涵蓋：
//   1. Apps Script 介面複習
//   2. 建立與更新圖表（柱狀圖、折線圖、圓餅圖）
//   3. 觸發器應用（靠近更新 → 圖表更新）
//   4. 實作：動態圖表生成
// ============================================================

// ============================================================
// 第一部分：建立柱狀圖
// ============================================================

/**
 * 建立部門業績柱狀圖
 */
function 建立柱狀圖() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("圖表資料");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("❌ 請先執行「初始化圖表資料」");
      return;
    }

    // 移除該表上既有的圖表（避免重複）
    var 舊圖表 = sheet.getCharts();
    for (var i = 0; i < 舊圖表.length; i++) {
      if (舊圖表[i].getOptions().get("title") === "部門季度業績") {
        sheet.removeChart(舊圖表[i]);
      }
    }

    // 資料範圍：A1:E7（部門 + Q1~Q4）
    var 資料範圍 = sheet.getRange("A1:E7");

    // 建立柱狀圖
    var chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)        // 柱狀圖
      .addRange(資料範圍)                              // 資料來源
      .setNumHeaders(1)                                // 將第一列設為標題（避免標籤顯示為 Column A, B 等）
      .setOption("useFirstColumnAsDomain", true)       // 將第一欄設為 X 軸標籤
      .setPosition(9, 1, 0, 0)                        // 位置：第 9 列、第 1 欄
      .setOption("title", "部門季度業績")              // 標題
      .setOption("titleTextStyle", {
        fontSize: 16,
        bold: true,
        color: "#1a237e"
      })
      .setOption("width", 700)                         // 寬度
      .setOption("height", 400)                        // 高度
      .setOption("legend", { position: "bottom" })     // 圖例位置
      .setOption("hAxis", {
        title: "部門",
        titleTextStyle: { fontSize: 12 }
      })
      .setOption("vAxis", {
        title: "業績 (NT$)",
        titleTextStyle: { fontSize: 12 },
        format: "#,##0"
      })
      .setOption("colors", ["#1a73e8", "#34a853", "#fbbc04", "#ea4335"])  // 自訂顏色
      .setOption("isStacked", false)                   // 非堆疊
      .setOption("series", {                            // 顯示數值資料標籤
        0: { dataLabel: "value" },
        1: { dataLabel: "value" },
        2: { dataLabel: "value" },
        3: { dataLabel: "value" }
      })
      .build();

    sheet.insertChart(chart);

    Logger.log("✅ 柱狀圖已建立！");
    SpreadsheetApp.getUi().alert("✅ 部門季度業績柱狀圖已生成！");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ 錯誤：" + 錯誤.message);
  }
}

/**
 * 建立堆疊柱狀圖
 */
function 建立堆疊柱狀圖() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("圖表資料");
    if (!sheet) return;

    var 資料範圍 = sheet.getRange("A1:E7");

    var chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(資料範圍)
      .setNumHeaders(1)                                // 將第一列設為標題
      .setOption("useFirstColumnAsDomain", true)       // 將第一欄設為 X 軸標籤
      .setPosition(9, 6, 0, 0)
      .setOption("title", "部門季度業績（堆疊）")
      .setOption("titleTextStyle", { fontSize: 14, bold: true })
      .setOption("width", 600)
      .setOption("height", 400)
      .setOption("isStacked", true)                    // ⬅️ 堆疊模式
      .setOption("legend", { position: "bottom" })
      .setOption("colors", ["#1a73e8", "#34a853", "#fbbc04", "#ea4335"])
      .setOption("vAxis", { format: "#,##0" })
      .setOption("series", {                            // 顯示數值資料標籤
        0: { dataLabel: "value" },
        1: { dataLabel: "value" },
        2: { dataLabel: "value" },
        3: { dataLabel: "value" }
      })
      .build();

    sheet.insertChart(chart);
    Logger.log("✅ 堆疊柱狀圖已建立！");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 第二部分：建立折線圖
// ============================================================

/**
 * 建立月度趨勢折線圖
 */
function 建立折線圖() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("月度趨勢");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("❌ 請先執行「初始化圖表資料」");
      return;
    }

    // 清除舊圖表
    sheet.getCharts().forEach(function(c) { sheet.removeChart(c); });

    var 資料範圍 = sheet.getRange("A1:D13"); // 月份 + 3 個系列

    var chart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)             // 折線圖
      .addRange(資料範圍)
      .setNumHeaders(1)                                // 將第一列設為標題
      .setOption("useFirstColumnAsDomain", true)       // 將第一欄設為 X 軸標籤
      .setPosition(15, 1, 0, 0)
      .setOption("title", "2026 年月度營收趨勢")
      .setOption("titleTextStyle", { fontSize: 16, bold: true, color: "#1a237e" })
      .setOption("width", 800)
      .setOption("height", 450)
      .setOption("curveType", "function")              // 平滑曲線
      .setOption("pointSize", 6)                        // 資料點大小
      .setOption("lineWidth", 3)                        // 線條寬度
      .setOption("legend", { position: "bottom" })
      .setOption("hAxis", {
        title: "月份",
        slantedText: true
      })
      .setOption("vAxis", {
        title: "金額 (NT$)",
        format: "#,##0",
        gridlines: { count: 6 }
      })
      .setOption("colors", ["#1a73e8", "#ea4335", "#34a853"])
      .build();

    sheet.insertChart(chart);

    Logger.log("✅ 折線圖已建立！");
    SpreadsheetApp.getUi().alert("✅ 月度趨勢折線圖已生成！");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 第三部分：建立圓餅圖
// ============================================================

/**
 * 建立部門預算佔比圓餅圖
 */
function 建立圓餅圖() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("圖表資料");
    if (!sheet) return;

    // 計算年度總額（G 欄）
    var 資料 = sheet.getRange("A2:E7").getValues();
    var 圓餅資料 = [["部門", "年度總額"]];

    for (var i = 0; i < 資料.length; i++) {
      var 年度總額 = 資料[i][1] + 資料[i][2] + 資料[i][3] + 資料[i][4];
      圓餅資料.push([資料[i][0], 年度總額]);
    }

    // 寫入圓餅圖資料
    sheet.getRange("H1:I1").setValues([["部門", "年度總額"]]);
    sheet.getRange("H1:I1").setFontWeight("bold").setBackground("#e0e0e0");
    for (var j = 0; j < 圓餅資料.length - 1; j++) {
      sheet.getRange(j + 2, 8, 1, 2).setValues([圓餅資料[j + 1]]);
    }
    sheet.getRange("I2:I7").setNumberFormat("#,##0");

    // 建立圓餅圖
    var 圓餅範圍 = sheet.getRange("H1:I7");
    var chart = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)              // 圓餅圖
      .addRange(圓餅範圍)
      .setPosition(25, 1, 0, 0)
      .setOption("title", "部門年度預算佔比")
      .setOption("titleTextStyle", { fontSize: 16, bold: true })
      .setOption("width", 600)
      .setOption("height", 400)
      .setOption("legend", { position: "right" })
      .setOption("pieHole", 0.4)                        // 環形圖效果（0~1）
      .setOption("colors", ["#1a73e8", "#34a853", "#fbbc04", "#ea4335", "#9c27b0", "#00bcd4"])
      .setOption("pieSliceText", "percentage")           // 顯示百分比
      .build();

    sheet.insertChart(chart);

    Logger.log("✅ 圓餅圖已建立！");
    SpreadsheetApp.getUi().alert("✅ 部門預算圓餅圖已生成！");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 第四部分：動態更新圖表
// ============================================================

/**
 * 更新圖表資料
 * 說明：修改資料後，圖表會自動更新（因為圖表綁定了 Range）
 * 這裡示範如何手動觸發資料更新
 */
function 模擬資料更新() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("圖表資料");
    if (!sheet) return;

    // 模擬 Q4 業績更新（隨機增減 10%）
    for (var i = 2; i <= 7; i++) {
      var 原值 = sheet.getRange(i, 5).getValue(); // E 欄 = Q4
      var 變動 = 原值 * (0.9 + Math.random() * 0.2); // ±10%
      sheet.getRange(i, 5).setValue(Math.round(變動));
    }

    // 圖表會自動根據範圍資料更新！
    Logger.log("✅ Q4 業績資料已更新，圖表將自動刷新！");
    SpreadsheetApp.getUi().alert("✅ Q4 資料已更新！\n圖表會自動反映新數據。");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

/**
 * 一鍵生成全部圖表
 */
function 一鍵生成圖表() {
  建立柱狀圖();
  建立堆疊柱狀圖();
  建立折線圖();
  建立圓餅圖();

  SpreadsheetApp.getUi().alert("✅ 全部圖表已生成！\n包含：柱狀圖、堆疊圖、折線圖、圓餅圖");
}

// ============================================================
// 初始化範例資料
// ============================================================

function 初始化圖表資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 圖表資料（部門季度）---
  var sheet1 = ss.getSheetByName("圖表資料");
  if (!sheet1) sheet1 = ss.insertSheet("圖表資料"); else sheet1.clear();

  var 標題1 = [["部門", "Q1", "Q2", "Q3", "Q4"]];
  var 資料1 = [
    ["業務部", 1500000, 1800000, 2000000, 2200000],
    ["行銷部", 800000, 900000, 1000000, 1100000],
    ["研發部", 2500000, 2600000, 2700000, 2800000],
    ["人資部", 500000, 520000, 540000, 560000],
    ["財務部", 400000, 420000, 430000, 450000],
    ["客服部", 600000, 650000, 700000, 750000]
  ];

  sheet1.getRange(1, 1, 1, 5).setValues(標題1);
  sheet1.getRange(2, 1, 資料1.length, 5).setValues(資料1);
  sheet1.getRange("A1:E1").setBackground("#1a237e").setFontColor("#fff").setFontWeight("bold");
  sheet1.getRange("B2:E7").setNumberFormat("#,##0");
  sheet1.setFrozenRows(1);
  for (var c = 1; c <= 5; c++) sheet1.autoResizeColumn(c);

  // --- 月度趨勢資料 ---
  var sheet2 = ss.getSheetByName("月度趨勢");
  if (!sheet2) sheet2 = ss.insertSheet("月度趨勢"); else sheet2.clear();

  var 標題2 = [["月份", "營收", "成本", "利潤"]];
  var 月度資料 = [];
  var 基礎營收 = 3000000;
  for (var m = 1; m <= 12; m++) {
    var 營收 = 基礎營收 + Math.floor(Math.random() * 1000000) + (m * 100000);
    var 成本 = Math.floor(營收 * (0.55 + Math.random() * 0.15));
    var 利潤 = 營收 - 成本;
    月度資料.push([m + "月", 營收, 成本, 利潤]);
  }

  sheet2.getRange(1, 1, 1, 4).setValues(標題2);
  sheet2.getRange(2, 1, 月度資料.length, 4).setValues(月度資料);
  sheet2.getRange("A1:D1").setBackground("#e65100").setFontColor("#fff").setFontWeight("bold");
  sheet2.getRange("B2:D13").setNumberFormat("#,##0");
  sheet2.setFrozenRows(1);
  for (var c2 = 1; c2 <= 4; c2++) sheet2.autoResizeColumn(c2);

  SpreadsheetApp.getUi().alert("✅ 圖表資料已建立！\n（圖表資料 + 月度趨勢）");
}

// ============================================================
// 自訂選單
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📚 Session 9 工具")
    .addItem("📦 初始化圖表資料", "初始化圖表資料")
    .addSeparator()
    .addItem("📊 建立柱狀圖", "建立柱狀圖")
    .addItem("📊 建立堆疊柱狀圖", "建立堆疊柱狀圖")
    .addItem("📈 建立折線圖", "建立折線圖")
    .addItem("🥧 建立圓餅圖", "建立圓餅圖")
    .addSeparator()
    .addItem("🚀 一鍵生成全部圖表", "一鍵生成圖表")
    .addItem("🔄 模擬資料更新", "模擬資料更新")
    .addToUi();
}
