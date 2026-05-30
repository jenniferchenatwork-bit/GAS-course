// ============================================================
// 進階練習：AI 月報簡報自動產生器
// 對應：Session 10（Slides 整合、跨服務操作）
// ============================================================

/**
 * 🚀 一鍵產生 AI 風格月報簡報
 * 整合 Sheets 數據 + 圖表 + Slides 排版
 */
function 產生AI月報簡報() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 開始 = new Date();

    SpreadsheetApp.getUi().alert("⏳ 正在生成 AI 月報簡報...\n這可能需要 15~30 秒。");

    // ===== 讀取資料 =====
    var 業績表 = ss.getSheetByName("月報資料");
    if (!業績表) { SpreadsheetApp.getUi().alert("❌ 請先初始化"); return; }
    var 資料 = 業績表.getDataRange().getValues();

    // 統計
    var 統計 = { 總營收: 0, 總成本: 0, 新客戶: 0, 專案完成: 0 };
    for (var i = 1; i < 資料.length; i++) {
      統計.總營收 += 資料[i][1] || 0;
      統計.總成本 += 資料[i][2] || 0;
      統計.新客戶 += 資料[i][3] || 0;
      統計.專案完成 += 資料[i][4] || 0;
    }
    統計.利潤 = 統計.總營收 - 統計.總成本;
    統計.利潤率 = ((統計.利潤 / 統計.總營收) * 100).toFixed(1);

    // ===== 建立簡報 =====
    var 月份 = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy 年 MM 月");
    var ppt = SlidesApp.create("📊 " + 月份 + " 營運月報");

    // --- P1：封面 ---
    var p1 = ppt.getSlides()[0];
    p1.getBackground().setSolidFill("#0a1628");
    p1.getPlaceholders().forEach(function(p) { p.remove(); });

    var badge = p1.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 260, 30, 200, 35);
    badge.getFill().setSolidFill("#1a73e8");
    badge.getBorder().setTransparent();
    var badgeText = p1.insertTextBox("🤖 AI 自動產生", 270, 33, 180, 28);
    badgeText.getText().getTextStyle().setFontSize(11).setForegroundColor("#ffffff");
    badgeText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    var 標題 = p1.insertTextBox("📊 " + 月份 + "\n營運月報", 40, 120, 640, 120);
    標題.getText().getTextStyle().setFontSize(40).setBold(true).setForegroundColor("#ffffff");
    標題.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    var 副標 = p1.insertTextBox(
      Utilities.formatDate(new Date(), "Asia/Taipei", "報告日期：yyyy/MM/dd") +
      " ｜ " + ss.getName(),
      40, 270, 640, 30
    );
    副標.getText().getTextStyle().setFontSize(12).setForegroundColor("#64b5f6");
    副標.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // --- P2：KPI 摘要 ---
    var p2 = ppt.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    p2.getBackground().setSolidFill("#fafafa");

    var p2Title = p2.insertTextBox("🎯 關鍵營運指標 (KPI)", 40, 15, 640, 40);
    p2Title.getText().getTextStyle().setFontSize(26).setBold(true).setForegroundColor("#0a1628");

    var kpis = [
      { label: "總營收", value: "NT$ " + (統計.總營收 / 10000).toFixed(0) + " 萬", color: "#1a73e8", icon: "💰" },
      { label: "淨利潤", value: "NT$ " + (統計.利潤 / 10000).toFixed(0) + " 萬", color: "#34a853", icon: "📈" },
      { label: "利潤率", value: 統計.利潤率 + "%", color: "#fbbc04", icon: "📊" },
      { label: "新客戶", value: 統計.新客戶 + " 家", color: "#ea4335", icon: "🤝" }
    ];

    kpis.forEach(function(kpi, idx) {
      var x = 30 + idx * 170;
      var card = p2.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, 75, 155, 140);
      card.getFill().setSolidFill(kpi.color);
      card.getBorder().setTransparent();

      var icon = p2.insertTextBox(kpi.icon, x + 10, 85, 135, 35);
      icon.getText().getTextStyle().setFontSize(28);
      icon.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

      var val = p2.insertTextBox(kpi.value, x + 10, 120, 135, 35);
      val.getText().getTextStyle().setFontSize(18).setBold(true).setForegroundColor("#ffffff");
      val.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

      var lbl = p2.insertTextBox(kpi.label, x + 10, 160, 135, 25);
      lbl.getText().getTextStyle().setFontSize(11).setForegroundColor("#e0e0e0");
      lbl.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    });

    // AI 分析重點
    var 分析標題 = p2.insertTextBox("🤖 AI 分析重點", 40, 240, 640, 30);
    分析標題.getText().getTextStyle().setFontSize(16).setBold(true).setForegroundColor("#0a1628");

    var 利潤評語 = Number(統計.利潤率) >= 30 ? "利潤率優秀，維持策略有效" :
                   Number(統計.利潤率) >= 20 ? "利潤率良好，可優化成本" : "利潤率偏低，建議檢視成本結構";
    var 客戶評語 = 統計.新客戶 >= 20 ? "客戶拓展快速，注意服務品質" :
                   統計.新客戶 >= 10 ? "客戶成長穩定" : "客戶拓展緩慢，建議加強行銷";

    var 分析 = p2.insertTextBox(
      "📌 " + 利潤評語 + "\n📌 " + 客戶評語 + "\n📌 本月完成 " + 統計.專案完成 + " 個專案",
      60, 275, 600, 100
    );
    分析.getText().getTextStyle().setFontSize(13).setForegroundColor("#37474f");

    // --- P3：圖表頁 ---
    var 圖表列表 = 業績表.getCharts();
    if (圖表列表.length > 0) {
      var p3 = ppt.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      p3.getBackground().setSolidFill("#ffffff");
      var ct = p3.insertTextBox("📈 營運趨勢圖表", 40, 10, 640, 35);
      ct.getText().getTextStyle().setFontSize(22).setBold(true).setForegroundColor("#0a1628");

      var blob = 圖表列表[0].getBlob();
      var img = p3.insertImage(blob);
      img.setLeft(40).setTop(55).setWidth(640).setHeight(350);
    }

    // --- P4：部門表格 ---
    var p4 = ppt.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    p4.getBackground().setSolidFill("#fafafa");
    var ttl = p4.insertTextBox("🏢 部門業績明細", 40, 10, 640, 35);
    ttl.getText().getTextStyle().setFontSize(22).setBold(true).setForegroundColor("#0a1628");

    var rows = Math.min(資料.length, 8);
    var cols = Math.min(資料[0].length, 5);
    var table = p4.insertTable(rows, cols);
    for (var ri = 0; ri < rows; ri++) {
      for (var ci = 0; ci < cols; ci++) {
        var v = 資料[ri][ci];
        if (typeof v === "number" && ri > 0) v = "NT$ " + v.toLocaleString();
        table.getCell(ri, ci).getText().setText(String(v));
        table.getCell(ri, ci).getText().getTextStyle().setFontSize(10);
        if (ri === 0) {
          table.getCell(ri, ci).getFill().setSolidFill("#0a1628");
          table.getCell(ri, ci).getText().getTextStyle().setForegroundColor("#ffffff").setBold(true);
        } else if (ri % 2 === 0) {
          table.getCell(ri, ci).getFill().setSolidFill("#e3f2fd");
        }
      }
    }

    // --- 結尾 ---
    var pEnd = ppt.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    pEnd.getBackground().setSolidFill("#0a1628");
    var end = pEnd.insertTextBox("謝謝聆聽\n\n🤖 本報告由 AI 自動產生", 40, 140, 640, 120);
    end.getText().getTextStyle().setFontSize(32).setBold(true).setForegroundColor("#ffffff");
    end.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // 儲存並關閉簡報以確保所有變更都寫入
    ppt.saveAndClose();

    // ===== 轉成 PDF 並寄送 Email =====
    var pptFile = DriveApp.getFileById(ppt.getId());
    var pdfBlob = pptFile.getAs(MimeType.PDF).setName(ppt.getName() + ".pdf");
    
    var 主管Email = "jenniferchenatwork@gmail.com";
    MailApp.sendEmail({
      to: 主管Email,
      subject: "📊 " + 月份 + " 營運月報簡報 (AI 自動生成)",
      body: "主管您好：\n\n本月的 AI 營運月報簡報已自動生成，請參閱附件 PDF 檔案，或是點擊下方連結直接線上瀏覽簡報。\n\n🔗 簡報線上瀏覽連結：\n" + ppt.getUrl() + "\n\n祝 順心\nAI 自動排版系統",
      attachments: [pdfBlob]
    });

    var 耗時 = Math.round((new Date() - 開始) / 1000);
    SpreadsheetApp.getUi().alert(
      "🎉 AI 月報簡報已完成！\n\n" +
      "📄 共 " + ppt.getSlides().length + " 頁 ｜ ⏱️ " + 耗時 + " 秒\n" +
      "📧 PDF 簡報已成功寄送至主管信箱：" + 主管Email + "\n\n" +
      "🔗 線上簡報連結：" + ppt.getUrl()
    );

  } catch (錯誤) {
    Logger.log("❌ " + 錯誤.message + "\n" + 錯誤.stack);
    SpreadsheetApp.getUi().alert("❌ " + 錯誤.message);
  }
}

function 初始化月報資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("月報資料");
  if (!sheet) sheet = ss.insertSheet("月報資料"); else sheet.clear();

  var 標題 = [["部門", "營收", "成本", "新客戶", "完成專案"]];
  var 資料 = [
    ["業務部", 4800000, 2400000, 12, 8],
    ["行銷部", 2100000, 1200000, 8, 5],
    ["研發部", 0, 3500000, 0, 12],
    ["客服部", 950000, 600000, 5, 15],
    ["人資部", 0, 780000, 0, 3],
    ["財務部", 0, 420000, 0, 6]
  ];

  sheet.getRange(1, 1, 1, 5).setValues(標題);
  sheet.getRange(2, 1, 資料.length, 5).setValues(資料);
  sheet.getRange("A1:E1").setBackground("#0a1628").setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange("B2:C7").setNumberFormat("#,##0");
  sheet.setFrozenRows(1);
  for (var c = 1; c <= 5; c++) sheet.autoResizeColumn(c);

  // 建立簡單柱狀圖
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange("A1:C7"))
    .setPosition(9, 1, 0, 0)
    .setOption("title", "部門營收 vs 成本")
    .setOption("width", 600).setOption("height", 350)
    .setOption("colors", ["#1a73e8", "#ea4335"])
    .setOption("vAxis", { format: "#,##0" })
    .build();
  sheet.insertChart(chart);

  SpreadsheetApp.getUi().alert("✅ 月報資料已建立！請執行「產生 AI 月報簡報」。");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🤖 AI 月報產生器")
    .addItem("📦 初始化月報資料", "初始化月報資料")
    .addItem("🚀 產生 AI 月報簡報", "產生AI月報簡報")
    .addToUi();
}
