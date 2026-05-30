// ============================================================
// Session 10：Google Slides 整合專案
// 日期：115/05/30（六）13:30~16:30
// 講師：林冠廷
// ============================================================
// 本課程涵蓋：
//   1. Google Slides 整合（簡報建立、插入圖表與文字）
//   2. 執行限制（跨服務操作限制）
//   3. 專案實作：一鍵產生含圖表的報表簡報
// ============================================================

// ============================================================
// 第一部分：Google Slides 基本操作
// ============================================================

/**
 * 建立新的 Google Slides 簡報ss
 * 說明：示範如何用程式碼建立和操作簡報
 */
function 建立基本簡報() {
  try {
    // 建立新簡報
    var presentation = SlidesApp.create("GAS 自動產生簡報 - " +
      Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd"));

    Logger.log("簡報已建立！");
    Logger.log("簡報 ID：" + presentation.getId());
    Logger.log("簡報 URL：" + presentation.getUrl());

    // --- 編輯第一張投影片（標題頁）---
    var 標題頁 = presentation.getSlides()[0];

    // 取得標題佔位符
    var 標題 = 標題頁.getPlaceholder(SlidesApp.PlaceholderType.CENTERED_TITLE);
    if (標題) {
      標題.asShape().getText().setText("📊 2026 年度營運報告");
      var 文字樣式 = 標題.asShape().getText().getTextStyle();
      文字樣式.setFontSize(36);
      文字樣式.setBold(true);
      文字樣式.setForegroundColor("#1a237e");
    }

    // 取得副標題佔位符
    var 副標題 = 標題頁.getPlaceholder(SlidesApp.PlaceholderType.SUBTITLE);
    if (副標題) {
      副標題.asShape().getText().setText(
        "自動產生日期：" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy 年 MM 月 dd 日") +
        "\n由 Google Apps Script 自動生成"
      );
    }

    // --- 設定背景色 ---
    標題頁.getBackground().setSolidFill("#e8eaf6");

    Logger.log("✅ 標題頁已設定");

    // 開啟簡報
    SpreadsheetApp.getUi().alert(
      "✅ 簡報已建立！\n\n" +
      "點選連結開啟：\n" + presentation.getUrl()
    );

    return presentation;

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 第二部分：在簡報中插入內容
// ============================================================

/**
 * 建立含有多頁內容的簡報
 */
function 建立多頁簡報() {
  try {
    var presentation = SlidesApp.create("多頁簡報範例 - " +
      Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd_HHmm"));

    // === 第 1 頁：標題頁 ===
    var 第1頁 = presentation.getSlides()[0];
    第1頁.getBackground().setSolidFill("#1a237e");

    // 清除預設佈局，改用自訂元素
    第1頁.getPlaceholders().forEach(function (p) { p.remove(); });

    // 加入標題文字方塊
    var 標題框 = 第1頁.insertTextBox("📊 季度營運分析報告", 50, 150, 620, 80);
    var 標題樣式 = 標題框.getText().getTextStyle();
    標題樣式.setFontSize(32).setBold(true).setForegroundColor("#ffffff");
    標題框.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // 加入日期
    var 日期框 = 第1頁.insertTextBox(
      Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy 年 MM 月 dd 日"),
      50, 250, 620, 40
    );
    日期框.getText().getTextStyle().setFontSize(18).setForegroundColor("#b0bec5");
    日期框.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // === 第 2 頁：摘要數據 ===
    var 第2頁 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    第2頁.getBackground().setSolidFill("#ffffff");

    // 頁面標題
    var 頁標題 = 第2頁.insertTextBox("📈 關鍵績效指標", 40, 20, 640, 50);
    頁標題.getText().getTextStyle().setFontSize(28).setBold(true).setForegroundColor("#1a237e");

    // 從試算表讀取資料
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("圖表資料");

    if (sheet) {
      var 資料 = sheet.getDataRange().getValues();

      // 建立表格
      var 表格 = 第2頁.insertTable(資料.length, 資料[0].length > 5 ? 5 : 資料[0].length);

      for (var i = 0; i < Math.min(資料.length, 7); i++) {
        for (var j = 0; j < Math.min(資料[0].length, 5); j++) {
          var 儲存格 = 表格.getCell(i, j);
          var 值 = 資料[i][j];

          // 格式化數字
          if (typeof 值 === "number" && i > 0) {
            值 = "NT$ " + 值.toLocaleString();
          }

          儲存格.getText().setText(String(值));
          儲存格.getText().getTextStyle().setFontSize(11);

          // 標題列格式
          if (i === 0) {
            儲存格.getFill().setSolidFill("#1a237e");
            儲存格.getText().getTextStyle()
              .setFontSize(12)
              .setBold(true)
              .setForegroundColor("#ffffff");
          } else if (i % 2 === 0) {
            儲存格.getFill().setSolidFill("#e8eaf6");
          }
        }
      }
    }

    // === 第 3 頁：重點摘要 ===
    var 第3頁 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    第3頁.getBackground().setSolidFill("#f5f5f5");

    var 摘要標題 = 第3頁.insertTextBox("💡 本季重點摘要", 40, 20, 640, 50);
    摘要標題.getText().getTextStyle().setFontSize(28).setBold(true).setForegroundColor("#1a237e");

    var 摘要內容 = 第3頁.insertTextBox(
      "✅ 研發部持續領先，Q4 預算達 280 萬\n\n" +
      "📈 業務部穩定成長，年增 46.7%\n\n" +
      "⚠️ 財務部預算最低，需評估人力配置\n\n" +
      "🎯 下季目標：整體營收成長 15%",
      60, 90, 600, 300
    );
    摘要內容.getText().getTextStyle().setFontSize(18).setForegroundColor("#37474f");
    摘要內容.getText().getListStyle().applyListPreset(SlidesApp.ListPreset.DISC_CIRCLE_SQUARE);

    // === 第 4 頁：結尾頁 ===
    var 第4頁 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    第4頁.getBackground().setSolidFill("#1a237e");

    var 結尾框 = 第4頁.insertTextBox("謝謝聆聽", 50, 180, 620, 60);
    結尾框.getText().getTextStyle().setFontSize(40).setBold(true).setForegroundColor("#ffffff");
    結尾框.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    var 聯絡框 = 第4頁.insertTextBox("📧 contact@company.com", 50, 260, 620, 40);
    聯絡框.getText().getTextStyle().setFontSize(16).setForegroundColor("#90caf9");
    聯絡框.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    Logger.log("✅ 多頁簡報已建立！共 " + presentation.getSlides().length + " 頁");
    SpreadsheetApp.getUi().alert(
      "✅ 多頁簡報已建立！\n\n" +
      "共 " + presentation.getSlides().length + " 頁\n" +
      "點選連結開啟：\n" + presentation.getUrl()
    );

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 第三部分：圖表插入簡報
// ============================================================

/**
 * 將試算表中的圖表插入到簡報
 */
function 插入圖表到簡報() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("圖表資料");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("❌ 請先執行「初始化圖表資料」並建立圖表");
      return;
    }

    // 取得所有圖表
    var 圖表列表 = sheet.getCharts();
    if (圖表列表.length === 0) {
      SpreadsheetApp.getUi().alert("❌ 找不到圖表！\n請先在 Session 9 中建立圖表。");
      return;
    }

    // 建立簡報
    var presentation = SlidesApp.create("含圖表簡報 - " +
      Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd_HHmm"));

    // 標題頁
    var 標題頁 = presentation.getSlides()[0];
    標題頁.getBackground().setSolidFill("#e8eaf6");
    var placeholders = 標題頁.getPlaceholders();
    if (placeholders.length > 0) {
      placeholders[0].asShape().getText().setText("📊 圖表分析報告");
      placeholders[0].asShape().getText().getTextStyle()
        .setFontSize(32).setBold(true).setForegroundColor("#1a237e");
    }

    // 為每個圖表建立一頁
    for (var i = 0; i < 圖表列表.length; i++) {
      var 新頁 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

      // 從試算表圖表取得圖片
      var 圖表圖片 = 圖表列表[i].getBlob();

      // 插入圖表圖片到簡報
      var image = 新頁.insertImage(圖表圖片);

      // 調整位置和大小
      image.setLeft(40);
      image.setTop(60);
      image.setWidth(640);
      image.setHeight(380);

      // 加入頁面標題
      var 圖表標題 = 圖表列表[i].getOptions().get("title") || "圖表 " + (i + 1);
      var 標題框 = 新頁.insertTextBox("📊 " + 圖表標題, 40, 10, 640, 45);
      標題框.getText().getTextStyle().setFontSize(22).setBold(true).setForegroundColor("#1a237e");
    }

    Logger.log("✅ 含圖表簡報已建立！共插入 " + 圖表列表.length + " 個圖表");
    SpreadsheetApp.getUi().alert(
      "✅ 含圖表簡報已建立！\n\n" +
      "共 " + (圖表列表.length + 1) + " 頁（包含 " + 圖表列表.length + " 個圖表）\n\n" +
      "點選連結開啟：\n" + presentation.getUrl()
    );

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 第四部分：🔧 專案實作 — 一鍵產生含圖表的報表簡報
// ============================================================

/**
 * 🚀 一鍵產生完整報表簡報
 * 說明：整合所有學過的技能
 *   1. 讀取試算表資料
 *   2. 計算統計資訊
 *   3. 建立圖表
 *   4. 產生簡報（含表格、圖表、摘要）
 */
function 一鍵產生報表簡報() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 開始時間 = new Date();

    SpreadsheetApp.getUi().alert("⏳ 開始生成報表簡報...\n\n這可能需要 10~30 秒，請稍候。");

    // ========== Step 1：讀取資料 ==========
    Logger.log("Step 1：讀取資料...");

    var 部門表 = ss.getSheetByName("圖表資料");
    var 趨勢表 = ss.getSheetByName("月度趨勢");

    if (!部門表 || !趨勢表) {
      SpreadsheetApp.getUi().alert("❌ 請先執行「初始化圖表資料」！");
      return;
    }

    var 部門資料 = 部門表.getDataRange().getValues();
    var 趨勢資料 = 趨勢表.getDataRange().getValues();

    // ========== Step 2：統計計算 ==========
    Logger.log("Step 2：計算統計...");

    var 部門統計 = [];
    var 年度總營收 = 0;
    for (var i = 1; i < 部門資料.length; i++) {
      var 年度合計 = 部門資料[i][1] + 部門資料[i][2] + 部門資料[i][3] + 部門資料[i][4];
      年度總營收 += 年度合計;
      部門統計.push({
        部門: 部門資料[i][0],
        Q1: 部門資料[i][1],
        Q4: 部門資料[i][4],
        年度: 年度合計,
        成長率: ((部門資料[i][4] - 部門資料[i][1]) / 部門資料[i][1] * 100).toFixed(1)
      });
    }

    // 排序：年度業績由高到低
    部門統計.sort(function (a, b) { return b.年度 - a.年度; });

    var 月度統計 = [];
    var 最高月營收 = 0, 最高月 = "";
    for (var j = 1; j < 趨勢資料.length; j++) {
      var 營收 = 趨勢資料[j][1];
      if (營收 > 最高月營收) {
        最高月營收 = 營收;
        最高月 = 趨勢資料[j][0];
      }
    }

    // ========== Step 3：確保圖表存在 ==========
    Logger.log("Step 3：檢查圖表...");

    var 圖表列表 = 部門表.getCharts();
    var 趨勢圖表列表 = 趨勢表.getCharts();

    // ========== Step 4：建立簡報 ==========
    Logger.log("Step 4：建立簡報...");

    var 簡報名 = "📊 營運分析報告 - " +
      Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy.MM.dd");
    var ppt = SlidesApp.create(簡報名);

    // --- 第 1 頁：封面 ---
    var 封面 = ppt.getSlides()[0];
    封面.getBackground().setSolidFill("#0d47a1");
    封面.getPlaceholders().forEach(function (p) { p.remove(); });

    var t1 = 封面.insertTextBox("📊 2026 年度營運分析報告", 40, 130, 640, 80);
    t1.getText().getTextStyle().setFontSize(36).setBold(true).setForegroundColor("#ffffff");
    t1.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    var t2 = 封面.insertTextBox(
      Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy 年 MM 月 dd 日") +
      " ｜ 自動產生報告",
      40, 230, 640, 40
    );
    t2.getText().getTextStyle().setFontSize(16).setForegroundColor("#90caf9");
    t2.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // 加入裝飾線
    var 線 = 封面.insertLine(
      SlidesApp.LineCategory.STRAIGHT, 200, 220, 520, 220
    );
    線.getLineFill().setSolidFill("#64b5f6");
    線.setWeight(2);

    // --- 第 2 頁：KPI 摘要 ---
    var KPI頁 = ppt.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    KPI頁.getBackground().setSolidFill("#fafafa");

    var kpiTitle = KPI頁.insertTextBox("🎯 關鍵績效指標 (KPI)", 40, 15, 640, 45);
    kpiTitle.getText().getTextStyle().setFontSize(26).setBold(true).setForegroundColor("#1a237e");

    // KPI 卡片
    var kpiData = [
      { 標籤: "年度總營收", 值: "NT$ " + (年度總營收 / 10000).toFixed(0) + " 萬", 色: "#1a73e8" },
      { 標籤: "部門數", 值: 部門統計.length + " 個", 色: "#34a853" },
      { 標籤: "最高月營收", 值: 最高月, 色: "#fbbc04" },
      { 標籤: "營收冠軍", 值: 部門統計[0].部門, 色: "#ea4335" }
    ];

    for (var k = 0; k < kpiData.length; k++) {
      var x = 40 + (k * 165);
      // 卡片背景
      var card = KPI頁.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, 80, 150, 120);
      card.getFill().setSolidFill(kpiData[k].色);
      card.getBorder().setWeight(0);

      // 卡片文字
      var 值框 = KPI頁.insertTextBox(kpiData[k].值, x + 10, 100, 130, 40);
      值框.getText().getTextStyle().setFontSize(18).setBold(true).setForegroundColor("#ffffff");
      值框.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

      var 標籤框 = KPI頁.insertTextBox(kpiData[k].標籤, x + 10, 150, 130, 30);
      標籤框.getText().getTextStyle().setFontSize(11).setForegroundColor("#e0e0e0");
      標籤框.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    }

    // 部門排行表格
    var 排行標題 = KPI頁.insertTextBox("🏆 部門年度業績排行", 40, 230, 640, 35);
    排行標題.getText().getTextStyle().setFontSize(18).setBold(true).setForegroundColor("#1a237e");

    var 表格列數 = Math.min(部門統計.length + 1, 7);
    var table = KPI頁.insertTable(表格列數, 4);

    // 表格標題
    var thData = ["排名", "部門", "年度業績", "Q1→Q4 成長率"];
    for (var th = 0; th < 4; th++) {
      table.getCell(0, th).getText().setText(thData[th]);
      table.getCell(0, th).getText().getTextStyle().setFontSize(10).setBold(true).setForegroundColor("#ffffff");
      table.getCell(0, th).getFill().setSolidFill("#1a237e");
    }

    for (var r = 0; r < Math.min(部門統計.length, 6); r++) {
      var d = 部門統計[r];
      var rowData = [String(r + 1), d.部門, "NT$ " + (d.年度 / 10000).toFixed(0) + " 萬", d.成長率 + "%"];
      for (var c = 0; c < 4; c++) {
        table.getCell(r + 1, c).getText().setText(rowData[c]);
        table.getCell(r + 1, c).getText().getTextStyle().setFontSize(10);
        if (r % 2 === 1) {
          table.getCell(r + 1, c).getFill().setSolidFill("#e8eaf6");
        }
      }
    }

    // --- 第 3-N 頁：圖表頁 ---
    if (圖表列表.length > 0) {
      for (var ci = 0; ci < 圖表列表.length; ci++) {
        var 圖表頁 = ppt.appendSlide(SlidesApp.PredefinedLayout.BLANK);
        圖表頁.getBackground().setSolidFill("#ffffff");

        var chartTitle = 圖表列表[ci].getOptions().get("title") || "圖表 " + (ci + 1);
        var ct = 圖表頁.insertTextBox("📊 " + chartTitle, 40, 10, 640, 40);
        ct.getText().getTextStyle().setFontSize(24).setBold(true).setForegroundColor("#1a237e");

        var blob = 圖表列表[ci].getBlob();
        var img = 圖表頁.insertImage(blob);
        img.setLeft(40);
        img.setTop(60);
        img.setWidth(640);
        img.setHeight(380);
      }
    }

    if (趨勢圖表列表.length > 0) {
      for (var ti = 0; ti < 趨勢圖表列表.length; ti++) {
        var 趨勢頁 = ppt.appendSlide(SlidesApp.PredefinedLayout.BLANK);
        趨勢頁.getBackground().setSolidFill("#ffffff");

        var trendTitle = 趨勢圖表列表[ti].getOptions().get("title") || "趨勢圖 " + (ti + 1);
        var tt = 趨勢頁.insertTextBox("📈 " + trendTitle, 40, 10, 640, 40);
        tt.getText().getTextStyle().setFontSize(24).setBold(true).setForegroundColor("#1a237e");

        var tBlob = 趨勢圖表列表[ti].getBlob();
        var tImg = 趨勢頁.insertImage(tBlob);
        tImg.setLeft(40);
        tImg.setTop(60);
        tImg.setWidth(640);
        tImg.setHeight(380);
      }
    }

    // --- 最後一頁：結語 ---
    var 結語頁 = ppt.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    結語頁.getBackground().setSolidFill("#0d47a1");

    var endText = 結語頁.insertTextBox("感謝聆聽", 40, 150, 640, 60);
    endText.getText().getTextStyle().setFontSize(40).setBold(true).setForegroundColor("#ffffff");
    endText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    var endSub = 結語頁.insertTextBox(
      "本報告由 Google Apps Script 自動產生\n" +
      "資料來源：" + ss.getName(),
      40, 230, 640, 50
    );
    endSub.getText().getTextStyle().setFontSize(14).setForegroundColor("#90caf9");
    endSub.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // ========== 完成 ==========
    var 結束時間 = new Date();
    var 耗時 = Math.round((結束時間 - 開始時間) / 1000);

    Logger.log("✅ 報表簡報已產生！耗時 " + 耗時 + " 秒");
    Logger.log("簡報 URL：" + ppt.getUrl());

    SpreadsheetApp.getUi().alert(
      "🎉 報表簡報已成功產生！\n\n" +
      "📄 共 " + ppt.getSlides().length + " 頁\n" +
      "⏱️ 耗時 " + 耗時 + " 秒\n\n" +
      "📎 簡報名稱：" + 簡報名 + "\n\n" +
      "🔗 點選連結開啟：\n" + ppt.getUrl()
    );

  } catch (錯誤) {
    Logger.log("❌ 報表簡報錯誤：" + 錯誤.message);
    Logger.log("錯誤堆疊：" + 錯誤.stack);
    SpreadsheetApp.getUi().alert("❌ 發生錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 初始化（沿用 Session 9 的資料）
// ============================================================

/**
 * 確認資料是否就緒
 */
function 檢查資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var 圖表表 = ss.getSheetByName("圖表資料");
  var 趨勢表 = ss.getSheetByName("月度趨勢");

  var 訊息 = "📋 資料檢查結果：\n\n";

  if (圖表表) {
    var 圖表數 = 圖表表.getCharts().length;
    訊息 += "✅ 圖表資料：已就緒（" + 圖表數 + " 個圖表）\n";
  } else {
    訊息 += "❌ 圖表資料：未建立\n";
  }

  if (趨勢表) {
    var 趨勢圖表數 = 趨勢表.getCharts().length;
    訊息 += "✅ 月度趨勢：已就緒（" + 趨勢圖表數 + " 個圖表）\n";
  } else {
    訊息 += "❌ 月度趨勢：未建立\n";
  }

  if (!圖表表 || !趨勢表) {
    訊息 += "\n⚠️ 請先執行「初始化圖表資料」！";
  }

  SpreadsheetApp.getUi().alert(訊息);
}

// ============================================================
// 自訂選單
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📚 Session 10 工具")
    .addItem("🔍 檢查資料", "檢查資料")
    .addSeparator()
    .addItem("📝 建立基本簡報", "建立基本簡報")
    .addItem("📑 建立多頁簡報", "建立多頁簡報")
    .addItem("📊 插入圖表到簡報", "插入圖表到簡報")
    .addSeparator()
    .addItem("🚀 一鍵產生報表簡報", "一鍵產生報表簡報")
    .addToUi();
}
