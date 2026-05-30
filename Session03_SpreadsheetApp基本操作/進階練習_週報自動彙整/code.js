// ============================================================
// 進階練習：多部門週報自動彙整系統
// 主題：辦公室自動化 AI — 週報管理
// 對應：Session 3（SpreadsheetApp、建立工作表、觸發器）
// ============================================================

/**
 * 自動為每個部門建立週報工作表
 */
function 建立本週週報() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 今天 = new Date();
    var 週一 = new Date(今天);
    週一.setDate(今天.getDate() - 今天.getDay() + 1);
    var 週五 = new Date(週一);
    週五.setDate(週一.getDate() + 4);

    var 週區間 = Utilities.formatDate(週一, "Asia/Taipei", "MMdd") + "-" +
                 Utilities.formatDate(週五, "Asia/Taipei", "MMdd");

    var 部門 = ["業務部", "行銷部", "研發部", "人資部", "財務部"];

    部門.forEach(function(部門名) {
      var 表名 = "週報_" + 部門名 + "_" + 週區間;

      // 跳過已存在的
      if (ss.getSheetByName(表名)) {
        Logger.log("⚠️ " + 表名 + " 已存在，跳過");
        return;
      }

      var sheet = ss.insertSheet(表名);

      // 標題
      sheet.getRange("A1:E1").merge();
      sheet.getRange("A1").setValue("📋 " + 部門名 + " 週報")
        .setFontSize(16).setFontWeight("bold").setBackground("#1565c0").setFontColor("#fff");

      sheet.getRange("A2").setValue("報告期間：" + 週區間);
      sheet.getRange("A3").setValue("填報人：").setFontWeight("bold");

      // 本週完成事項
      sheet.getRange("A5:E5").merge();
      sheet.getRange("A5").setValue("✅ 本週完成事項").setFontSize(13).setFontWeight("bold").setBackground("#e8f5e9");
      sheet.getRange("A6:E6").setValues([["序號", "工作項目", "負責人", "狀態", "備註"]]);
      sheet.getRange("A6:E6").setBackground("#43a047").setFontColor("#fff").setFontWeight("bold");
      for (var i = 7; i <= 12; i++) sheet.getRange(i, 1).setValue(i - 6);

      // 下週預計工作
      sheet.getRange("A14:E14").merge();
      sheet.getRange("A14").setValue("📝 下週預計工作").setFontSize(13).setFontWeight("bold").setBackground("#fff3e0");
      sheet.getRange("A15:E15").setValues([["序號", "工作項目", "負責人", "優先級", "預計完成"]]);
      sheet.getRange("A15:E15").setBackground("#ef6c00").setFontColor("#fff").setFontWeight("bold");
      for (var j = 16; j <= 21; j++) sheet.getRange(j, 1).setValue(j - 15);

      // 問題與需求
      sheet.getRange("A23:E23").merge();
      sheet.getRange("A23").setValue("⚠️ 問題與支援需求").setFontSize(13).setFontWeight("bold").setBackground("#fce4ec");
      sheet.getRange("A24:E24").setValues([["序號", "問題描述", "影響範圍", "建議解法", "需要支援"]]);
      sheet.getRange("A24:E24").setBackground("#c62828").setFontColor("#fff").setFontWeight("bold");

      // 欄寬
      sheet.setColumnWidth(1, 60);
      sheet.setColumnWidth(2, 280);
      sheet.setColumnWidth(3, 100);
      sheet.setColumnWidth(4, 100);
      sheet.setColumnWidth(5, 150);
      sheet.setFrozenRows(1);

      Logger.log("✅ " + 表名 + " 已建立");
    });

    SpreadsheetApp.getUi().alert("✅ 本週週報工作表已為 " + 部門.length + " 個部門建立！");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ 錯誤：" + 錯誤.message);
  }
}

/**
 * 自動彙整：合併所有部門週報到一個「週報總覽」
 */
function 彙整週報() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var 總覽表 = ss.getSheetByName("週報總覽");
    if (總覽表) 總覽表.clear(); else 總覽表 = ss.insertSheet("週報總覽", 0);

    // 標題
    總覽表.getRange("A1:F1").merge();
    總覽表.getRange("A1").setValue("📊 各部門週報彙整總覽")
      .setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
    總覽表.getRange("A2").setValue("彙整時間：" +
      Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm"));

    var 目前列 = 4;

    sheets.forEach(function(sheet) {
      if (sheet.getName().indexOf("週報_") !== 0) return;

      var 部門名 = sheet.getName().replace("週報_", "").split("_")[0];

      // 部門標題
      總覽表.getRange(目前列, 1, 1, 6).merge();
      總覽表.getRange(目前列, 1).setValue("🏢 " + 部門名)
        .setFontSize(14).setFontWeight("bold").setBackground("#e3f2fd");
      目前列++;

      // 讀取完成事項（第 7~12 列，A~E 欄）
      var 完成資料 = sheet.getRange(7, 1, 6, 5).getValues();
      var 有資料 = false;

      完成資料.forEach(function(row) {
        if (row[1] && String(row[1]).trim() !== "") {
          總覽表.getRange(目前列, 1, 1, 5).setValues([row]);
          有資料 = true;
          目前列++;
        }
      });

      if (!有資料) {
        總覽表.getRange(目前列, 1).setValue("（尚未填寫）").setFontColor("#999");
        目前列++;
      }

      目前列++; // 空一列
    });

    for (var c = 1; c <= 5; c++) {
      總覽表.autoResizeColumn(c);
      var 目前寬度 = 總覽表.getColumnWidth(c);
      總覽表.setColumnWidth(c, 目前寬度 + 30); // 增加 30 像素緩衝
    }

    SpreadsheetApp.getUi().alert("✅ 週報彙整完成！請查看「週報總覽」工作表。");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

/**
 * 設定每週五下午 5 點自動彙整觸發器
 */
function 設定週五自動彙整() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "彙整週報") ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger("彙整週報")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(17)
    .create();

  SpreadsheetApp.getUi().alert("✅ 每週五 17:00 自動彙整已設定！");
}

/**
 * 設定每週一早上自動建立新週報
 */
function 設定週一自動建立() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "建立本週週報") ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger("建立本週週報")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();

  SpreadsheetApp.getUi().alert("✅ 每週一 08:00 自動建立週報已設定！");
}

/**
 * 為現有的週報工作表填入範例資料
 */
function 新增週報範例資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var 有更新 = false;

  // 定義各部門的專屬範例資料（讓模擬效果更真實）
  var 部門範例 = {
    "業務部": {
      "完成": [
        [1, "拜訪 A 客戶進行續約談判", "張小明", "✅ 已完成", "合約已簽署"],
        [2, "開發 3 位潛在客戶並寄送提案", "李小華", "⏳ 進行中", "待客戶回覆"],
        [3, "更新 CRM 系統客戶資料", "王大同", "✅ 已完成", "無"]
      ],
      "預計": [
        [1, "參加產業年度展覽會", "張小明", "🔥 高", "2026/05/12"],
        [2, "業務月會簡報準備", "李小華", "⚡️ 中", "2026/05/14"]
      ],
      "問題": [1, "競爭對手推出低價促銷", "全體業務", "研擬對應折扣方案", "需市場部支援"]
    },
    "行銷部": {
      "完成": [
        [1, "製作本季廣告視覺素材", "林美玲", "✅ 已完成", "已交付設計稿"],
        [2, "分析 Facebook 廣告轉換率", "趙小芳", "✅ 已完成", "ROAS 達到 4.5"],
        [3, "策劃 618 促銷活動提案", "林美玲", "⏳ 進行中", "預計週五定案"]
      ],
      "預計": [
        [1, "與網紅進行合作洽談", "趙小芳", "🔥 高", "2026/05/15"],
        [2, "SEO 關鍵字優化調整", "林美玲", "⚡️ 中", "2026/05/16"]
      ],
      "問題": [1, "廣告帳戶餘額不足", "行銷組", "向財務部申請預算", "需急件處理"]
    },
    "研發部": {
      "完成": [
        [1, "修復 API 逾時偶發問題", "陳大文", "✅ 已完成", "已部署至測試機"],
        [2, "開發新功能 A 測試介面", "吳資深", "⏳ 進行中", "前端串接完畢"],
        [3, "執行 Code Review 與重構", "陳大文", "✅ 已完成", "提升效能 20%"]
      ],
      "預計": [
        [1, "資料庫結構遷移作業", "吳資深", "🔥 高", "2026/05/13"],
        [2, "撰寫新功能技術文件", "陳大文", "⚡️ 中", "2026/05/17"]
      ],
      "問題": [1, "開發環境資源不足", "研發組", "申請增加雲端實例", "需主管核准"]
    },
    "人資部": {
      "完成": [
        [1, "招募資深工程師面試(3位)", "許人資", "✅ 已完成", "1位進入二面"],
        [2, "舉辦新進員工教育訓練", "許人資", "✅ 已完成", "滿意度 95%"],
        [3, "更新員工年度考核基準", "張經理", "⏳ 進行中", "草案擬定中"]
      ],
      "預計": [
        [1, "籌備端午節員工活動", "許人資", "⚡️ 中", "2026/05/20"],
        [2, "修訂員工手冊福利專章", "張經理", "🔥 高", "2026/05/14"]
      ],
      "問題": [1, "員工離職率略有上升", "管理層", "進行離職訪談分析", "需各部門回饋"]
    },
    "財務部": {
      "完成": [
        [1, "審核本月各部門報支單據", "財務專員", "✅ 已完成", "全數入帳"],
        [2, "準備年度稅務申報報表", "財務主管", "⏳ 進行中", "完成 70%"],
        [3, "核對銀行往來調節表", "財務專員", "✅ 已完成", "無誤"]
      ],
      "預計": [
        [1, "撥發本月員工薪資", "財務專員", "🔥 高", "2026/05/10"],
        [2, "各專案預算執行狀況查核", "財務主管", "⚡️ 中", "2026/05/18"]
      ],
      "問題": [1, "部分發票報支不符規範", "各部門", "加強報支流程宣導", "無"]
    }
  };

  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    if (name.indexOf("週報_") !== 0) return;

    var 部門名 = name.replace("週報_", "").split("_")[0];
    var data = 部門範例[部門名];
    
    if (!data) return; // 若不在範例清單中則跳過

    // 填入填報人
    sheet.getRange("A3").setValue("填報人：範例機器人 (" + 部門名 + ")");

    // 1. 填入本週完成事項
    sheet.getRange(7, 1, data.完成.length, 5).setValues(data.完成);

    // 2. 填入下週預計工作
    sheet.getRange(16, 1, data.預計.length, 5).setValues(data.預計);

    // 3. 填入問題與需求
    sheet.getRange(25, 1, 1, 5).setValues([data.問題]);
    
    有更新 = true;
  });

  if (有更新) {
    SpreadsheetApp.getUi().alert("✅ 已成功為各部門（業務、行銷、研發、人資、財務）填入「專屬」範例資料！\n現在您可以點選「彙整所有週報」來測試彙整功能。");
  } else {
    SpreadsheetApp.getUi().alert("⚠️ 找不到週報工作表，請先執行「建立本週週報」。");
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🤖 智慧週報系統")
    .addItem("📋 建立本週週報", "建立本週週報")
    .addItem("📝 新增週報範例資料", "新增週報範例資料")
    .addItem("📊 彙整所有週報", "彙整週報")
    .addSeparator()
    .addItem("⏰ 設定週一自動建立", "設定週一自動建立")
    .addItem("⏰ 設定週五自動彙整", "設定週五自動彙整")
    .addToUi();
}


