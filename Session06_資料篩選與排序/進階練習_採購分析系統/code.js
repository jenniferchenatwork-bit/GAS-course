// ============================================================
// 進階練習：智慧採購分析與供應商評比
// 對應：Session 6（filter、sort、資料摘要）
// ============================================================

/**
 * 供應商綜合評比與排名（包含採購總金額）
 */
function 供應商評比() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("採購紀錄");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("❌ 請先執行「初始化採購資料」");
      return;
    }

    var 資料 = sheet.getDataRange().getValues();
    var 標題 = 資料[0];
    
    var 紀錄 = [];
    for (var i = 1; i < 資料.length; i++) {
      var obj = {};
      for (var j = 0; j < 標題.length; j++) {
        obj[標題[j]] = 資料[i][j];
      }
      紀錄.push(obj);
    }

    // --- 1. 依供應商分組統計與計算總額 ---
    var 供應商統計 = {};
    var 至今總採購額 = 0;

    紀錄.forEach(function(r) {
      var 供應商 = r["供應商"];
      var 金額 = r["金額"] || 0;
      
      至今總採購額 += 金額;

      if (!供應商統計[供應商]) {
        供應商統計[供應商] = {
          訂單數: 0,
          總金額: 0,
          準時次數: 0,
          退貨次數: 0,
          品質總分: 0
        };
      }
      
      var 統計 = 供應商統計[供應商];
      統計.訂單數++;
      統計.總金額 += 金額;
      
      if (r["交貨狀態"] === "準時") 統計.準時次數++;
      if (r["退貨"] === "是") 統計.退貨次數++;
      統計.品質總分 += (r["品質評分"] || 0);
    });

    // --- 2. 計算綜合評分 ---
    var 排名陣列 = [];
    
    for (var 名稱 in 供應商統計) {
      var 統計 = 供應商統計[名稱];
      
      var 品質平均 = 統計.品質總分 / 統計.訂單數;
      var 準時率 = (統計.準時次數 / 統計.訂單數) * 100;
      var 退貨率 = (統計.退貨次數 / 統計.訂單數) * 100;
      
      // 假設價格評分為 85 分 (因範例資料單價固定，此處以固定分數做計算基準)
      var 價格評分 = 85; 

      // 綜合評分公式：品質(40%) + 準時率(30%) + (100 - 退貨率)(20%) + 價格(10%)
      var 綜合評分 = (品質平均 * 0.4) + (準時率 * 0.3) + ((100 - 退貨率) * 0.2) + (價格評分 * 0.1);

      排名陣列.push({
        供應商: 名稱,
        訂單數: 統計.訂單數,
        總金額: 統計.總金額,
        品質: 品質平均,
        準時率: 準時率,
        退貨率: 退貨率,
        綜合評分: 綜合評分
      });
    }

    // 依綜合評分由高到低排序
    排名陣列.sort(function(a, b) {
      return b.綜合評分 - a.綜合評分;
    });

    // --- 3. 輸出至另一個分頁 ---
    var 報表名稱 = "供應商排名報表";
    var 報表 = ss.getSheetByName(報表名稱);
    if (報表) 報表.clear(); 
    else 報表 = ss.insertSheet(報表名稱);

    // 寫入總部資訊
    報表.getRange("A1").setValue("🏆 供應商綜合評比與採購總額").setFontSize(16).setFontWeight("bold");
    報表.getRange("A2").setValue("至今採購總金額：$" + 至今總採購額.toLocaleString()).setFontSize(12).setFontColor("#d32f2f").setFontWeight("bold");
    報表.getRange("A3").setValue("製表日期：" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd"));

    // 寫入表頭
    var 表頭 = [["排名", "供應商", "訂單數", "採購總金額", "品質平均", "準時率", "退貨率", "綜合評分", "評等"]];
    報表.getRange(5, 1, 1, 9).setValues(表頭).setBackground("#1565c0").setFontColor("#fff").setFontWeight("bold");

    // 寫入資料
    var 寫入資料 = [];
    排名陣列.forEach(function(資料, 索引) {
      var 排名 = 索引 + 1;
      var 評等 = 資料.綜合評分 >= 80 ? "⭐ A 級" : (資料.綜合評分 >= 60 ? "🔵 B 級" : "🔴 C 級");
      
      寫入資料.push([
        排名,
        資料.供應商,
        資料.訂單數,
        資料.總金額,
        資料.品質.toFixed(1),
        資料.準時率.toFixed(1) + "%",
        資料.退貨率.toFixed(1) + "%",
        資料.綜合評分.toFixed(2),
        評等
      ]);
    });

    if (寫入資料.length > 0) {
      報表.getRange(6, 1, 寫入資料.length, 9).setValues(寫入資料);
      報表.getRange(6, 4, 寫入資料.length, 1).setNumberFormat("#,##0"); // 金額加上千分位
      // 標註第一名背景顏色
      報表.getRange(6, 1, 1, 9).setBackground("#e8f5e9");
    }

    // 自動調整欄寬
    for (var c = 1; c <= 9; c++) 報表.autoResizeColumn(c);

    SpreadsheetApp.getUi().alert("✅ 供應商排名報表已生成！\n至今採購總額為：$" + 至今總採購額.toLocaleString());

  } catch (錯誤) {
    Logger.log("❌ 發生錯誤：" + 錯誤.message);
  }
}

/**
 * 採購需求預測（根據歷史消耗趨勢）
 */
function 採購預測() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("採購紀錄");
  if (!sheet) return;

  var 資料 = sheet.getDataRange().getValues();
  var 品項統計 = {};

  for (var i = 1; i < 資料.length; i++) {
    var 品項 = 資料[i][1]; // B: 品項
    var 數量 = 資料[i][3]; // D: 數量
    if (!品項統計[品項]) 品項統計[品項] = [];
    品項統計[品項].push(數量);
  }

  var 預測 = [];
  for (var name in 品項統計) {
    var 歷史 = 品項統計[name];
    var 平均 = 歷史.reduce(function(a, b) { return a + b; }, 0) / 歷史.length;
    var 建議量 = Math.ceil(平均 * 1.2); // 多 20% 安全庫存

    預測.push(name + "：月均 " + Math.round(平均) + " → 建議採購 " + 建議量);
  }

  SpreadsheetApp.getUi().alert("📊 採購預測\n\n" + 預測.join("\n"));
}

/**
 * 自動產生採購建議書並結合 Gmail 發送（進階挑戰）
 */
function 產生並發送採購建議書() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("採購紀錄");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("❌ 請先執行「初始化採購資料」");
      return;
    }

    var 資料 = sheet.getDataRange().getValues();
    var 標題 = 資料[0];
    var 紀錄 = [];
    for (var i = 1; i < 資料.length; i++) {
      var obj = {};
      for (var j = 0; j < 標題.length; j++) obj[標題[j]] = 資料[i][j];
      紀錄.push(obj);
    }

    // --- 1. 計算供應商評比，找出推薦（第一名）供應商 ---
    var 供應商統計 = {};
    var 品項單價 = {}; // 紀錄各品項單價
    var 品項消耗 = {}; // 紀錄各品項消耗量

    紀錄.forEach(function(r) {
      var 供應商 = r["供應商"];
      var 品項 = r["品項"];
      var 數量 = r["數量"];
      var 單價 = r["單價"];

      // 紀錄單價與消耗
      品項單價[品項] = 單價;
      if (!品項消耗[品項]) 品項消耗[品項] = [];
      品項消耗[品項].push(數量);

      if (!供應商統計[供應商]) {
        供應商統計[供應商] = { 訂單數: 0, 準時次數: 0, 退貨次數: 0, 品質總分: 0 };
      }
      var 統計 = 供應商統計[供應商];
      統計.訂單數++;
      if (r["交貨狀態"] === "準時") 統計.準時次數++;
      if (r["退貨"] === "是") 統計.退貨次數++;
      統計.品質總分 += (r["品質評分"] || 0);
    });

    var 排名陣列 = [];
    for (var 名稱 in 供應商統計) {
      var 統計 = 供應商統計[名稱];
      var 品質平均 = 統計.品質總分 / 統計.訂單數;
      var 準時率 = (統計.準時次數 / 統計.訂單數) * 100;
      var 退貨率 = (統計.退貨次數 / 統計.訂單數) * 100;
      // 使用與之前相同的綜合評分公式，價格以基準分 85 計算
      var 綜合評分 = (品質平均 * 0.4) + (準時率 * 0.3) + ((100 - 退貨率) * 0.2) + (85 * 0.1);
      排名陣列.push({
        供應商: 名稱, 綜合評分: 綜合評分, 準時率: 準時率, 退貨率: 退貨率
      });
    }
    
    // 降冪排序，取出第一名做為推薦供應商
    排名陣列.sort(function(a, b) { return b.綜合評分 - a.綜合評分; });
    var 冠軍 = 排名陣列[0]; 

    // --- 2. 預測下次採購需求 ---
    var 採購清單 = [];
    var 預估總額 = 0;
    for (var p in 品項消耗) {
      var 歷史 = 品項消耗[p];
      var 平均 = 歷史.reduce(function(a, b) { return a + b; }, 0) / 歷史.length;
      var 建議量 = Math.ceil(平均 * 1.2); // 包含 20% 安全庫存
      var 小計 = 建議量 * 品項單價[p];
      預估總額 += 小計;
      採購清單.push(" - " + p + "：" + 建議量 + " 單位 (預估金額：$" + 小計.toLocaleString() + ")");
    }

    // --- 3. 準備 Email 內容 ---
    var 信件主旨 = "【採購建議書】下期採購合作與評比結果 - " + 冠軍.供應商;
    var 信件內容 = 
      "致 " + 冠軍.供應商 + " 團隊：\n\n" +
      "感謝貴公司過去的合作，以下為本期的供應商評比結果：\n" +
      "✅ 綜合評分：" + 冠軍.綜合評分.toFixed(1) + " 分（本期排名第 1）\n" +
      "✅ 準時交貨率：" + 冠軍.準時率.toFixed(1) + "%\n" +
      "✅ 退貨率：" + 冠軍.退貨率.toFixed(1) + "%\n\n" +
      "恭喜貴公司獲得本次評比的最高評價！\n" +
      "基於上述優異表現，我們預計向貴公司進行下一期的採購，以下為預估的品項與數量：\n\n" +
      採購清單.join("\n") + "\n" +
      "-----------------------------------------\n" +
      "💰 預估採購總金額：$" + 預估總額.toLocaleString() + "\n\n" +
      "請確認是否能配合上述需求與現有報價，謝謝！\n\n" +
      "採購部 敬上\n" + 
      "日期：" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd");

    // --- 4. 發送 Email ---
    var 目標信箱 = "jenniferchenatwork@gmail.com";
    GmailApp.sendEmail(目標信箱, 信件主旨, 信件內容, {
      cc: 目標信箱 // 同時副本給同一個信箱
    });

    // --- 5. 介面回饋 ---
    SpreadsheetApp.getUi().alert(
      "✅ 採購建議書已成功發送！\n\n" +
      "🏆 推薦供應商：" + 冠軍.供應商 + "\n" +
      "💰 預估總額：$" + 預估總額.toLocaleString() + "\n" +
      "📧 收件人與副本：" + 目標信箱
    );

  } catch (錯誤) {
    Logger.log("❌ 發生錯誤：" + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ 發送失敗：" + 錯誤.message);
  }
}

function 初始化採購資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("採購紀錄");
  if (!sheet) sheet = ss.insertSheet("採購紀錄"); else sheet.clear();

  var 標題 = [["供應商", "品項", "單價", "數量", "金額", "交貨狀態", "退貨", "品質評分", "日期"]];
  var 供應商 = ["宏達文具", "大同辦公", "金鼎耗材", "永豐科技", "佳能事務"];
  var 品項 = ["A4影印紙", "碳粉匣", "原子筆", "資料夾", "白板筆"];
  var 資料 = [];

  for (var i = 0; i < 40; i++) {
    var s = 供應商[Math.floor(Math.random() * 供應商.length)];
    var p = 品項[Math.floor(Math.random() * 品項.length)];
    var 單價 = [150, 1200, 15, 25, 45][品項.indexOf(p)];
    var 數量 = Math.floor(Math.random() * 50) + 5;
    var 交貨 = Math.random() > 0.2 ? "準時" : "延遲";
    var 退貨 = Math.random() > 0.85 ? "是" : "否";
    var 品質 = Math.floor(Math.random() * 40) + 60;
    var 日期 = new Date(2026, Math.floor(Math.random() * 4), Math.floor(Math.random() * 28) + 1);

    資料.push([s, p, 單價, 數量, 單價 * 數量, 交貨, 退貨, 品質, 日期]);
  }

  sheet.getRange(1, 1, 1, 9).setValues(標題);
  sheet.getRange(2, 1, 資料.length, 9).setValues(資料);
  sheet.getRange("A1:I1").setBackground("#6a1b9a").setFontColor("#fff").setFontWeight("bold");
  sheet.getRange("E2:E41").setNumberFormat("#,##0");
  sheet.getRange("I2:I41").setNumberFormat("yyyy/mm/dd");
  sheet.setFrozenRows(1);
  for (var c = 1; c <= 9; c++) sheet.autoResizeColumn(c);

  SpreadsheetApp.getUi().alert("✅ 40 筆採購紀錄已建立！");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🤖 智慧採購管理")
    .addItem("📦 初始化採購資料", "初始化採購資料")
    .addItem("🏆 供應商評比", "供應商評比")
    .addItem("📊 採購預測", "採購預測")
    .addSeparator()
    .addItem("📧 產生並發送採購建議書", "產生並發送採購建議書")
    .addToUi();
}
