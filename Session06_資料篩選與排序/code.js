// ============================================================
// Session 6：資料篩選與排序
// 日期：115/05/16（六）13:30~16:30
// 講師：林冠廷
// ============================================================
// 本課程涵蓋：
//   1. 資料篩選 (filter)、排序 (sort)
//   2. 條件判斷與資料摘要
//   3. 執行限制（記憶體與時間）
//   4. 實作：自動化資料處理工具
// ============================================================

// ============================================================
// 第一部分：資料篩選 (Filter)
// ============================================================

/**
 * 篩選銷售資料 — 依條件過濾
 */
function 篩選示範() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("銷售紀錄");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ 請先執行「初始化銷售資料」");
    return;
  }

  var 資料 = sheet.getDataRange().getValues();
  var 標題 = 資料[0];

  // 轉成物件陣列
  var 銷售 = [];
  for (var i = 1; i < 資料.length; i++) {
    var obj = {};
    for (var j = 0; j < 標題.length; j++) {
      obj[標題[j]] = 資料[i][j];
    }
    銷售.push(obj);
  }

  // --- 篩選 1：金額大於 5000 的訂單 ---
  var 大額訂單 = 銷售.filter(function(item) {
    return item["金額"] > 5000;
  });
  Logger.log("💰 大額訂單（>5000）：" + 大額訂單.length + " 筆");
  大額訂單.forEach(function(item) {
    Logger.log("  " + item["客戶"] + " - " + item["商品"] + " - $" + item["金額"]);
  });

  // --- 篩選 2：特定業務人員的業績 ---
  var 指定業務 = "王小明";
  var 業務訂單 = 銷售.filter(function(item) {
    return item["業務"] === 指定業務;
  });
  Logger.log("\n👤 " + 指定業務 + " 的訂單：" + 業務訂單.length + " 筆");

  // --- 篩選 3：多條件篩選（金額 > 3000 且 本月的訂單）---
  var 今天 = new Date();
  var 本月訂單 = 銷售.filter(function(item) {
    var 訂單日期 = new Date(item["日期"]);
    return item["金額"] > 3000 &&
           訂單日期.getMonth() === 今天.getMonth() &&
           訂單日期.getFullYear() === 今天.getFullYear();
  });
  Logger.log("\n📅 本月大額訂單：" + 本月訂單.length + " 筆");
}

// ============================================================
// 第二部分：資料排序 (Sort)
// ============================================================

/**
 * 排序銷售資料
 */
function 排序示範() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("銷售紀錄");
  if (!sheet) return;

  var 資料 = sheet.getDataRange().getValues();
  var 標題 = 資料.shift(); // 取出標題列

  // --- 排序 1：依金額由大到小 ---
  var 金額排序 = 資料.slice().sort(function(a, b) {
    return b[4] - a[4]; // 第 5 欄（金額），降冪
  });
  Logger.log("===== 金額排序（高→低）=====");
  for (var i = 0; i < Math.min(5, 金額排序.length); i++) {
    Logger.log(金額排序[i][1] + " - $" + 金額排序[i][4]);
  }

  // --- 排序 2：依日期排序 ---
  var 日期排序 = 資料.slice().sort(function(a, b) {
    return new Date(a[5]) - new Date(b[5]); // 第 6 欄（日期），升冪
  });
  Logger.log("\n===== 日期排序（舊→新）=====");
  for (var j = 0; j < Math.min(5, 日期排序.length); j++) {
    Logger.log(日期排序[j][5] + " - " + 日期排序[j][1]);
  }

  // --- 排序 3：多欄位排序（先依業務，再依金額）---
  var 多欄排序 = 資料.slice().sort(function(a, b) {
    // 先比較業務名稱
    if (a[0] < b[0]) return -1;
    if (a[0] > b[0]) return 1;
    // 業務相同則比較金額（降冪）
    return b[4] - a[4];
  });
  Logger.log("\n===== 多欄排序（業務 + 金額）=====");
  for (var k = 0; k < Math.min(8, 多欄排序.length); k++) {
    Logger.log(多欄排序[k][0] + " - " + 多欄排序[k][1] + " - $" + 多欄排序[k][4]);
  }
}

// ============================================================
// 第三部分：條件判斷與資料摘要
// ============================================================

/**
 * 生成銷售摘要報告
 * 說明：綜合運用篩選、排序、統計，產生完整報告
 */
function 生成銷售摘要() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("銷售紀錄");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("❌ 請先執行「初始化銷售資料」");
      return;
    }

    var 資料 = sheet.getDataRange().getValues();
    var 標題 = 資料[0];

    // 轉換為物件陣列
    var 銷售 = [];
    for (var i = 1; i < 資料.length; i++) {
      var obj = {};
      for (var j = 0; j < 標題.length; j++) obj[標題[j]] = 資料[i][j];
      銷售.push(obj);
    }

    // --- 業務業績統計 ---
    var 業務統計 = {};
    銷售.forEach(function(item) {
      var 業務 = item["業務"];
      if (!業務統計[業務]) {
        業務統計[業務] = { 筆數: 0, 總額: 0, 最大單筆: 0 };
      }
      業務統計[業務].筆數++;
      業務統計[業務].總額 += item["金額"];
      if (item["金額"] > 業務統計[業務].最大單筆) {
        業務統計[業務].最大單筆 = item["金額"];
      }
    });

    // --- 商品銷售統計 ---
    var 商品統計 = {};
    銷售.forEach(function(item) {
      var 商品 = item["商品"];
      if (!商品統計[商品]) {
        商品統計[商品] = { 數量: 0, 總額: 0 };
      }
      商品統計[商品].數量 += item["數量"];
      商品統計[商品].總額 += item["金額"];
    });

    // --- 建立摘要工作表 ---
    var 摘要表 = ss.getSheetByName("銷售摘要");
    if (摘要表) 摘要表.clear(); else 摘要表 = ss.insertSheet("銷售摘要");

    // 標題
    摘要表.getRange("A1").setValue("📊 銷售分析摘要報告");
    摘要表.getRange("A1").setFontSize(16).setFontWeight("bold");
    摘要表.getRange("A2").setValue("報告日期：" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm"));
    摘要表.getRange("A3").setValue("總交易筆數：" + 銷售.length + " 筆");

    // 業務業績表
    var 列 = 5;
    摘要表.getRange(列, 1).setValue("🏆 業務業績排行");
    摘要表.getRange(列, 1).setFontSize(13).setFontWeight("bold");
    列++;

    摘要表.getRange(列, 1, 1, 4).setValues([["業務", "成交筆數", "總業績", "最大單筆"]]);
    摘要表.getRange(列, 1, 1, 4).setBackground("#1a73e8").setFontColor("#fff").setFontWeight("bold");
    列++;

    // 依總額排序業務
    var 業務排序 = Object.keys(業務統計).sort(function(a, b) {
      return 業務統計[b].總額 - 業務統計[a].總額;
    });

    業務排序.forEach(function(業務) {
      var s = 業務統計[業務];
      摘要表.getRange(列, 1, 1, 4).setValues([[業務, s.筆數, s.總額, s.最大單筆]]);
      摘要表.getRange(列, 3, 1, 2).setNumberFormat("#,##0");
      列++;
    });

    列 += 2;

    // 商品銷售表
    摘要表.getRange(列, 1).setValue("📦 商品銷售統計");
    摘要表.getRange(列, 1).setFontSize(13).setFontWeight("bold");
    列++;

    摘要表.getRange(列, 1, 1, 3).setValues([["商品", "銷售數量", "銷售總額"]]);
    摘要表.getRange(列, 1, 1, 3).setBackground("#34a853").setFontColor("#fff").setFontWeight("bold");
    列++;

    var 商品排序 = Object.keys(商品統計).sort(function(a, b) {
      return 商品統計[b].總額 - 商品統計[a].總額;
    });

    商品排序.forEach(function(商品) {
      var s = 商品統計[商品];
      摘要表.getRange(列, 1, 1, 3).setValues([[商品, s.數量, s.總額]]);
      摘要表.getRange(列, 3).setNumberFormat("#,##0");
      列++;
    });

    // 自動調整欄寬
    for (var c = 1; c <= 4; c++) 摘要表.autoResizeColumn(c);

    Logger.log("✅ 銷售摘要已生成！");
    SpreadsheetApp.getUi().alert("✅ 銷售摘要報告已生成！\n請查看「銷售摘要」工作表。");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 第四部分：自動化資料處理工具
// ============================================================

/**
 * 自動化資料清理
 * 說明：移除空白列、修正格式、標記異常值
 */
function 自動資料清理() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("銷售紀錄");
    if (!sheet) return;

    var 資料 = sheet.getDataRange().getValues();
    var 清理計數 = { 空白列: 0, 格式修正: 0, 異常值: 0 };

    // 從最後一列往前檢查（避免刪除列後索引錯亂）
    for (var i = 資料.length - 1; i >= 1; i--) {
      var 是空白列 = 資料[i].every(function(cell) {
        return cell === "" || cell === null || cell === undefined;
      });

      if (是空白列) {
        sheet.deleteRow(i + 1);
        清理計數.空白列++;
        continue;
      }

      // 檢查金額異常（負數或非數字）
      var 金額 = 資料[i][4];
      if (typeof 金額 !== "number" || 金額 < 0) {
        sheet.getRange(i + 1, 5).setBackground("#ffcdd2");  // 標記紅色
        sheet.getRange(i + 1, 5).setNote("⚠️ 異常值：金額應為正數");
        清理計數.異常值++;
      }
    }

    var 訊息 = "📋 資料清理完成！\n" +
      "移除空白列：" + 清理計數.空白列 + "\n" +
      "標記異常值：" + 清理計數.異常值;

    Logger.log(訊息);
    SpreadsheetApp.getUi().alert(訊息);

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

/**
 * 依條件匯出篩選結果到新工作表
 */
function 匯出篩選結果() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("銷售紀錄");
    if (!sheet) return;

    var 資料 = sheet.getDataRange().getValues();
    var 標題 = 資料[0];

    // 篩選：金額 > 3000 的訂單
    var 篩選結果 = [];
    for (var i = 1; i < 資料.length; i++) {
      if (資料[i][4] > 3000) {
        篩選結果.push(資料[i]);
      }
    }

    // 依金額排序（降冪）
    篩選結果.sort(function(a, b) { return b[4] - a[4]; });

    // 建立匯出工作表
    var 匯出表名 = "篩選匯出_" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd");
    var 匯出表 = ss.getSheetByName(匯出表名);
    if (匯出表) 匯出表.clear(); else 匯出表 = ss.insertSheet(匯出表名);

    // 寫入標題
    匯出表.getRange(1, 1, 1, 標題.length).setValues([標題]);
    匯出表.getRange(1, 1, 1, 標題.length).setBackground("#e65100").setFontColor("#fff").setFontWeight("bold");

    // 寫入篩選結果
    if (篩選結果.length > 0) {
      匯出表.getRange(2, 1, 篩選結果.length, 標題.length).setValues(篩選結果);
      匯出表.getRange(2, 5, 篩選結果.length, 1).setNumberFormat("#,##0");
    }

    匯出表.setFrozenRows(1);
    for (var c = 1; c <= 標題.length; c++) 匯出表.autoResizeColumn(c);

    SpreadsheetApp.getUi().alert("✅ 已匯出 " + 篩選結果.length + " 筆金額 > $3,000 的訂單。");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 初始化範例資料
// ============================================================

function 初始化銷售資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("銷售紀錄");
  if (!sheet) sheet = ss.insertSheet("銷售紀錄"); else sheet.clear();

  var 標題 = [["業務", "客戶", "商品", "數量", "金額", "日期"]];
  var 業務 = ["王小明", "李小華", "張美玲", "陳大文"];
  var 客戶 = ["台北科技", "新竹電子", "台中製造", "高雄物流", "花蓮文創", "桃園光電"];
  var 商品 = ["筆電", "伺服器", "路由器", "平板", "印表機", "監控系統"];

  var 資料 = [];
  for (var i = 0; i < 30; i++) {
    var 業務員 = 業務[Math.floor(Math.random() * 業務.length)];
    var 客戶名 = 客戶[Math.floor(Math.random() * 客戶.length)];
    var 商品名 = 商品[Math.floor(Math.random() * 商品.length)];
    var 數量 = Math.floor(Math.random() * 20) + 1;
    var 單價 = [500, 1200, 2500, 3800, 5500, 8000, 12000][Math.floor(Math.random() * 7)];
    var 金額 = 數量 * 單價;
    var 日期 = new Date(2026, Math.floor(Math.random() * 4) + 1, Math.floor(Math.random() * 28) + 1);

    資料.push([業務員, 客戶名, 商品名, 數量, 金額, 日期]);
  }

  sheet.getRange(1, 1, 1, 6).setValues(標題);
  sheet.getRange(2, 1, 資料.length, 6).setValues(資料);

  sheet.getRange("A1:F1").setBackground("#9c27b0").setFontColor("#fff").setFontWeight("bold");
  sheet.getRange("E2:E31").setNumberFormat("#,##0");
  sheet.getRange("F2:F31").setNumberFormat("yyyy/mm/dd");
  sheet.setFrozenRows(1);
  for (var c = 1; c <= 6; c++) sheet.autoResizeColumn(c);

  SpreadsheetApp.getUi().alert("✅ 已建立 30 筆銷售紀錄！");
}

// ============================================================
// 第五部分：進階銷售分析
// ============================================================

/**
 * 生成進階銷售分析（每月統計與排名）
 */
function 生成進階銷售分析() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("銷售紀錄");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("❌ 請先執行「初始化銷售資料」");
      return;
    }

    var 資料 = sheet.getDataRange().getValues();
    var 標題 = 資料[0];

    // 轉成物件陣列
    var 銷售 = [];
    for (var i = 1; i < 資料.length; i++) {
      var obj = {};
      for (var j = 0; j < 標題.length; j++) {
        obj[標題[j]] = 資料[i][j];
      }
      銷售.push(obj);
    }

    // 統計物件
    var 每月統計 = {};
    var 年度客戶 = {};

    銷售.forEach(function(item) {
      var 日期 = item["日期"];
      if (!(日期 instanceof Date)) 日期 = new Date(日期);
      if (isNaN(日期.getTime())) return;
      
      var 月份 = Utilities.formatDate(日期, "Asia/Taipei", "yyyy/MM");
      var 客戶 = item["客戶"];
      var 業務 = item["業務"];
      var 商品 = item["商品"];
      var 金額 = item["金額"] || 0;

      // 初始化每月統計
      if (!每月統計[月份]) {
        每月統計[月份] = { 總額: 0, 筆數: 0, 客戶資料: {} };
      }
      每月統計[月份].總額 += 金額;
      每月統計[月份].筆數++;

      // 每月客戶資料
      if (!每月統計[月份].客戶資料[客戶]) {
        每月統計[月份].客戶資料[客戶] = { 總額: 0, 業務清單: {}, 商品清單: {} };
      }
      每月統計[月份].客戶資料[客戶].總額 += 金額;
      每月統計[月份].客戶資料[客戶].業務清單[業務] = true;
      每月統計[月份].客戶資料[客戶].商品清單[商品] = true;

      // 年度客戶資料
      if (!年度客戶[客戶]) {
        年度客戶[客戶] = { 總額: 0, 業務清單: {}, 商品清單: {} };
      }
      年度客戶[客戶].總額 += 金額;
      年度客戶[客戶].業務清單[業務] = true;
      年度客戶[客戶].商品清單[商品] = true;
    });

    // 建立新分頁
    var 分析表 = ss.getSheetByName("進階銷售分析");
    if (分析表) 分析表.clear(); 
    else 分析表 = ss.insertSheet("進階銷售分析");

    var 列 = 1;

    // --- 第一部分：每月銷售對照表 ---
    分析表.getRange(列, 1).setValue("📅 每月銷售對照表").setFontSize(14).setFontWeight("bold");
    列++;
    分析表.getRange(列, 1, 1, 3).setValues([["月份", "總筆數", "總金額"]]).setBackground("#1a73e8").setFontColor("#fff").setFontWeight("bold");
    列++;

    var 月份排序 = Object.keys(每月統計).sort();
    
    月份排序.forEach(function(月) {
      var 統計 = 每月統計[月];
      分析表.getRange(列, 1, 1, 3).setValues([[月, 統計.筆數, 統計.總額]]);
      分析表.getRange(列, 3).setNumberFormat("#,##0");
      列++;
    });

    列 += 2;

    // --- 第二部分：每月前 5 名客戶 ---
    分析表.getRange(列, 1).setValue("🏆 每月最高前 5 名客戶").setFontSize(14).setFontWeight("bold");
    列++;

    月份排序.forEach(function(月) {
      分析表.getRange(列, 1).setValue("【" + 月 + "】").setFontWeight("bold");
      列++;
      分析表.getRange(列, 1, 1, 5).setValues([["排名", "客戶", "負責業務", "商品", "消費總額"]]).setBackground("#34a853").setFontColor("#fff").setFontWeight("bold");
      列++;

      var 月份客戶 = 每月統計[月].客戶資料;
      var 月客戶排序 = Object.keys(月份客戶).sort(function(a, b) {
        return 月份客戶[b].總額 - 月份客戶[a].總額;
      });

      var 排名 = 1;
      for (var i = 0; i < Math.min(5, 月客戶排序.length); i++) {
        var 客戶名 = 月客戶排序[i];
        var 資料 = 月份客戶[客戶名];
        var 業務字串 = Object.keys(資料.業務清單).join(", ");
        var 商品字串 = Object.keys(資料.商品清單).join(", ");
        
        分析表.getRange(列, 1, 1, 5).setValues([[排名, 客戶名, 業務字串, 商品字串, 資料.總額]]);
        分析表.getRange(列, 5).setNumberFormat("#,##0");
        列++;
        排名++;
      }
      列++; // 月與月之間空一列
    });

    // --- 第三部分：年度前 10 名客戶 ---
    分析表.getRange(列, 1).setValue("🌟 年度總排行：前 10 名客戶").setFontSize(14).setFontWeight("bold");
    列++;
    分析表.getRange(列, 1, 1, 5).setValues([["排名", "客戶", "負責業務", "商品", "消費總額"]]).setBackground("#fbbc04").setFontColor("#000").setFontWeight("bold");
    列++;

    var 年客戶排序 = Object.keys(年度客戶).sort(function(a, b) {
      return 年度客戶[b].總額 - 年度客戶[a].總額;
    });

    var 總排名 = 1;
    for (var j = 0; j < Math.min(10, 年客戶排序.length); j++) {
      var 客戶名 = 年客戶排序[j];
      var 資料 = 年度客戶[客戶名];
      var 業務字串 = Object.keys(資料.業務清單).join(", ");
      var 商品字串 = Object.keys(資料.商品清單).join(", ");
      
      分析表.getRange(列, 1, 1, 5).setValues([[總排名, 客戶名, 業務字串, 商品字串, 資料.總額]]);
      分析表.getRange(列, 5).setNumberFormat("#,##0");
      列++;
      總排名++;
    }

    // 自動調整欄寬
    for (var c = 1; c <= 5; c++) 分析表.autoResizeColumn(c);

    SpreadsheetApp.getUi().alert("✅ 每月進階銷售分析與排名報告已生成！\n請查看「進階銷售分析」工作表。");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 自訂選單
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📚 Session 6 工具")
    .addItem("📦 初始化銷售資料", "初始化銷售資料")
    .addItem("🔍 篩選示範", "篩選示範")
    .addItem("📶 排序示範", "排序示範")
    .addSeparator()
    .addItem("📊 生成銷售摘要", "生成銷售摘要")
    .addItem("🧹 自動資料清理", "自動資料清理")
    .addItem("📤 匯出篩選結果", "匯出篩選結果")
    .addItem("📈 生成進階銷售分析", "生成進階銷售分析")
    .addToUi();
}
