// ============================================================
// Session 4：Ranges 與批次處理
// 日期：115/05/09（六）13:30~16:30
// 講師：林冠廷
// ============================================================
// 本課程涵蓋：
//   1. Ranges（儲存格範圍）存取與修改
//   2. 執行限制（批次處理最佳化）
//   3. 陣列 (Array) 基礎
//   4. 實作：批次讀寫資料，自動匯入範例
// ============================================================

// ============================================================
// 第一部分：Range 進階操作
// ============================================================

/**
 * Range 取得方式大全
 * 說明：示範各種 getRange 的用法
 */
function Range取得方式() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("庫存資料");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ 請先執行「初始化庫存資料」");
    return;
  }

  // 方式 1：A1 表示法
  var r1 = sheet.getRange("A1");
  Logger.log("A1 = " + r1.getValue());

  // 方式 2：A1 範圍表示法
  var r2 = sheet.getRange("A1:C3");
  Logger.log("A1:C3 = " + JSON.stringify(r2.getValues()));

  // 方式 3：數字表示法 getRange(列, 欄)
  var r3 = sheet.getRange(1, 1);  // 等同 A1
  Logger.log("(1,1) = " + r3.getValue());

  // 方式 4：數字範圍 getRange(起始列, 起始欄, 列數, 欄數)
  var r4 = sheet.getRange(2, 1, 5, 4);  // 從 A2 開始，5列4欄
  Logger.log("(2,1,5,4) = " + JSON.stringify(r4.getValues()));

  // 方式 5：整欄 / 整列
  var r5 = sheet.getRange("A:A");  // 整個 A 欄
  var r6 = sheet.getRange("2:2");  // 整個第 2 列
  Logger.log("第 2 列 = " + JSON.stringify(r6.getValues()));
}

/**
 * Range 修改操作
 * 說明：修改儲存格的值與屬性
 */
function Range修改操作() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("庫存資料");
  if (!sheet) return;

  // --- 單一儲存格操作 ---
  var cell = sheet.getRange("A1");
  cell.setValue("商品名稱");
  cell.setBackground("#e8f5e9");
  cell.setFontWeight("bold");
  cell.setFontSize(12);
  cell.setHorizontalAlignment("center");
  cell.setNote("這是商品名稱欄位");  // 加入附註

  // --- 範圍操作 ---
  var range = sheet.getRange("B2:B11");
  range.setNumberFormat("#,##0");       // 數字格式：千分位
  range.setHorizontalAlignment("right"); // 靠右

  // --- 使用公式 ---
  // F2 = D2 * E2（庫存量 × 單價）
  sheet.getRange("F2").setFormula("=D2*E2");

  // 批次設定公式（F2:F11）
  for (var i = 2; i <= 11; i++) {
    sheet.getRange("F" + i).setFormula("=D" + i + "*E" + i);
  }

  Logger.log("✅ Range 修改完成");
}

// ============================================================
// 第二部分：陣列 (Array) 基礎
// ============================================================

/**
 * 陣列基礎操作示範
 */
function 陣列基礎() {
  // --- 一維陣列 ---
  var 水果 = ["蘋果", "香蕉", "橘子", "葡萄", "芒果"];

  Logger.log("陣列長度：" + 水果.length);          // 5
  Logger.log("第一個元素：" + 水果[0]);             // 蘋果
  Logger.log("最後一個：" + 水果[水果.length - 1]); // 芒果

  // 新增元素
  水果.push("西瓜");
  Logger.log("push 後：" + 水果);            // 末尾加入

  // 刪除最後一個
  var 被移除 = 水果.pop();
  Logger.log("pop 移除：" + 被移除);          // 西瓜

  // 搜尋元素
  var 索引 = 水果.indexOf("橘子");
  Logger.log("橘子的索引：" + 索引);           // 2

  // --- 二維陣列（模擬試算表資料）---
  var 表格 = [
    ["姓名", "年齡", "部門"],     // 標題列
    ["王小明", 25, "業務部"],
    ["李小華", 30, "行銷部"],
    ["張美玲", 28, "人資部"]
  ];

  Logger.log("第 2 列第 1 欄：" + 表格[1][0]); // 王小明
  Logger.log("第 3 列第 3 欄：" + 表格[2][2]); // 行銷部

  // 遍歷二維陣列
  for (var i = 1; i < 表格.length; i++) {
    Logger.log(表格[i][0] + "，" + 表格[i][1] + " 歲，" + 表格[i][2]);
  }
}

// ============================================================
// 第三部分：批次處理最佳化
// ============================================================

/**
 * ❌ 不良示範：逐格讀寫（效能極差）
 * 說明：每次 getValue/setValue 都是一次 API 呼叫，
 *       大量操作時會非常慢且容易超時
 */
function 不良示範_逐格讀寫() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("庫存資料");
  if (!sheet) return;

  var 開始 = new Date().getTime();

  // ❌ 這樣做很慢！每次迴圈都會呼叫 API
  for (var i = 2; i <= 11; i++) {
    var 商品 = sheet.getRange(i, 1).getValue();  // API 呼叫 ×1
    var 單價 = sheet.getRange(i, 2).getValue();  // API 呼叫 ×1
    var 庫存 = sheet.getRange(i, 4).getValue();  // API 呼叫 ×1
    var 總值 = 單價 * 庫存;
    sheet.getRange(i, 6).setValue(總值);          // API 呼叫 ×1
    // 每列 4 次 API 呼叫，10 列 = 40 次！
  }

  var 結束 = new Date().getTime();
  Logger.log("❌ 逐格讀寫耗時：" + (結束 - 開始) + " 毫秒");
}

/**
 * ✅ 正確示範：批次讀寫（效能最佳）
 * 說明：一次讀取全部、處理完再一次寫回
 */
function 正確示範_批次讀寫() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("庫存資料");
  if (!sheet) return;

  var 開始 = new Date().getTime();

  // ✅ 一次讀取所有資料（1 次 API 呼叫）
  var 所有資料 = sheet.getDataRange().getValues();

  // 在記憶體中處理（0 次 API 呼叫）
  var 結果 = [];
  for (var i = 1; i < 所有資料.length; i++) {
    var 單價 = 所有資料[i][1];   // B 欄：單價
    var 庫存 = 所有資料[i][3];   // D 欄：庫存量
    var 總值 = 單價 * 庫存;
    結果.push([總值]);
  }

  // ✅ 一次寫入所有結果（1 次 API 呼叫）
  sheet.getRange(2, 6, 結果.length, 1).setValues(結果);

  var 結束 = new Date().getTime();
  Logger.log("✅ 批次讀寫耗時：" + (結束 - 開始) + " 毫秒");
  // 通常比逐格讀寫快 10 倍以上！
}

// ============================================================
// 第四部分：批次匯入外部資料
// ============================================================

/**
 * 自動匯入範例資料（模擬外部資料匯入）
 * 說明：示範如何將大量資料批次寫入試算表
 */
function 批次匯入資料() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 表名 = "匯入資料_" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd_HHmmss");
    var sheet = ss.insertSheet(表名);

    // 模擬產生 50 筆訂單資料
    var 標題 = [["訂單編號", "客戶名稱", "商品", "數量", "單價", "金額", "日期", "狀態"]];
    var 資料 = [];

    var 客戶列表 = ["台北公司", "新竹企業", "台中行銷", "高雄科技", "花蓮文創"];
    var 商品列表 = ["筆記型電腦", "印表機", "投影機", "平板電腦", "耳機", "滑鼠", "鍵盤"];
    var 狀態列表 = ["已出貨", "處理中", "已取消", "已完成"];

    for (var i = 0; i < 50; i++) {
      var 訂單編號 = "ORD-" + String(i + 1).padStart(4, "0");
      var 客戶 = 客戶列表[Math.floor(Math.random() * 客戶列表.length)];
      var 商品 = 商品列表[Math.floor(Math.random() * 商品列表.length)];
      var 數量 = Math.floor(Math.random() * 10) + 1;
      var 單價 = Math.floor(Math.random() * 50000) + 500;
      var 金額 = 數量 * 單價;
      var 日期 = new Date(2026, Math.floor(Math.random() * 12), Math.floor(Math.random() * 28) + 1);
      var 狀態 = 狀態列表[Math.floor(Math.random() * 狀態列表.length)];

      資料.push([訂單編號, 客戶, 商品, 數量, 單價, 金額, 日期, 狀態]);
    }

    // 批次寫入標題
    sheet.getRange(1, 1, 1, 8).setValues(標題);

    // 批次寫入資料（一次寫入 50 筆！）
    sheet.getRange(2, 1, 資料.length, 8).setValues(資料);

    // 格式化
    var 標題範圍 = sheet.getRange("A1:H1");
    標題範圍.setBackground("#ea4335");
    標題範圍.setFontColor("#ffffff");
    標題範圍.setFontWeight("bold");
    標題範圍.setHorizontalAlignment("center");

    // 數字格式
    sheet.getRange("E2:F51").setNumberFormat("#,##0");
    sheet.getRange("G2:G51").setNumberFormat("yyyy/mm/dd");

    sheet.setFrozenRows(1);
    for (var j = 1; j <= 8; j++) {
      sheet.autoResizeColumn(j);
    }

    Logger.log("✅ 已匯入 " + 資料.length + " 筆訂單到「" + 表名 + "」");
    SpreadsheetApp.getUi().alert("✅ 已匯入 " + 資料.length + " 筆訂單！\n工作表：" + 表名);

  } catch (錯誤) {
    Logger.log("❌ 匯入錯誤：" + 錯誤.message);
  }
}

/**
 * 批次更新庫存（讀取→計算→寫回）
 * 說明：一次讀取所有庫存，計算庫存總值後批次寫回，並在低庫存時發送預警信
 */
function 批次更新庫存() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("庫存資料");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("❌ 找不到「庫存資料」");
      return;
    }

    // Step 1：一次讀取全部資料
    var 資料 = sheet.getDataRange().getValues();

    // Step 2：在記憶體中計算
    var 庫存總值列 = [];
    var 庫存狀態列 = [];
    var 總庫存值 = 0;
    
    // 新增：用來收集低於安全庫存的商品清單
    var 預警清單 = [];

    for (var i = 1; i < 資料.length; i++) {
      var 商品名稱 = 資料[i][0]; // A: 商品名稱
      var 單價 = 資料[i][1];     // B: 單價
      var 安全庫存 = 資料[i][2]; // C: 安全庫存量
      var 目前庫存 = 資料[i][3]; // D: 目前庫存
      var 庫存總值 = 單價 * 目前庫存;

      庫存總值列.push([庫存總值]);

      // 判斷庫存狀態，並在低於安全庫存時加入預警清單
      var 狀態;
      if (目前庫存 <= 0) {
        狀態 = "🔴 缺貨";
        預警清單.push("🔴 " + 商品名稱 + " (目前庫存: " + 目前庫存 + " / 安全庫存: " + 安全庫存 + ")");
      } else if (目前庫存 < 安全庫存) {
        狀態 = "🟡 低庫存";
        預警清單.push("🟡 " + 商品名稱 + " (目前庫存: " + 目前庫存 + " / 安全庫存: " + 安全庫存 + ")");
      } else {
        狀態 = "🟢 正常";
      }
      庫存狀態列.push([狀態]);

      總庫存值 += 庫存總值;
    }

    // Step 3：一次寫回所有結果
    var 資料筆數 = 資料.length - 1;
    sheet.getRange(2, 6, 資料筆數, 1).setValues(庫存總值列);  // F 欄：庫存總值
    sheet.getRange(2, 7, 資料筆數, 1).setValues(庫存狀態列);  // G 欄：狀態

    // 數字格式
    sheet.getRange(2, 6, 資料筆數, 1).setNumberFormat("#,##0");

    // 在最後寫入總計
    var 總計列 = 資料.length + 1;
    sheet.getRange(總計列, 5).setValue("總庫存值：");
    sheet.getRange(總計列, 5).setFontWeight("bold");
    sheet.getRange(總計列, 6).setValue(總庫存值);
    sheet.getRange(總計列, 6).setNumberFormat("#,##0").setFontWeight("bold").setBackground("#e8f5e9");

    // Step 4：判斷是否需要發送預警信
    var 提示訊息 = "✅ 庫存更新完成！\n總庫存值：NT$ " + 總庫存值.toLocaleString();
    
    if (預警清單.length > 0) {
      // 修改為指定的收件人信箱 (可使用逗號分隔多個信箱)
      var 接收信箱 = "jenniferchenatwork@gmail.com"; 
      var 信件主旨 = "⚠️ 庫存預警通知：有商品低於安全庫存！";
      var 信件內容 = "系統偵測到以下商品庫存不足，請盡快安排補貨：\n\n" + 
                     預警清單.join("\n") + 
                     "\n\n請至系統確認詳細資訊。";
      
      MailApp.sendEmail(接收信箱, 信件主旨, 信件內容);
      
      Logger.log("📧 已發送庫存預警信給：" + 接收信箱);
      提示訊息 += "\n\n📧 偵測到缺貨商品，已發送預警信至 " + 接收信箱 + "！";
    }

    Logger.log("✅ 庫存更新完成！總庫存值：NT$" + 總庫存值.toLocaleString());
    
    // 加上 try...catch：因為如果是被「排程」自動觸發執行，背景不會有試算表 UI 畫面可以顯示 alert
    // 若沒有 try...catch 捕捉，排程執行時遇到 getUi().alert() 會發生錯誤而中斷。
    try {
      SpreadsheetApp.getUi().alert(提示訊息);
    } catch (e) {
      Logger.log("背景排程執行完畢。");
    }

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 排程自動化
// ============================================================

/**
 * 建立排程：每週一早上自動執行「批次更新庫存」
 * 您可以手動點擊執行此函數，或是從試算表選單中點擊
 */
function 設定每週一發送排程() {
  // 1. 先清除舊的同名排程，避免重複建立導致發送多封信
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "批次更新庫存") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 2. 建立新的排程：每週一早上 8:00~9:00 之間自動執行
  ScriptApp.newTrigger("批次更新庫存")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
    
  Logger.log("✅ 已成功設定每週一早上 8:00 自動檢查庫存的排程！");
  
  try {
    SpreadsheetApp.getUi().alert("✅ 已成功設定每週一早上 8:00 自動檢查庫存的排程！\n系統每週一將會自動在背景執行並發送預警信。");
  } catch (e) {}
}

// ============================================================
// 初始化範例資料
// ============================================================

/**
 * 建立「庫存資料」工作表
 */
function 初始化庫存資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("庫存資料");

  if (!sheet) {
    sheet = ss.insertSheet("庫存資料");
  } else {
    sheet.clear();
  }

  var 標題 = [["商品名稱", "單價", "安全庫存量", "目前庫存", "單位", "庫存總值", "狀態"]];
  var 資料 = [
    ["A4 影印紙", 150, 50, 120, "包", "", ""],
    ["原子筆", 15, 100, 45, "支", "", ""],
    ["資料夾", 25, 80, 200, "個", "", ""],
    ["釘書針", 35, 30, 8, "盒", "", ""],
    ["便條紙", 20, 60, 0, "本", "", ""],
    ["膠帶", 30, 40, 55, "捲", "", ""],
    ["白板筆", 45, 20, 15, "支", "", ""],
    ["計算機", 350, 10, 12, "台", "", ""],
    ["印表機碳粉", 1200, 5, 3, "個", "", ""],
    ["滑鼠墊", 80, 15, 25, "個", "", ""]
  ];

  sheet.getRange(1, 1, 1, 7).setValues(標題);
  sheet.getRange(2, 1, 資料.length, 7).setValues(資料);

  var 標題範圍 = sheet.getRange("A1:G1");
  標題範圍.setBackground("#ff6d01");
  標題範圍.setFontColor("#ffffff");
  標題範圍.setFontWeight("bold");
  標題範圍.setHorizontalAlignment("center");

  sheet.getRange("B2:B11").setNumberFormat("#,##0");
  sheet.setFrozenRows(1);
  for (var i = 1; i <= 7; i++) sheet.autoResizeColumn(i);

  Logger.log("✅ 庫存資料已建立！");
  SpreadsheetApp.getUi().alert("✅ 庫存資料已建立！\n請執行「批次更新庫存」計算結果。");
}

// ============================================================
// 自訂選單
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📚 Session 4 工具")
    .addItem("📦 初始化庫存資料", "初始化庫存資料")
    .addItem("📊 批次更新庫存", "批次更新庫存")
    .addItem("📥 批次匯入訂單資料", "批次匯入資料")
    .addSeparator()
    .addItem("📝 Range 取得方式", "Range取得方式")
    .addItem("🎨 Range 修改操作", "Range修改操作")
    .addItem("📚 陣列基礎", "陣列基礎")
    .addSeparator()
    .addItem("❌ 不良示範：逐格讀寫", "不良示範_逐格讀寫")
    .addItem("✅ 正確示範：批次讀寫", "正確示範_批次讀寫")
    .addSeparator()
    .addItem("⏰ 設定每週一發送排程", "設定每週一發送排程")
    .addToUi();
}