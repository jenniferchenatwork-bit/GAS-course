/**
 * ============================================================================
 * Session 11：新產品營業額與毛利額彙整系統 (實戰專案)
 * 說明：將當月新產品營業額資料彙整到銷售毛利額工作表，具備欄位動態對齊、
 *       缺失數值自動計算、批次處理最佳化，以及自動化儲存格格式設定。
 * 
 * 來源工作表名稱 (預設)："26'04" (可依實際月份修改，例如 "26'05")
 * 目的工作表名稱 (預設)："當月新產品銷售資料"
 * 篩選規則：僅保留「布號設定日期」在 4 月份的產品資訊。
 * ============================================================================
 */

function 彙整當月新產品銷售資料() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // ==========================================
    // 1. 設定工作表名稱（可依需求修改）
    // ==========================================
    var SOURCE_SHEET_NAME = "26'04"; // 4月份新產品營業額工作表
    var TARGET_SHEET_NAME = "當月新產品銷售資料"; // 彙整目標工作表
    
    var sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
    if (!sourceSheet) {
      SpreadsheetApp.getUi().alert("❌ 找不到來源工作表：「" + SOURCE_SHEET_NAME + "」，請確認工作表名稱是否正確。");
      return;
    }
    
    // ==========================================
    // 2. 確保目標工作表存在，若不存在則自動建立並套用標題與樣式
    // ==========================================
    var targetSheet = ss.getSheetByName(TARGET_SHEET_NAME);
    var isNewTargetSheet = false;
    
    // 定義標準的彙整欄位順序 (包含日期、類別、布號、數量、金額、售價、成本、毛利率)
    var targetHeaders = ["日期", "類別", "布號", "數量", "金額", "售價", "成本", "毛利率"];
    
    if (!targetSheet) {
      targetSheet = ss.insertSheet(TARGET_SHEET_NAME);
      isNewTargetSheet = true;
      Logger.log("ℹ️ 目標工作表「" + TARGET_SHEET_NAME + "」不存在，已自動建立。");
    }
    
    // ==========================================
    // 3. 讀取來源資料與建立動態欄位對照表 (Header Aliases)
    // ==========================================
    var sourceDataRange = sourceSheet.getDataRange();
    var sourceValues = sourceDataRange.getValues();
    
    if (sourceValues.length <= 1) {
      SpreadsheetApp.getUi().alert("⚠️ 來源工作表「" + SOURCE_SHEET_NAME + "」中除了標題列外無其他資料！");
      return;
    }
    
    var sourceHeaders = sourceValues[0];
    
    // 定義欄位別名系統，增加程式對不同表格欄位名稱的容錯度與智慧對齊
    var headerAliases = {
      "日期": ["布號設定日期", "設定日期", "建立日期", "開發日期", "日期", "Date", "Setup Date"],
      "類別": ["類別", "產品類別", "Category", "分類", "產品分類"],
      "布號": ["布號", "編號", "商品代碼", "布種", "品號", "Fabric"],
      "數量": ["數量", "件數", "銷量", "銷售數量", "Quantity"],
      "金額": ["金額", "營業額", "銷貨收入", "銷售額", "Revenue", "Amount"],
      "毛利率": ["毛利率", "利潤率", "毛利率%", "Margin", "Gross Profit Margin"],
      "售價": ["售價", "單價", "售單價", "Price", "Unit Price"],
      "成本": ["成本價", "單位成本", "成本單價", "成本", "Cost Price", "Cost"]
    };
    
    // 尋找來源工作表欄位索引
    var sourceIndexes = {};
    for (var key in headerAliases) {
      sourceIndexes[key] = -1; // 預設 -1 代表沒找到
      var aliases = headerAliases[key];
      
      for (var i = 0; i < sourceHeaders.length; i++) {
        var cleanHeader = String(sourceHeaders[i]).trim();
        if (aliases.indexOf(cleanHeader) !== -1) {
          sourceIndexes[key] = i;
          break;
        }
      }
    }
    
    // 輸出欄位查找日誌，方便除錯
    Logger.log("🔍 欄位搜尋結果：");
    for (var colName in sourceIndexes) {
      Logger.log("   - " + colName + " 欄位索引：" + (sourceIndexes[colName] !== -1 ? "第 " + (sourceIndexes[colName] + 1) + " 欄" : "未找到"));
    }
    
    // 檢查關鍵篩選欄位是否存在
    if (sourceIndexes["日期"] === -1) {
      SpreadsheetApp.getUi().alert("❌ 來源工作表中找不到「布號設定日期」或相關的日期欄位，無法進行 4 月份產品之篩選！\n請確認欄位名稱。");
      return;
    }
    
    // ==========================================
    // 4. 讀取目標工作表的欄位結構（若為舊表則動態對齊欄位）
    // ==========================================
    var currentTargetHeaders = [];
    if (isNewTargetSheet) {
      currentTargetHeaders = targetHeaders;
      // 寫入全新工作表標題列
      targetSheet.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders]);
      // 美化全新標題列
      var headerRange = targetSheet.getRange(1, 1, 1, targetHeaders.length);
      headerRange.setBackground("#1a365d") // 深藍高質感
                 .setFontColor("#ffffff")
                 .setFontWeight("bold")
                 .setHorizontalAlignment("center")
                 .setFontFamily("Microsoft JhengHei");
      targetSheet.setFrozenRows(1);
    } else {
      var targetRange = targetSheet.getDataRange();
      var targetValues = targetRange.getValues();
      currentTargetHeaders = targetValues[0];
      
      // 如果目標表是空的（連標題都沒有），寫入預設標題
      if (currentTargetHeaders.length === 0 || (currentTargetHeaders.length === 1 && currentTargetHeaders[0] === "")) {
        currentTargetHeaders = targetHeaders;
        targetSheet.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders]);
        var headerRange = targetSheet.getRange(1, 1, 1, targetHeaders.length);
        headerRange.setBackground("#1a365d")
                   .setFontColor("#ffffff")
                   .setFontWeight("bold")
                   .setHorizontalAlignment("center")
                   .setFontFamily("Microsoft JhengHei");
        targetSheet.setFrozenRows(1);
      }
    }
    
    // 找出目標工作表每個標題對應的欄位索引
    var targetColMapping = {};
    for (var t = 0; t < currentTargetHeaders.length; t++) {
      targetColMapping[String(currentTargetHeaders[t]).trim()] = t;
    }
    
    // ==========================================
    // 5. 批次處理並智慧計算與篩選數值
    // ==========================================
    var rowsToInsert = [];
    var processedCount = 0;
    var skippedCount = 0;
    
    for (var r = 1; r < sourceValues.length; r++) {
      var sourceRow = sourceValues[r];
      
      // 取得基礎關鍵欄位，如果類別與布號皆為空，則視為空行跳過
      var category = getValueByHeader(sourceRow, sourceIndexes["類別"]);
      var fabricNo = getValueByHeader(sourceRow, sourceIndexes["布號"]);
      
      if (!category && !fabricNo) {
        continue; // 跳過空行
      }
      
      // --- 篩選布號設定日期在 4 月份的產品 ---
      var setupDateRaw = getValueByHeader(sourceRow, sourceIndexes["日期"]);
      var setupDate = null;
      
      if (setupDateRaw instanceof Date) {
        setupDate = setupDateRaw;
      } else if (setupDateRaw) {
        setupDate = new Date(setupDateRaw);
      }
      
      // 若無法解析日期，或日期月份不為 4 月 (getMonth() === 3)，則過濾排除
      if (!setupDate || isNaN(setupDate.getTime()) || setupDate.getMonth() !== 3) {
        skippedCount++;
        continue;
      }
      
      // 讀取各欄位之數值 (若無則先預設 null)
      var qty = parseNumber(getValueByHeader(sourceRow, sourceIndexes["數量"]));
      var amount = parseNumber(getValueByHeader(sourceRow, sourceIndexes["金額"]));
      var margin = parseNumber(getValueByHeader(sourceRow, sourceIndexes["毛利率"]));
      var price = parseNumber(getValueByHeader(sourceRow, sourceIndexes["售價"]));
      var costPrice = parseNumber(getValueByHeader(sourceRow, sourceIndexes["成本"]));
      
      // --- 智慧公式計算 (以記憶體高速運算填補遺漏欄位) ---
      
      // A. 售價 (Unit Price)
      if (price === null && amount !== null && qty > 0) {
        price = amount / qty;
      }
      
      // B. 成本價 / 成本 (Unit Cost)
      if (costPrice === null) {
        if (price !== null && margin !== null) {
          costPrice = price * (1 - margin);
        }
      }
      
      // C. 金額 (Amount / Revenue)
      if (amount === null) {
        if (price !== null && qty !== null) {
          amount = price * qty;
        }
      }
      
      // D. 毛利率 (Gross Profit Margin)
      if (margin === null) {
        if (price > 0 && costPrice !== null) {
          margin = (price - costPrice) / price;
        }
      }
      
      // 格式化日期字串 (yyyy-MM-dd)
      var dateStr = Utilities.formatDate(setupDate, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
      
      // 建立對應於目標工作表欄位順序的列資料
      var outputRow = new Array(currentTargetHeaders.length).fill("");
      
      // 將資料對齊寫入對應的目標欄位中
      mapValueToTarget(outputRow, targetColMapping, "日期", dateStr);
      mapValueToTarget(outputRow, targetColMapping, "類別", category);
      mapValueToTarget(outputRow, targetColMapping, "布號", fabricNo);
      mapValueToTarget(outputRow, targetColMapping, "數量", qty);
      mapValueToTarget(outputRow, targetColMapping, "金額", amount);
      mapValueToTarget(outputRow, targetColMapping, "售價", price);
      mapValueToTarget(outputRow, targetColMapping, "成本", costPrice);
      mapValueToTarget(outputRow, targetColMapping, "毛利率", margin);
      
      rowsToInsert.push(outputRow);
      processedCount++;
    }
    
    // ==========================================
    // 6. 批次寫入目標工作表 (Batch Writing)
    // ==========================================
    if (rowsToInsert.length > 0) {
      var startRow = targetSheet.getLastRow() + 1;
      var numRows = rowsToInsert.length;
      var numCols = currentTargetHeaders.length;
      
      // 一次性寫入所有資料
      var writeRange = targetSheet.getRange(startRow, 1, numRows, numCols);
      writeRange.setValues(rowsToInsert);
      
      // ==========================================
      // 7. 儲存格格式與美化設定 (Styling)
      // ==========================================
      
      for (var col = 0; col < currentTargetHeaders.length; col++) {
        var hName = String(currentTargetHeaders[col]).trim();
        var colRange = targetSheet.getRange(startRow, col + 1, numRows, 1);
        
        // 數字與英文採用 Calibri，中文字採用微軟正黑體
        colRange.setFontFamily("Calibri").setFontSize(11);
        
        if (hName === "數量") {
          colRange.setNumberFormat("#,##0"); // 數量千分位格式
          colRange.setHorizontalAlignment("right");
        } else if (hName === "金額") {
          colRange.setNumberFormat("$#,##0"); // 貨幣千分位格式
          colRange.setHorizontalAlignment("right");
        } else if (hName === "售價" || hName === "成本") {
          colRange.setNumberFormat("$#,##0.00"); // 價格保留兩位小數
          colRange.setHorizontalAlignment("right");
        } else if (hName === "毛利率") {
          colRange.setNumberFormat("0.0%"); // 百分比顯示
          colRange.setHorizontalAlignment("right");
          
          // 條件式格式化輔助：對低毛利與高毛利產品做顏色識別
          var rawValues = colRange.getValues();
          for (var ri = 0; ri < rawValues.length; ri++) {
            var val = parseNumber(rawValues[ri][0]);
            if (val !== null && val < 0.15) {
              // 毛利低於 15% 標記淡紅色
              targetSheet.getRange(startRow + ri, col + 1).setBackground("#ffebee").setFontColor("#c62828");
            } else if (val !== null && val > 0.40) {
              // 毛利高於 40% 標記淡綠色
              targetSheet.getRange(startRow + ri, col + 1).setBackground("#e8f5e9").setFontColor("#2e7d32");
            }
          }
        } else {
          colRange.setFontFamily("Microsoft JhengHei");
          colRange.setHorizontalAlignment("center");
        }
      }
      
      // 自動調整欄寬
      for (var c = 1; c <= currentTargetHeaders.length; c++) {
        targetSheet.autoResizeColumn(c);
      }
      
      var msg = "🎉 彙整完成！\n" + 
                "📊 來源工作表：「" + SOURCE_SHEET_NAME + "」\n" + 
                "📥 成功匯入 「" + processedCount + "」 筆 4 月份新產品銷售資料至 「" + TARGET_SHEET_NAME + "」！\n" +
                "🚫 已自動過濾排除 「" + skippedCount + "」 筆非 4 月份之產品資料。";
      
      Logger.log("✅ " + msg);
      SpreadsheetApp.getUi().alert(msg);
      
    } else {
      SpreadsheetApp.getUi().alert("ℹ️ 未偵測到符合 4 月份之新產品資料。\n（已過濾排除 " + skippedCount + " 筆非 4 月份資料）");
    }
    
  } catch (error) {
    Logger.log("❌ 執行發生錯誤：" + error.message);
    SpreadsheetApp.getUi().alert("❌ 彙整失敗，錯誤原因：\n" + error.message);
  }
}

/**
 * 輔助函式：根據欄位索引安全取得陣列值
 */
function getValueByHeader(rowArray, index) {
  if (index !== undefined && index >= 0 && index < rowArray.length) {
    return rowArray[index];
  }
  return null;
}

/**
 * 輔助函式：將解析過的數值安全寫入對應的目標陣列位置
 */
function mapValueToTarget(outputRow, targetColMapping, headerName, value) {
  if (headerName in targetColMapping) {
    var targetIndex = targetColMapping[headerName];
    outputRow[targetIndex] = value !== null ? value : "";
  }
}

/**
 * 輔助函式：將文字數值轉換為 float 浮點數，防呆過濾空白與符號
 */
function parseNumber(value) {
  if (value === "" || value === null || value === undefined) {
    return null;
  }
  if (typeof value === "number") {
    return value;
  }
  var cleaned = String(value).replace(/[\$,\s%]/g, "");
  var num = parseFloat(cleaned);
  
  if (isNaN(num)) {
    return null;
  }
  
  if (String(value).indexOf("%") !== -1) {
    return num / 100;
  }
  
  return num;
}

/**
 * 自動註冊選單事件
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚀 新產品彙整工具")
    .addItem("📊 彙整當月新產品資料", "彙整當月新產品銷售資料")
    .addToUi();
}
