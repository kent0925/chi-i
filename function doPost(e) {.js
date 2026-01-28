// --- 系統設定與初始化 API ---
function doGet(e) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // [新增] 統編查詢 API (GET)
    if (e.parameter.action === 'lookup') {
        var taxId = e.parameter.taxId;
        if (!taxId) return ContentService.createTextOutput(JSON.stringify({ error: "No ID" })).setMimeType(ContentService.MimeType.JSON);

        // [新增] 優先查詢歷史紀錄
        var historyName = checkVendorHistory(taxId);
        if (historyName) {
            return ContentService.createTextOutput(JSON.stringify({
                name: historyName,
                taxId: taxId,
                source: 'history',
                error: ""
            })).setMimeType(ContentService.MimeType.JSON);
        }

        var filterStr = "Business_Accounting_NO eq " + taxId;
        // 1. 公司登記, 2. 商業登記, 3. 分公司登記
        var urls = [
            "https://data.gcis.nat.gov.tw/od/data/api/9D17AE0D-09B5-4732-A8F4-81ADED04B679?$format=json&$filter=" + encodeURIComponent(filterStr),
            "https://data.gcis.nat.gov.tw/od/data/api/5F6402A4-EB18-440A-B010-903760CA1663?$format=json&$filter=" + encodeURIComponent(filterStr),
            "https://data.gcis.nat.gov.tw/od/data/api/F7BC70B5-75A4-4710-97A4-AB05DDC69199?$format=json&$filter=" + encodeURIComponent(filterStr)
        ];

        var resultName = "";
        var lastError = "";

        for (var i = 0; i < urls.length; i++) {
            try {
                var resp = UrlFetchApp.fetch(urls[i], { muteHttpExceptions: true });
                var json = JSON.parse(resp.getContentText());
                if (json && json.length > 0) {
                    // 取得屬性名稱 (不同 API 屬性名可能略有不同：Company_Name, Business_Name, Branch_Name)
                    var firstRow = json[0];
                    resultName = firstRow.Company_Name || firstRow.Business_Name || firstRow.Branch_Name || firstRow.Company_Name_A || "";
                    if (resultName) break;
                }
            } catch (err) {
                lastError = err.toString();
            }
        }

        return ContentService.createTextOutput(JSON.stringify({
            name: resultName,
            taxId: taxId,
            error: resultName ? "" : (lastError || "查無登記資料")
        })).setMimeType(ContentService.MimeType.JSON);
    }

    // [新增] 取得銀行代碼列表 (GET) - 動態抓取開放資料
    if (e.parameter.action === 'getBanks') {
        try {
            var url = "https://raw.githubusercontent.com/nczz/taiwan-banks-list/main/banks_sort_by_codes.json";
            var resp = UrlFetchApp.fetch(url);
            var fullList = JSON.parse(resp.getContentText());
            var bankMap = {};
            fullList.forEach(function (item) {
                // bank_code 可能重複，我們只取唯一代碼對應的名稱
                if (!bankMap[item.bank_code]) {
                    bankMap[item.bank_code] = item.name;
                }
            });
            var banks = Object.keys(bankMap).sort().map(function (code) {
                return { code: code, name: bankMap[code] };
            });
            return ContentService.createTextOutput(JSON.stringify(banks)).setMimeType(ContentService.MimeType.JSON);
        } catch (err) {
            // 回退機制：若抓取失敗，顯示基本名單
            var fallback = [
                { code: "004", name: "臺灣銀行" }, { code: "005", name: "土地銀行" },
                { code: "006", name: "合作金庫" }, { code: "007", name: "第一銀行" },
                { code: "008", name: "華南銀行" }, { code: "009", name: "彰化銀行" },
                { code: "012", name: "台北富邦" }, { code: "013", name: "國泰世華" },
                { code: "017", name: "兆豐銀行" }, { code: "700", name: "中華郵政" },
                { code: "807", name: "永豐銀行" }, { code: "808", name: "玉山銀行" },
                { code: "812", name: "台新銀行" }, { code: "822", name: "中國信託" }
            ];
            return ContentService.createTextOutput(JSON.stringify(fallback)).setMimeType(ContentService.MimeType.JSON);
        }
    }

    // [新增] 取得全局財務概況 (GET)
    if (e.parameter.action === 'getGlobalFinancials') {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var quoteData = ss.getSheetByName("報價紀錄") ? ss.getSheetByName("報價紀錄").getDataRange().getValues() : [];
        var costData = ss.getSheetByName("成本紀錄") ? ss.getSheetByName("成本紀錄").getDataRange().getValues() : [];
        var accData = ss.getSheetByName("會計紀錄") ? ss.getSheetByName("會計紀錄").getDataRange().getValues() : [];

        var projectMap = {};

        // 1. 報價數據 (Receivable)
        for (var i = 1; i < quoteData.length; i++) {
            var row = quoteData[i];
            var pid = row[5]; // Project Name
            if (!pid) continue;
            if (!projectMap[pid]) projectMap[pid] = { name: pid, recNoTax: 0, payNoTax: 0, recGrand: 0, payGrand: 0, received: 0, paid: 0 };
            projectMap[pid].recNoTax += parseFloat(row[11] || 0); // Subtotal No Tax
            projectMap[pid].recGrand += parseFloat(row[13] || 0); // GrandTotal
        }

        // 2. 成本數據 (Payable)
        for (var i = 1; i < costData.length; i++) {
            var row = costData[i];
            var pid = row[4]; // Project Name
            if (!pid) continue;
            if (!projectMap[pid]) projectMap[pid] = { name: pid, recNoTax: 0, payNoTax: 0, recGrand: 0, payGrand: 0, received: 0, paid: 0 };
            projectMap[pid].payNoTax += parseFloat(row[12] || 0); // Subtotal No Tax
            projectMap[pid].payGrand += parseFloat(row[14] || 0); // GrandTotal
        }

        // 3. 會計數據 (Actual Received/Paid)
        for (var i = 1; i < accData.length; i++) {
            var row = accData[i];
            var pid = row[4]; // Project Name
            var amt = parseFloat(row[5] || 0);
            if (!pid || !projectMap[pid]) continue;

            // 從 JSON 中取得 accType
            try {
                var json = JSON.parse(row[7]);
                var type = json.accType; // receivable / payable
                if (type === 'receivable') {
                    projectMap[pid].received += amt;
                } else if (type === 'payable') {
                    projectMap[pid].paid += amt;
                }
            } catch (e) { }
        }

        var result = Object.values(projectMap).filter(p => p.recGrand > 0 || p.payGrand > 0);
        return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }

    // [新增] 取得專案財務資料 (GET) for Bonus & Payable Validation
    if (e.parameter.action === 'getProjectFinancials') {
        var projectId = e.parameter.projectId;
        if (!projectId) return ContentService.createTextOutput(JSON.stringify({ error: "No ProjectID" })).setMimeType(ContentService.MimeType.JSON);

        // 1. 取得報價總金額 (Receivable Total) & 客戶
        var quoteSheet = ss.getSheetByName("報價紀錄");
        var quoteData = quoteSheet ? quoteSheet.getDataRange().getValues() : [];
        var totalReceivable = 0;
        var paidAmount = 0;
        var customerName = "";
        var salesPerson = "";

        // 假設 "報價紀錄" 結構: [..., Customer(E), Project(F), ..., GrandTotal(N), Status(R), JSON(S)] (based on doPost logic)
        // Check doPost: rowData = [timestamp, orderId, version, date, customer, project, internalId, sales, mobile, lineId, lineName, subNoTax, tax, grandTotal, ...]
        // Index: 0, 1, 2, 3, 4(Cust), 5(Proj), 6, 7, 8, 9, 10, 11(Sub), 12(Tax), 13(Grand)

        for (var i = 1; i < quoteData.length; i++) {
            var row = quoteData[i];
            // Match Project Name (Col F, Index 5)
            if (row[5] === projectId) {
                totalReceivable += parseFloat(row[11] || 0);
                customerName = row[4];
                salesPerson = row[7] || "";
            }
        }

        // 2. 取得成本與發票 (Payable Total & Invoices)
        var costSheet = ss.getSheetByName("成本紀錄");
        var costData = costSheet ? costSheet.getDataRange().getValues() : [];
        var totalPayable = 0;
        var invoices = [];
        var vendors = [];
        var workers = [];
        var buyers = [];
        // Check doPost Cost: [.., Project(E), Vendor(F), TaxId(G), InvNo(H), ..., Sub(K), Tax(L), Grand(M), ..., JSON(R)]
        // Index: 0, 1, 2, 3, 4(Col E, Index 4), 5, 6, 7, ..., 11(Sub), 12(Tax), 13(Grand) -> Incorrect indexing in comments?
        // Let's re-verify doPost Cost Logic.
        // row = [time, ordId, ver, date, project, vendor, taxid, invNo, sales, mobile, worker, buyer, sub, tax, grand, summary, json, url, invUrl]
        // Index: 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18
        // Project is Index 4.
        // Subtotal is Index 12.

        for (var i = 1; i < costData.length; i++) {
            var row = costData[i];
            if (row[4] === projectId) {
                totalPayable += parseFloat(row[12] || 0);

                // Collect Vendor
                var vName = row[5];
                if (vName && vendors.indexOf(vName) === -1) vendors.push(vName);

                // Collect Personnel
                var wStr = row[10] || ""; // Workers list (string)
                wStr.split(',').forEach(function (p) { if (p && workers.indexOf(p.trim()) === -1) workers.push(p.trim()); });

                var bStr = row[11] || ""; // Buyers list
                bStr.split(',').forEach(function (p) { if (p && buyers.indexOf(p.trim()) === -1) buyers.push(p.trim()); });

                // Collect Invoice Info for Validation
                var invNo = row[7];
                if (invNo) {
                    invoices.push({
                        no: invNo,
                        vendor: vName,
                        amount: row[14],
                        url: row[18]
                    });
                }
            }
        }

        return ContentService.createTextOutput(JSON.stringify({
            customer: customerName,
            sales: salesPerson,
            receivable: totalReceivable,
            payable: totalPayable,
            invoices: invoices,
            vendors: vendors,
            workers: workers,
            buyers: buyers
        })).setMimeType(ContentService.MimeType.JSON);
    }

    // 1. 取得或建立「系統設定」分頁
    var configSheet = ss.getSheetByName("系統設定");
    if (!configSheet) {
        configSheet = ss.insertSheet("系統設定");
        configSheet.appendRow(["人員名單(共用)", "單位清單", "廠商名單"]); // 標題修正
        configSheet.getRange(2, 1).setValue("王小明");
        configSheet.getRange(2, 2).setValue("式");
        configSheet.getRange(2, 3).setValue("廠商A");
    }

    // 2. 讀取清單 (共用人員名單)
    var lastRow = configSheet.getLastRow();
    var sales = [], units = [], vendors = [];

    if (lastRow > 1) {
        // 讀取 A~C 欄
        var data = configSheet.getRange(2, 1, lastRow - 1, 3).getValues();
        data.forEach(function (row) {
            if (row[0]) sales.push(row[0]);
            if (row[1]) units.push(row[1]);
            if (row[2]) vendors.push(row[2]); // C欄: 廠商
        });

        // [新增] 移除重複項目 (Deduplicate)
        sales = [...new Set(sales)];
        units = [...new Set(units)];
        vendors = [...new Set(vendors)];
    } else {
        sales = ["業務A"]; units = ["式"]; vendors = ["廠商A"];
    }

    // 3. 工務與採購共用 Sales 名單
    var workers = sales;
    var buyers = sales;

    // 4. 計算今日專案ID序號 (格式: YYYYMMDDxxxx)
    var quoteSheet = ss.getSheetByName("報價紀錄");
    var nextSeq = "0001";
    var todayStr = Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd");

    // [新增] 歷史專案列表
    var projectList = [];

    if (quoteSheet && quoteSheet.getLastRow() > 1) {
        // [修改] 抓取更多欄位 (例如 20欄) 以確保能抓到 JSON
        var maxCols = 20;

        // [優化] 限制只抓取最後 150 筆資料，避免資料量過大導致 Timeout
        var lastRow = quoteSheet.getLastRow();
        var limit = 150;
        var startRow = Math.max(2, lastRow - limit + 1);
        var numRows = lastRow - startRow + 1;

        var qData = quoteSheet.getRange(startRow, 1, numRows, maxCols).getValues();

        qData.forEach(function (row) {
            // ID 計算邏輯: 專案ID在 G欄 (Index 6)
            var pid = row[6];
            if (pid && pid.toString().indexOf(todayStr) === 0 && pid.toString().length === 12) {
                var currentSeq = parseInt(pid.toString().substring(8));
                var potNext = currentSeq + 1;
                var potSeq = ("0000" + potNext).slice(-4);
                if (potSeq > nextSeq) nextSeq = potSeq;
            }

            // [新增] 智慧尋找 JSON 欄位
            // 因欄位結構可能變動 (舊資料在 O欄/Index 14, 新資料在 P欄/Index 15, 最新 Q欄/Index 16...等)
            // 我們從後往前找，找到第一個以 "{" 開頭的字串當作 JSON
            var jsonStr = "";
            for (var k = row.length - 1; k >= 10; k--) {
                var val = row[k];
                // [修改] 放寬判斷標準: 只要是字串且以 { 開頭就嘗試當作 JSON (不強制檢查 items)
                if (typeof val === 'string' && val.trim().indexOf('{') === 0) {
                    jsonStr = val;
                    break;
                }
            }
            if (!jsonStr) {
                // Fallback: 嘗試直接指定 Index (針對已知欄位位置嘗試)
                // 優先順序: R(17) -> Q(16) -> P(15) -> O(14)
                if (row[17] && row[17].toString().indexOf('{') === 0) jsonStr = row[17];
                else if (row[16] && row[16].toString().indexOf('{') === 0) jsonStr = row[16];
                else if (row[15] && row[15].toString().indexOf('{') === 0) jsonStr = row[15];
                else if (row[14] && row[14].toString().indexOf('{') === 0) jsonStr = row[14];
            }

            // 專案列表
            // [修正] 忽略沒有單號的空行或異常資料
            if (row[1]) {
                projectList.push({
                    id: row[1], // 單號
                    date: Utilities.formatDate(new Date(row[0]), "GMT+8", "yyyy-MM-dd"),
                    customer: row[4],
                    project: row[5],
                    sales: row[7],
                    salesMobile: row[8],
                    fax: row[15] || "", // [新增] 傳真 (Index 15/Column P)
                    status: row[16] || "", // [新增] 狀態 (Index 16/Column Q)
                    json: jsonStr // 使用智慧偵測的 JSON
                });
            }
        });
    }
    projectList.reverse();

    // [新增] 成本紀錄列表 (供 cost.html 編輯/查詢)
    var costSheet = ss.getSheetByName("成本紀錄");
    var costList = [];
    if (costSheet && costSheet.getLastRow() > 1) {
        var cData = costSheet.getRange(2, 1, costSheet.getLastRow() - 1, 19).getValues();
        cData.forEach(function (row) {
            // 只回傳有單號的紀錄
            if (row[1]) {
                costList.push({
                    id: row[1], // 成本單號
                    date: Utilities.formatDate(new Date(row[0]), "GMT+8", "yyyy-MM-dd"),
                    project: row[4], // 來源專案名稱
                    vendor: row[5],
                    grandTotal: row[14],
                    json: row[16] // 詳細資料
                });
            }
        });
        costList.reverse(); // 最新在最前
    }

    var result = {
        sales: sales,
        units: units,
        vendors: vendors, // [新增]
        workers: workers,
        buyers: buyers,
        nextProjectId: todayStr + nextSeq,
        projects: projectList,
        costs: costList
    };

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    try {
        var data = JSON.parse(e.postData.contents);

        // --- [修改] 通用設定儲存邏輯 ---
        if (data.action === "save_config") {
            var configSheet = ss.getSheetByName("系統設定");
            if (!configSheet) configSheet = ss.insertSheet("系統設定");

            var colIndex = 0;
            switch (data.configType) {
                case 'sales': colIndex = 1; break;  // A
                case 'unit': colIndex = 2; break;   // B
                case 'vendor': colIndex = 3; break; // C [新增]
                case 'worker': colIndex = 1; break; // 共用 sales
                case 'buyer': colIndex = 1; break;  // 共用 sales
            }

            if (colIndex > 0 && data.value) {
                // 找出該欄最後一列
                var lastRow = configSheet.getLastRow();
                var rangeLetter;
                if (colIndex === 1) {
                    rangeLetter = "A";
                } else if (colIndex === 2) {
                    rangeLetter = "B";
                } else if (colIndex === 3) {
                    rangeLetter = "C";
                } else {
                    // Should not happen with current configTypes, but as a fallback
                    // or if new types are added without updating this logic.
                    // For now, it will only be 1, 2, or 3.
                    rangeLetter = "A"; // Default to A if something unexpected happens
                }
                var values = configSheet.getRange(rangeLetter + "1:" + rangeLetter).getValues();
                // [修正] 二維陣列平面化後再過濾空值，確保計數正確
                var realLast = values.flat().filter(function (c) { return c !== ""; }).length + 1;
                configSheet.getRange(realLast, colIndex).setValue(data.value);
            }

            return ContentService.createTextOutput(JSON.stringify({ status: "success" }));
        }

        // --- 核心邏輯 ---
        var systemType = data.type;
        var sheetName = "";
        var headerRow = [];

        if (systemType === 'quotation') {
            sheetName = "報價紀錄";
            // [更新] 對應文件 2.1 報價紀錄 (A-S)
            headerRow = [
                "時間戳記", "單號", "版號", "日期", "客戶/廠商", "專案名稱",
                "專案ID(內控)", "業務人員", "業務手機", "LINE ID", "LINE 名稱",
                "未稅金額", "稅金", "總金額", "備註",
                "傳真", "狀態", "詳細資料(JSON)", "檔案連結"
            ];
        } else if (systemType === 'cost') {
            sheetName = "成本紀錄";
            // [更新] 對應文件 2.2 成本紀錄 (A-S)
            headerRow = [
                "時間戳記", "成本單號", "版號", "日期", "專案名稱(來源)", "廠商名稱", "廠商統編",
                "發票號碼", "業務人員", "業務手機", "工務人員", "採購人員",
                "未稅金額", "稅金", "總金額", "備註",
                "詳細資料(JSON)", "分析表連結", "發票連結"
            ];
        } else {
            sheetName = systemType === 'accounting' ? "會計紀錄" : "其他紀錄";
            // [更新] 對應文件 2.3 會計紀錄 (A-F...)
            headerRow = ["時間戳記", "單號", "日期", "對象名稱", "說明", "金額", "備註", "JSON", "連結"];
        }

        var sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
            sheet = ss.insertSheet(sheetName);
            sheet.appendRow(headerRow);
        }

        // --- 版本控制邏輯 (維持不變) ---
        var version = 0;
        var fileSuffix = "";
        var orderId = data.orderId;

        // [新增] 追加單據 ID 邏輯 (Q-YYYYMMDDxxxxA)
        if (systemType === 'quotation' && data.isAddition) {
            var rootId = "";
            var isRevisionOfAddition = false;

            if (data.originId) {
                // 檢查是否已經是追加單 (結尾是英文)
                var match = data.originId.match(/^(Q-\d{12})([A-Z])?(-[\d]+)?$/);
                if (match) {
                    rootId = match[1]; // Q-YYYYMMDDxxxx
                    if (match[2]) isRevisionOfAddition = true; // 已經有後綴 (A)
                }
            } else if (data.internalProjectId) {
                // 新單據但標記為追加 (罕見，但也防呆)
                rootId = "Q-" + data.internalProjectId.substring(0, 12);
            }

            if (rootId && !isRevisionOfAddition) {
                // 產生新的追加後綴 (從 Root 衍生)
                var nextChar = 'A';
                var rows = sheet.getDataRange().getValues();
                var maxChar = '';

                // 掃描所有內控 ID (Col G -> Index 6)
                // 內控 ID 格式: YYYYMMDDxxxx[A-Z]?
                var rootInternal = rootId.substring(2); // Remove Q-

                for (var i = 1; i < rows.length; i++) {
                    var iId = rows[i][6] ? rows[i][6].toString() : "";
                    if (iId.indexOf(rootInternal) === 0) {
                        var suffix = iId.substring(rootInternal.length);
                        // 只找第一層後綴 (A, B...), 忽略 -2, -3 ver
                        // 實際上內控ID欄位通常不帶版號 (版號在Col C)
                        // 若內控ID為 YYYYMMDDxxxxA
                        if (suffix.length === 1 && suffix >= 'A' && suffix <= 'Z') {
                            if (suffix > maxChar) maxChar = suffix;
                        }
                    }
                }

                if (maxChar) {
                    nextChar = String.fromCharCode(maxChar.charCodeAt(0) + 1);
                }

                // 設定新 ID
                internalId = rootInternal + nextChar;
                orderId = "Q-" + internalId;

                // 這是新的一條脈絡 (A, B...)，所以視為 V1，不算 Revision
                data.originId = null;
                fileSuffix = ""; // 第一版無後綴

                // 更新日期顯示 (追加 A)
                if (data.date.indexOf("(") === -1) {
                    data.date = data.date + " (" + nextChar + ")";
                }
            } else if (isRevisionOfAddition) {
                // 既有追加單的修訂 (A -> A-2)
                // 維持 originId，讓下方標準邏輯處理版號
                // 但確保日期有後綴?
                var currentSuffix = data.originId.match(/([A-Z])(-[\d]+)?$/)[1];
                if (data.date.indexOf("(") === -1) {
                    data.date = data.date + " (" + currentSuffix + ")";
                }
            }
        }

        if (data.originId) {
            var rows = sheet.getDataRange().getValues();
            var maxVer = 0;
            for (var i = 1; i < rows.length; i++) {
                var dbId = rows[i][1].toString();
                if (dbId.indexOf(data.originId) === 0) {
                    var v = parseInt(rows[i][2] || 0);
                    if (v > maxVer) maxVer = v;
                }
            }
            version = maxVer + 1;
            orderId = data.originId + "-" + version;
            fileSuffix = "-" + version;
        }

        // --- 檔案處理 ---
        var fileUrl = data.fileUrl || "";
        if (data.image) {
            var folderId = "";
            var folderName = "工程管理系統_未分類";

            // [更新] 依照系統類別指定資料夾 ID
            if (systemType === 'quotation') {
                folderId = "1jCLKQTw4LRrGeFUIOhWFsfFPhIMqSk3m";
                folderName = "報價單圖檔";
            } else if (systemType === 'cost') {
                folderId = "13d4nBjU3zCrBfWw2SOIviofddZ7GCGWS";
                folderName = "成本單圖檔";
            } else if (systemType === 'accounting') {
                folderId = "1ys0WG8oWuzt2-NHz_AbomAQgbFER-eUx";
                folderName = "會計單圖檔";
            }

            var folder;
            if (folderId) {
                try {
                    folder = DriveApp.getFolderById(folderId);
                } catch (e) {
                    // ID 錯誤或無權限時的回退機制
                    folder = getOrCreateFolder(folderName);
                }
            } else {
                folder = getOrCreateFolder(folderName);
            }

            var subFolderName;

            // 定義子資料夾名稱規則: 年份+日期+客戶+工程名稱
            // 由於 data.date 格式通常為 YYYY-MM-DD，這裡直接使用
            var dateStr = (data.date || "").replace(/-/g, ""); // Remove dashes for YYYYMMDD
            var safeCustomer = (data.customer || data.vendorName || "無客戶").replace(/[\/\\:*?"<>|]/g, "_");
            var safeProject = (data.project || "無專案").replace(/[\/\\:*?"<>|]/g, "_");

            if (systemType === 'cost') {
                // 成本系統: 年份+日期+廠商+工程名稱 (Cost usually tracks by Project, but user rule: Customer+Project. Cost has Project. Let's use Project as main context)
                // User rule: "報價單上之年份+日期+客戶+工程名稱" (Specific to Quotation?)
                // User also said: "各系統資料夾將相關圖檔置放統一資料夾內" -> "Unified folder within each system folder"
                // And "資料夾名稱設定規則 報價單上之..."
                // Let's apply YYYYMMDD_Customer_Project to all systems where possible.
                // For Cost, "Customer" might be the Project Owner (data.customer)
                subFolderName = dateStr + "_" + safeCustomer + "_" + safeProject;
            } else if (systemType === 'accounting') {
                // Accounting might not have "Project", but we mapped it earlier.
                subFolderName = dateStr + "_" + safeCustomer + "_" + safeProject;
            } else {
                // Quotation
                subFolderName = dateStr + "_" + safeCustomer + "_" + safeProject;
            }

            // 在系統主資料夾下，取得或建立該子資料夾
            folder = getOrCreateSubFolder(folder, subFolderName);

            var safeProjectName = (data.project || "專案").replace(/[\/\\:*?"<>|]/g, "_");
            // 檔名加上版號，若已成交加上「最終版」
            var statusSuffix = (data.status === 'closed' || data.status === '已成交') ? "_最終版" : "";

            // [修正] 會計系統可能沒有 project 名稱，改用對象名稱 + 摘要
            var namePart = safeProjectName;
            if (systemType === 'accounting') {
                namePart = (data.customer || "廠商") + "_" + (data.project || "摘要"); // data.project 在 accounting 中被 map 到 summary
            }

            var fileName = [data.date, namePart].join("_") + statusSuffix + fileSuffix + ".jpg";

            // 報價單檔名格式維持: 日期_客戶_專案...
            if (systemType === 'quotation') {
                fileName = [data.date, data.customer, safeProjectName].join("_") + statusSuffix + fileSuffix + ".jpg";
            }

            fileUrl = saveBase64ToDrive(folder, data.image, fileName);
            data.fileUrl = fileUrl;
        }

        var invoiceUrl = "";
        if (data.invoiceImage) {
            // 發票憑證也放入同一個專案子資料夾?
            // User said: "各系統資料夾將相關圖檔置放統一資料夾內"
            // Usually Invoice is for Cost. 
            // Previous logic: getOrCreateFolder("廠商發票憑證") -> Root level folder.
            // Now: "Unified folder". 
            // It makes strict sense to put invoices related to this Cost Record into the SAME subfolder as the Cost Sheet image.

            // Get root folder for invoices? Or put inside the same Project folder in Cost System?
            // "各系統資料夾...置放統一資料夾內"
            // If I am in Cost System -> Root is Cost Folder -> Subfolder is Project Folder.
            // So Invoice should go into Cost/ProjectFolder/

            var folderId = "13d4nBjU3zCrBfWw2SOIviofddZ7GCGWS"; // Cost Folder Default
            var rootFolder;
            try { rootFolder = DriveApp.getFolderById(folderId); } catch (e) { rootFolder = getOrCreateFolder("成本單圖檔"); }

            var dateStr = (data.date || "").replace(/-/g, "");
            var safeCustomer = (data.customer || "無客戶").replace(/[\/\\:*?"<>|]/g, "_"); // Cost has data.customer? Yes from payload.
            var safeProject = (data.project || "無專案").replace(/[\/\\:*?"<>|]/g, "_");
            var subFolderName = dateStr + "_" + safeCustomer + "_" + safeProject;

            var targetFolder = getOrCreateSubFolder(rootFolder, subFolderName);

            var safeVendor = (data.vendorName || "廠商").replace(/[\/\\:*?"<>|]/g, "_");
            var invFileName = [data.date, safeVendor, (data.invoiceNo || ""), orderId].join("_") + ".jpg";
            invoiceUrl = saveBase64ToDrive(targetFolder, data.invoiceImage, invFileName);
            data.invoiceUrl = invoiceUrl;
        }

        // --- [新增] 自動生成內控 ID (針對新報價單) ---
        var internalId = data.internalProjectId || "";
        if (systemType === 'quotation' && !data.originId && !internalId) {
            var todayStr = Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd");
            var nextSeq = 1;
            var rows = sheet.getDataRange().getValues();
            for (var i = 1; i < rows.length; i++) {
                var pid = rows[i][6] ? rows[i][6].toString() : "";
                if (pid.indexOf(todayStr) === 0 && pid.length === 12) {
                    var seq = parseInt(pid.substring(8));
                    if (seq >= nextSeq) nextSeq = seq + 1;
                }
            }
            internalId = todayStr + ("0000" + nextSeq).slice(-4);
        }

        // [新增] 如果是新報價單，自動生成正式單號 (Format: Q-YYYYMMDDxxxx)
        if (!orderId && internalId) {
            orderId = "Q-" + internalId;
        }

        // --- [修改] 寫入資料 (加入 salesMobile) ---
        var rowData = [];
        var timestamp = new Date();
        var lineId = (data.lineUser && data.lineUser.userId) ? data.lineUser.userId : "";
        var lineName = (data.lineUser && data.lineUser.displayName) ? data.lineUser.displayName : "";
        var mobile = data.salesMobile || "";


        // --- 根據 SystemType 處理寫入 ---
        if (systemType === 'accounting') {
            // 欄位定義: 
            // A: 時間, B: ID, C: Date, D: Customer(對象), E: Project(摘要), F: Amount, G: Note
            // H: JSON(Full Data), I: FileUrl
            // 這裡我們需要儲存額外的會計資訊 (Bank info, Type, etc.)
            // 將其全部放入 JSON (Column H) 中即可，前端已將這些打包進 data

            // 如果是 "Bonus" (獎金分配)，格式可能稍有不同，但仍可共用此結構
            // 對象名稱 = 人員名稱, 摘要 = 專案名稱 + "獎金分配"

            rowData = [
                timestamp, orderId, data.date,
                data.customer || data.vendorName || "", // 對象
                data.project || "一般帳務", // 摘要
                data.grandTotal,
                data.note,
                JSON.stringify(data),
                fileUrl
            ];
            sheet.appendRow(rowData);
        } else if (systemType === 'quotation') {


            rowData = [
                timestamp, orderId, version, data.date, data.customer, data.project,
                internalId, data.salesPerson || "", mobile, lineId, lineName,
                data.subtotalNoTax, data.taxAmount, data.grandTotal, data.note,
                data.fax || "", data.status || "", // [新增] Fax(Column P), Status(Column Q)
                JSON.stringify(data), fileUrl
            ];
            sheet.appendRow(rowData);
        } else if (systemType === 'cost') {
            var items = data.items || [{
                vendorName: data.vendorName,
                vendorTaxId: data.vendorTaxId,
                invoiceNo: data.invoiceNo,
                costSummary: data.costSummary,
                subtotalNoTax: data.subtotalNoTax,
                taxAmount: data.taxAmount,
                grandTotal: data.grandTotal,
                invoiceImage: data.invoiceImage // 支援單一舊格式
            }];

            items.forEach(function (item, idx) {
                var itemInvoiceUrl = "";
                var itemInvoiceUrl = "";
                if (item.invoiceImage) {
                    // Use the same subfolder logic for item images in Cost System
                    var folderId = "13d4nBjU3zCrBfWw2SOIviofddZ7GCGWS";
                    var rootFolder;
                    try { rootFolder = DriveApp.getFolderById(folderId); } catch (e) { rootFolder = getOrCreateFolder("成本單圖檔"); }

                    var dateStr = (data.date || "").replace(/-/g, "");
                    var safeCustomer = (data.customer || "無客戶").replace(/[\/\\:*?"<>|]/g, "_");
                    var safeProject = (data.project || "無專案").replace(/[\/\\:*?"<>|]/g, "_");
                    var subFolderName = dateStr + "_" + safeCustomer + "_" + safeProject;

                    var targetFolder = getOrCreateSubFolder(rootFolder, subFolderName);

                    var safeVendor = (item.vendorName || "廠商").replace(/[\/\\:*?"<>|]/g, "_");
                    var errorSuffix = item.isError ? "_錯誤" : "";
                    var invFileName = [data.date, safeVendor, (item.invoiceNo || ""), orderId, idx].join("_") + errorSuffix + ".jpg";
                    itemInvoiceUrl = saveBase64ToDrive(targetFolder, item.invoiceImage, invFileName);
                }

                var row = [
                    timestamp, orderId, version, data.date, data.project,
                    item.vendorName || "", item.vendorTaxId || "",
                    item.invoiceNo || "", data.salesPerson, mobile, // [更新] 加入業務手機 (Column J)
                    data.worker, data.buyer, // [更新] 變數名稱依照前端 payload (worker, buyer)
                    item.subtotalNoTax || 0,
                    item.taxAmount || 0,
                    item.grandTotal || 0,
                    (item.costSummary || "") + " " + (data.note || ""),
                    JSON.stringify(item), fileUrl, itemInvoiceUrl || invoiceUrl
                ];
                sheet.appendRow(row);
                // 更新回傳用的第一個 URL (相容性)
                if (idx === 0) invoiceUrl = itemInvoiceUrl;
            });
        } else {
            // [修正] 對應 headerRow，加入對象名稱 (data.customer or vendor)
            rowData = [timestamp, orderId, data.date, data.customer || data.vendorName || "", data.project || "一般紀錄", data.grandTotal, data.note, JSON.stringify(data), fileUrl];
            sheet.appendRow(rowData);
        }

        return ContentService.createTextOutput(JSON.stringify({
            status: "success",
            message: "儲存成功 (" + (data.items ? data.items.length : 1) + "筆)",
            id: orderId,
            internalId: internalId,
            fileUrl: fileUrl,
            invoiceUrl: invoiceUrl
        })).setMimeType(ContentService.MimeType.JSON);

    } catch (error) {
        return ContentService.createTextOutput(JSON.stringify({
            status: "error",
            message: error.toString()
        })).setMimeType(ContentService.MimeType.JSON);
    }
}

// 輔助函式
function getOrCreateFolder(name) {
    var folders = DriveApp.getFoldersByName(name);
    if (folders.hasNext()) return folders.next();
    return DriveApp.createFolder(name);
}

function getOrCreateSubFolder(parentFolder, name) {
    if (!parentFolder) return getOrCreateFolder(name); // Fallback
    var folders = parentFolder.getFoldersByName(name);
    if (folders.hasNext()) return folders.next();
    return parentFolder.createFolder(name);
}

function saveBase64ToDrive(folder, base64, name) {
    var base64Data = base64.split(',')[1] || base64;
    var decoded = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decoded, "image/jpeg", name);
    return folder.createFile(blob).getUrl();
}

function checkVendorHistory(taxId) {
    try {
        var start = new Date();
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        // 嘗試從「系統設定」的廠商名單找 (C欄)
        // 假設 C欄格式為 "名稱 (統編)" 或單純名稱，若無統編欄位則難以精確比對
        // 這裡改為搜尋「成本紀錄」的 G欄(統編) 與 F欄(廠商名稱)
        var sheet = ss.getSheetByName("成本紀錄");
        if (!sheet) return null;

        // 搜尋最近 500 筆即可，避免效能問題
        var lastRow = sheet.getLastRow();
        if (lastRow < 2) return null;

        var startRow = Math.max(2, lastRow - 500);
        var numRows = lastRow - startRow + 1;
        // G欄=Index 6, F欄=Index 5
        var data = sheet.getRange(startRow, 6, numRows, 2).getValues(); // F:名稱, G:統編 ? No. Cost header: ..., "廠商名稱", "廠商統編" (F, G)
        // Adjust for index. Cost header: 1:Time, 2:Id, 3:Ver, 4:Date, 5:Proj, 6:VendorName(F), 7:TaxId(G)
        // getRange(row, column) -> 6 is VendorName, 7 is TaxId

        var data = sheet.getRange(startRow, 6, numRows, 2).getValues();
        // data[i][0] = VendorName, data[i][1] = TaxId

        for (var i = data.length - 1; i >= 0; i--) {
            if (data[i][1] && data[i][1].toString() === taxId) {
                return data[i][0];
            }
        }
        return null;
    } catch (e) {
        return null;
    }
}