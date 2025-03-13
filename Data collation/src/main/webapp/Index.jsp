<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.sql.*"%>
<jsp:useBean id='objDBConfig' scope='application' class='hitstd.group.tool.database.DBConfig' />

<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>鄉鎮人口統計Excel解析器</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        body {
            font-family: 'Microsoft JhengHei', Arial, sans-serif;
            line-height: 1.6;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background-color: #fff;
            padding: 20px;
            border-radius: 5px;
            box-shadow: 0 2px 5px rgba(0, 0, 0, 0.1);
        }
        h1, h2 {
            color: #2c3e50;
        }
        .upload-section {
            margin-bottom: 20px;
            padding: 15px;
            background-color: #f8f9fa;
            border-radius: 4px;
        }
        input[type="file"] {
            margin-right: 10px;
        }
        button {
            padding: 8px 15px;
            background-color: #3498db;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            margin-right: 5px;
        }
        button:hover {
            background-color: #2980b9;
        }
        #fileInfo {
            margin: 10px 0;
            font-style: italic;
        }
        .progress-container {
            margin: 15px 0;
            background-color: #eee;
            border-radius: 10px;
            height: 20px;
            display: none;
        }
        .progress-bar {
            height: 100%;
            border-radius: 10px;
            background-color: #3498db;
            width: 0%;
            transition: width 0.3s;
        }
        .result-section {
            margin-top: 20px;
        }
        .tab-buttons {
            margin-bottom: 15px;
        }
        .tab-button {
            padding: 10px 15px;
            border: none;
            background-color: #f2f2f2;
            cursor: pointer;
            margin-right: 5px;
            border-radius: 4px 4px 0 0;
        }
        .tab-button.active {
            background-color: #3498db;
            color: white;
        }
        .tab-content {
            display: none;
        }
        .tab-content.active {
            display: block;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 15px;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
        }
        tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        #resultSummary {
            margin: 10px 0;
            font-weight: bold;
        }
        .export-btn {
            margin-top: 15px;
            background-color: #27ae60;
        }
        .export-btn:hover {
            background-color: #2ecc71;
        }
        .error-message {
            color: #e74c3c;
            background-color: #fadbd8;
            padding: 10px;
            border-radius: 4px;
            margin: 10px 0;
            display: none;
        }
        .success-message {
            color: #27ae60;
            background-color: #d4efdf;
            padding: 10px;
            border-radius: 4px;
            margin: 10px 0;
            display: none;
        }
        .info-panel {
            background-color: #eaf2f8;
            border-left: 4px solid #3498db;
            padding: 10px 15px;
            margin: 15px 0;
            font-size: 14px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>鄉鎮人口統計Excel解析器</h1>
        
        <div class="upload-section">
            <h2>1. 選擇Excel檔案</h2>
            <input type="file" id="excelFile" accept=".xls,.xlsx" />
            <button id="parseBtn">解析檔案</button>
            <div id="fileInfo"></div>
            <div class="progress-container" id="progressContainer">
                <div class="progress-bar" id="progressBar"></div>
            </div>
            <div id="errorMessage" class="error-message"></div>
            <div id="successMessage" class="success-message"></div>
        </div>
        
        <div class="result-section" id="resultSection" style="display: none;">
            <h2>2. 解析結果</h2>
            <div id="resultSummary"></div>
            
            <div class="tab-buttons">
                <button id="originalTabBtn" class="tab-button active">原始預覽</button>
                <button id="unmergedTabBtn" class="tab-button">解除合併後預覽</button>
                <button id="processedTabBtn" class="tab-button">處理後資料</button>
            </div>
            
            <div id="originalTab" class="tab-content active">
                <div class="info-panel">
                    顯示原始Excel檔案的預覽，包含合併儲存格。
                </div>
                <div class="table-container">
                    <table id="originalTable"></table>
                </div>
            </div>
            
            <div id="unmergedTab" class="tab-content">
                <div class="info-panel">
                    顯示解除所有合併儲存格後的Excel預覽。
                </div>
                <div class="table-container">
                    <table id="unmergedTable"></table>
                </div>
            </div>
            
            <div id="processedTab" class="tab-content">
                <div class="info-panel">
                    顯示處理後的結構化資料。
                </div>
                <div class="table-container">
                    <table id="processedTable">
                        <thead>
                            <tr>
                                <th>年分</th>
                                <th>月份</th>
                                <th>縣市別</th>
                                <th>鄉鎮別</th>
                                <th>性別</th>
                                <th>戶數</th>
                                <th>人口數</th>
                            </tr>
                        </thead>
                        <tbody id="processedBody"></tbody>
                    </table>
                </div>
            </div>
            
            <div style="margin-top: 15px;">
                <button id="exportXLSX" class="export-btn">匯出 Excel</button>
            </div>
        </div>
    </div>

    <script>
        document.addEventListener('DOMContentLoaded', function() {
            // DOM 元素
            var excelFileInput = document.getElementById('excelFile');
            var parseBtn = document.getElementById('parseBtn');
            var fileInfo = document.getElementById('fileInfo');
            var progressContainer = document.getElementById('progressContainer');
            var progressBar = document.getElementById('progressBar');
            var errorMessage = document.getElementById('errorMessage');
            var successMessage = document.getElementById('successMessage');
            var resultSection = document.getElementById('resultSection');
            var resultSummary = document.getElementById('resultSummary');
            var originalTable = document.getElementById('originalTable');
            var unmergedTable = document.getElementById('unmergedTable');
            var processedBody = document.getElementById('processedBody');
            var exportXLSXBtn = document.getElementById('exportXLSX');
            var originalTabBtn = document.getElementById('originalTabBtn');
            var unmergedTabBtn = document.getElementById('unmergedTabBtn');
            var processedTabBtn = document.getElementById('processedTabBtn');
            var originalTab = document.getElementById('originalTab');
            var unmergedTab = document.getElementById('unmergedTab');
            var processedTab = document.getElementById('processedTab');
            
            // 存儲解析後的資料
            var originalWorkbook = null;
            var unmergedWorkbook = null;
            var parsedData = [];
            var fileName = '';
            var fileYear = '';
            var fileMonth = '';
            
            // 標籤切換
            originalTabBtn.addEventListener('click', function() {
                originalTabBtn.classList.add('active');
                unmergedTabBtn.classList.remove('active');
                processedTabBtn.classList.remove('active');
                originalTab.classList.add('active');
                unmergedTab.classList.remove('active');
                processedTab.classList.remove('active');
            });
            
            unmergedTabBtn.addEventListener('click', function() {
                unmergedTabBtn.classList.add('active');
                originalTabBtn.classList.remove('active');
                processedTabBtn.classList.remove('active');
                unmergedTab.classList.add('active');
                originalTab.classList.remove('active');
                processedTab.classList.remove('active');
            });
            
            processedTabBtn.addEventListener('click', function() {
                processedTabBtn.classList.add('active');
                originalTabBtn.classList.remove('active');
                unmergedTabBtn.classList.remove('active');
                processedTab.classList.add('active');
                originalTab.classList.remove('active');
                unmergedTab.classList.remove('active');
            });
            
            // 解析按鈕點擊事件
            parseBtn.addEventListener('click', function() {
                var file = excelFileInput.files[0];
                if (!file) {
                    showError('請先選擇Excel檔案');
                    return;
                }
                
                fileName = file.name;
                fileInfo.textContent = '檔案: ' + fileName + ' (' + (file.size / 1024).toFixed(2) + ' KB)';
                
                // 嘗試從檔名中提取年月
                var regexResult = /(\d+)年(\d+)月/.exec(fileName);
                if (regexResult) {
                    fileYear = regexResult[1]; // 民國年
                    fileMonth = regexResult[2];
                    fileInfo.textContent += ' - 民國' + fileYear + '年' + fileMonth + '月資料';
                }
                
                hideError();
                hideSuccess();
                showProgressBar();
                
                // 讀取檔案
                var reader = new FileReader();
                
                reader.onprogress = function(e) {
                    if (e.lengthComputable) {
                        var percentLoaded = Math.round((e.loaded / e.total) * 100);
                        updateProgressBar(percentLoaded / 3); // 讀取佔整體進度的三分之一
                    }
                };
                
                reader.onload = function(e) {
                    try {
                        // 開始解析
                        updateProgressBar(33);
                        var data = new Uint8Array(e.target.result);
                        parseExcelFile(data);
                        
                        // 完成
                        updateProgressBar(100);
                        setTimeout(function() {
                            hideProgressBar();
                        }, 500);
                    } catch (error) {
                        showError('解析檔案時發生錯誤: ' + error.message);
                        console.error(error);
                        hideProgressBar();
                    }
                };
                
                reader.onerror = function() {
                    showError('讀取檔案時發生錯誤');
                    hideProgressBar();
                };
                
                reader.readAsArrayBuffer(file);
            });
            
            // 匯出Excel按鈕事件
            exportXLSXBtn.addEventListener('click', function() {
                if (parsedData.length === 0) {
                    showError('沒有資料可匯出');
                    return;
                }
                
                try {
                    // 創建新的工作簿
                    var wb = XLSX.utils.book_new();
                    
                    // 創建工作表數據
                    var wsData = [['年分', '月份', '縣市別', '鄉鎮別', '性別', '戶數', '人口數']];
                    
                    // 添加資料行
                    parsedData.forEach(function(row) {
                        wsData.push([
                            row.year,
                            row.month,
                            row.city,
                            row.district,
                            row.gender,
                            row.households,
                            row.population
                        ]);
                    });
                    
                    // 將數據轉換為工作表
                    var ws = XLSX.utils.aoa_to_sheet(wsData);
                    
                    // 添加工作表到工作簿
                    XLSX.utils.book_append_sheet(wb, ws, '鄉鎮人口統計');
                    
                    // 導出為XLSX文件
                    XLSX.writeFile(wb, '人口統計_' + fileYear + '年' + fileMonth + '月.xlsx');
                    showSuccess('已成功匯出Excel檔案');
                } catch (error) {
                    showError('匯出Excel時發生錯誤: ' + error.message);
                    console.error(error);
                }
            });
            
            // 解析Excel函數
            function parseExcelFile(data) {
                // 使用SheetJS讀取Excel
                originalWorkbook = XLSX.read(data, { type: 'array' });
                
                // 獲取第一個工作表
                var firstSheetName = originalWorkbook.SheetNames[0];
                var worksheet = originalWorkbook.Sheets[firstSheetName];
                
                // 顯示原始Excel預覽
                displayOriginalPreview(worksheet);
                
                // 解除合併儲存格
                updateProgressBar(66);
                unmergedWorkbook = unmergeCells(originalWorkbook);
                var unmergedWorksheet = unmergedWorkbook.Sheets[firstSheetName];
                
                // 顯示解除合併後的Excel預覽
                displayUnmergedPreview(unmergedWorksheet);
                
                // 從解除合併儲存格後的工作表提取資料
                var jsonData = XLSX.utils.sheet_to_json(unmergedWorksheet, { header: 1 });
                
                // 提取資料 (根據明確指示)
                var extractedData = extractDataFromSpecificFormat(jsonData);
                
                // 顯示處理後資料
                displayProcessedData(extractedData);
            }
            
            // 解除合併儲存格
            function unmergeCells(workbook) {
                // 複製原始工作簿，避免修改原始資料
                var newWorkbook = XLSX.utils.book_new();
                
                // 處理每個工作表
                workbook.SheetNames.forEach(function(sheetName) {
                    var originalSheet = workbook.Sheets[sheetName];
                    
                    // 複製工作表
                    var newSheet = {};
                    Object.keys(originalSheet).forEach(function(key) {
                        if (key !== '!merges') { // 不複製合併資訊
                            newSheet[key] = originalSheet[key];
                        }
                    });
                    
                    // 如果有合併儲存格，執行解除合併
                    if (originalSheet['!merges']) {
                        originalSheet['!merges'].forEach(function(merge) {
                            // 獲取左上角儲存格的值
                            var topLeftCellAddress = XLSX.utils.encode_cell({r: merge.s.r, c: merge.s.c});
                            var topLeftCellValue = originalSheet[topLeftCellAddress];
                            
                            // 複製左上角的值到合併區域的所有儲存格
                            if (topLeftCellValue) {
                                for (var r = merge.s.r; r <= merge.e.r; r++) {
                                    for (var c = merge.s.c; c <= merge.e.c; c++) {
                                        var cellAddress = XLSX.utils.encode_cell({r: r, c: c});
                                        // 複製值和樣式，但使用新的儲存格地址
                                        newSheet[cellAddress] = {
                                            t: topLeftCellValue.t, // 類型
                                            v: topLeftCellValue.v, // 值
                                            // 可以複製其他屬性，如樣式等
                                        };
                                    }
                                }
                            }
                        });
                    }
                    
                    // 添加工作表到新工作簿
                    XLSX.utils.book_append_sheet(newWorkbook, newSheet, sheetName);
                });
                
                return newWorkbook;
            }
            
            // 顯示原始Excel預覽
            function displayOriginalPreview(worksheet) {
                // 獲取工作表範圍
                var range = XLSX.utils.decode_range(worksheet['!ref']);
                
                // 清空原始表格
                originalTable.innerHTML = '';
                
                // 創建表頭行
                var headerRow = document.createElement('tr');
                for (var c = range.s.c; c <= range.e.c; c++) {
                    var th = document.createElement('th');
                    th.textContent = XLSX.utils.encode_col(c);
                    headerRow.appendChild(th);
                }
                
                // 創建表頭元素
                var thead = document.createElement('thead');
                thead.appendChild(headerRow);
                originalTable.appendChild(thead);
                
                // 創建表身
                var tbody = document.createElement('tbody');
                
                // 建立一個二維陣列來追蹤已經處理過的合併儲存格
                var processedCells = Array(range.e.r + 1).fill().map(() => Array(range.e.c + 1).fill(false));
                
                // 顯示前20行數據
                var maxRows = Math.min(range.e.r + 1, 20);
                for (var r = range.s.r; r < maxRows; r++) {
                    var tr = document.createElement('tr');
                    
                    for (var c = range.s.c; c <= range.e.c; c++) {
                        // 如果這個儲存格已經被處理過（作為合併儲存格的一部分），則跳過
                        if (processedCells[r][c]) {
                            continue;
                        }
                        
                        var td = document.createElement('td');
                        var cellAddress = XLSX.utils.encode_cell({r: r, c: c});
                        var cell = worksheet[cellAddress];
                        
                        // 檢查是否為合併儲存格的一部分
                        var mergeInfo = null;
                        if (worksheet['!merges']) {
                            for (var i = 0; i < worksheet['!merges'].length; i++) {
                                var merge = worksheet['!merges'][i];
                                if (r >= merge.s.r && r <= merge.e.r && 
                                    c >= merge.s.c && c <= merge.e.c) {
                                    mergeInfo = merge;
                                    break;
                                }
                            }
                        }
                        
                        // 如果是合併儲存格的一部分
                        if (mergeInfo) {
                            // 只處理合併儲存格的左上角
                            if (r === mergeInfo.s.r && c === mergeInfo.s.c) {
                                // 設置 colspan 和 rowspan
                                td.colSpan = mergeInfo.e.c - mergeInfo.s.c + 1;
                                td.rowSpan = mergeInfo.e.r - mergeInfo.s.r + 1;
                                
                                // 顯示左上角儲存格的值
                                if (cell) {
                                    td.textContent = cell.v;
                                }
                                
                                // 標記所有被這個合併儲存格覆蓋的儲存格為已處理
                                for (var mr = mergeInfo.s.r; mr <= mergeInfo.e.r; mr++) {
                                    for (var mc = mergeInfo.s.c; mc <= mergeInfo.e.c; mc++) {
                                        processedCells[mr][mc] = true;
                                    }
                                }
                                
                                tr.appendChild(td);
                            }
                            // 其他部分的合併儲存格已經被標記為已處理，會被跳過
                        } else {
                            // 不是合併儲存格的一部分，直接顯示
                            if (cell) {
                                td.textContent = cell.v;
                            }
                            tr.appendChild(td);
                        }
                    }
                    
                    // 只有當行中有內容時才添加到表格
                    if (tr.children.length > 0) {
                        tbody.appendChild(tr);
                    }
                }
                
                originalTable.appendChild(tbody);
                
                // 如果有更多行，添加提示信息
                if (range.e.r + 1 > 20) {
                    var infoRow = document.createElement('tr');
                    var infoCell = document.createElement('td');
                    infoCell.colSpan = range.e.c + 1;
                    infoCell.style.textAlign = 'center';
                    infoCell.style.fontStyle = 'italic';
                    infoCell.textContent = '... 僅顯示前 20 行 (共 ' + (range.e.r + 1) + ' 行)';
                    infoRow.appendChild(infoCell);
                    tbody.appendChild(infoRow);
                }
            }
            
            // 顯示解除合併後的Excel預覽
            function displayUnmergedPreview(worksheet) {
                // 獲取工作表範圍
                var range = XLSX.utils.decode_range(worksheet['!ref']);
                
                // 清空解除合併表格
                unmergedTable.innerHTML = '';
                
                // 創建表頭行
                var headerRow = document.createElement('tr');
                for (var c = range.s.c; c <= range.e.c; c++) {
                    var th = document.createElement('th');
                    th.textContent = XLSX.utils.encode_col(c);
                    headerRow.appendChild(th);
                }
                
                // 創建表頭元素
                var thead = document.createElement('thead');
                thead.appendChild(headerRow);
                unmergedTable.appendChild(thead);
                
                // 創建表身
                var tbody = document.createElement('tbody');
                
                // 顯示前20行數據
                var maxRows = Math.min(range.e.r + 1, 20);
                for (var r = range.s.r; r < maxRows; r++) {
                    var tr = document.createElement('tr');
                    
                    for (var c = range.s.c; c <= range.e.c; c++) {
                        var td = document.createElement('td');
                        var cellAddress = XLSX.utils.encode_cell({r: r, c: c});
                        var cell = worksheet[cellAddress];
                        
                        if (cell) {
                            td.textContent = cell.v;
                        }
                        
                        tr.appendChild(td);
                    }
                    
                    tbody.appendChild(tr);
                }
                
                unmergedTable.appendChild(tbody);
                
                // 如果有更多行，添加提示信息
                if (range.e.r + 1 > 20) {
                    var infoRow = document.createElement('tr');
                    var infoCell = document.createElement('td');
                    infoCell.colSpan = range.e.c + 1;
                    infoCell.style.textAlign = 'center';
                    infoCell.style.fontStyle = 'italic';
                    infoCell.textContent = '... 僅顯示前 20 行 (共 ' + (range.e.r + 1) + ' 行)';
                    infoRow.appendChild(infoCell);
                    tbody.appendChild(infoRow);
                }
            }
            
            // 根據特定格式提取資料
            function extractDataFromSpecificFormat(jsonData) {
                var result = [];
                
                try {
                    // 尋找標題行中包含「中華民國」的行以確定資料開始的位置
                    var titleRowIndex = -1;
                    for (var i = 0; i < Math.min(5, jsonData.length); i++) {
                        var row = jsonData[i];
                        if (row && row[0] && String(row[0]).includes('中華民國')) {
                            titleRowIndex = i;
                            
                            // 嘗試從標題中提取年月
                            if (!fileYear || !fileMonth) {
                                var titleText = String(row[0]);
                                var yearMonthMatch = titleText.match(/(\d+)年(\d+)月/);
                                if (yearMonthMatch) {
                                    fileYear = yearMonthMatch[1];
                                    fileMonth = yearMonthMatch[2];
                                }
                            }
                            break;
                        }
                    }
                    
                    // 如果沒找到標題行，就假設資料從第一行開始
                    if (titleRowIndex === -1) {
                        titleRowIndex = 0;
                    }
                    
                    // 定位到縣市資料開始的行
                    // 尋找縣市資料（通常含「臺北市」的行）
                    var cityRowIndex = -1;
                    var cityName = '';
                    
                    for (var i = 0; i < jsonData.length; i++) {
                        var row = jsonData[i];
                        if (row && row[0] && 
                            (String(row[0]).includes('臺北市') || 
                             String(row[0]).includes('台北市') || 
                             String(row[0]).includes('新北市') || 
                             String(row[0]).includes('臺中市') || 
                             String(row[0]).includes('高雄市'))) {
                            cityRowIndex = i;
                            cityName = String(row[0]);
                            break;
                        }
                    }
                    
                    if (cityRowIndex === -1) {
                        throw new Error('無法找到縣市資料行');
                    }
                    
                    console.log('縣市行:', cityRowIndex);
                    console.log('縣市名稱:', cityName);
                    console.log('該行數據:', jsonData[cityRowIndex]);
                    
                    // 處理縣市總計行
                    var cityRow = jsonData[cityRowIndex];
                    if (cityRow[1] !== undefined && cityRow[3] !== undefined && cityRow[4] !== undefined) {
                        // 男性資料
                        result.push({
                            year: fileYear,
                            month: fileMonth,
                            city: cityName,
                            district: '總計',
                            gender: '男',
                            households: cityRow[1], // 戶數
                            population: cityRow[3]  // 男性人口數
                        });
                        
                        // 女性資料
                        result.push({
                            year: fileYear,
                            month: fileMonth,
                            city: cityName,
                            district: '總計',
                            gender: '女',
                            households: cityRow[1], // 戶數
                            population: cityRow[4]  // 女性人口數
                        });
                    }
                    
                    // 從縣市行之後開始處理鄉鎮資料
                    for (var i = cityRowIndex + 1; i < jsonData.length; i++) {
                        var row = jsonData[i];
                        
                        // 確保行有效且有資料
                        if (!row || !row[0]) continue;
                        
                        var districtName = String(row[0]);
                        
                        // 跳過空行或非鄉鎮名的行 (排除空白和不包含區/鄉/鎮的名稱)
                        if (!districtName || districtName.trim() === '' || 
                            !(districtName.includes('區') || districtName.includes('鄉') || 
                              districtName.includes('鎮') || districtName.includes('市'))) continue;
                              
                        // 檢查是否已經到了下一個縣市（如果有）
                        if (districtName.includes('縣') || 
                            (districtName.includes('市') && districtName.length <= 3)) {
                            break;
                        }
                        
                        console.log('處理鄉鎮:', districtName, '行數據:', row);
                        
                        // 確保有人口數資料
                        if (row[1] === undefined || row[3] === undefined || row[4] === undefined) continue;
                        
                        // 男性資料
                        result.push({
                            year: fileYear,
                            month: fileMonth,
                            city: cityName,
                            district: districtName,
                            gender: '男',
                            households: row[1], // 戶數
                            population: row[3]  // 男性人口數
                        });
                        
                        // 女性資料
                        result.push({
                            year: fileYear,
                            month: fileMonth,
                            city: cityName,
                            district: districtName,
                            gender: '女',
                            households: row[1], // 戶數
                            population: row[4]  // 女性人口數
                        });
                    }
                    
                    if (result.length === 0) {
                        throw new Error('未能提取到任何有效資料');
                    }
                    
                    showSuccess('成功解析資料，共 ' + result.length + ' 筆記錄');
                    } catch (error) {
                    showError('資料解析錯誤: ' + error.message);
                    console.error('提取資料錯誤:', error);
                }
                
                return result;
            }
            
            // 顯示處理後資料
            function displayProcessedData(data) {
                parsedData = data;
                
                // 清空表格
                processedBody.innerHTML = '';
                
                // 更新摘要
                resultSummary.textContent = '共解析出 ' + data.length + ' 筆記錄';
                
                // 只顯示前 100 筆資料，避免瀏覽器變慢
                var displayData = data.slice(0, 100);
                
                // 添加資料到表格
                for (var i = 0; i < displayData.length; i++) {
                    var item = displayData[i];
                    var row = document.createElement('tr');
                    
                    var yearCell = document.createElement('td');
                    yearCell.textContent = item.year;
                    
                    var monthCell = document.createElement('td');
                    monthCell.textContent = item.month;
                    
                    var cityCell = document.createElement('td');
                    cityCell.textContent = item.city;
                    
                    var districtCell = document.createElement('td');
                    districtCell.textContent = item.district;
                    
                    var genderCell = document.createElement('td');
                    genderCell.textContent = item.gender;
                    
                    var householdsCell = document.createElement('td');
                    householdsCell.textContent = item.households;
                    
                    var populationCell = document.createElement('td');
                    populationCell.textContent = item.population;
                    
                    row.appendChild(yearCell);
                    row.appendChild(monthCell);
                    row.appendChild(cityCell);
                    row.appendChild(districtCell);
                    row.appendChild(genderCell);
                    row.appendChild(householdsCell);
                    row.appendChild(populationCell);
                    
                    processedBody.appendChild(row);
                }
                
                // 顯示結果區段
                resultSection.style.display = 'block';
                
                // 如果有 100 筆以上記錄，添加注意文字
                if (data.length > 100) {
                    resultSummary.textContent += ' (只顯示前 100 筆)';
                }
            }
            
            // 顯示錯誤訊息
            function showError(message) {
                errorMessage.textContent = message;
                errorMessage.style.display = 'block';
            }
            
            // 隱藏錯誤訊息
            function hideError() {
                errorMessage.style.display = 'none';
            }
            
            // 顯示成功訊息
            function showSuccess(message) {
                successMessage.textContent = message;
                successMessage.style.display = 'block';
            }
            
            // 隱藏成功訊息
            function hideSuccess() {
                successMessage.style.display = 'none';
            }
            
            // 顯示進度條
            function showProgressBar() {
                progressContainer.style.display = 'block';
                updateProgressBar(0);
            }
            
            // 隱藏進度條
            function hideProgressBar() {
                progressContainer.style.display = 'none';
            }
            
            // 更新進度條
            function updateProgressBar(percent) {
                progressBar.style.width = percent + '%';
            }
        });
    </script>
</body>
</html>