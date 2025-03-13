<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.sql.*"%>
<jsp:useBean id='objDBConfig' scope='application' class='hitstd.group.tool.database.DBConfig' />

<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>檔案預覽與合併儲存格解析工具</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.3.2/papaparse.min.js"></script>
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
        <h1>檔案預覽與合併儲存格解析工具</h1>
        
        <div class="upload-section">
            <h2>1. 選擇檔案</h2>
            <input type="file" id="excelFile" accept=".xls,.xlsx,.csv" />
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
                    顯示處理後的結構化資料，自動識別標題列及偵測年份與縣市。
                </div>
                <div id="detectedInfo" class="info-panel" style="background-color: #d4efdf; border-left: 4px solid #27ae60; margin-bottom: 15px; display: none;">
                    <h3 style="margin-top: 0;">偵測到的基本資訊</h3>
                    <div id="detectedYear">年份: <span id="yearValue">-</span></div>
                    <div id="detectedCity">縣市: <span id="cityValue">-</span></div>
                </div>
                <div class="table-container">
                    <table id="processedTable"></table>
                </div>
                <div style="margin-top: 15px;">
                    <button id="generateStructuredData" class="export-btn" style="background-color: #8e44ad;">生成整理後資料</button>
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
            var processedTable = document.getElementById('processedTable');
            var exportXLSXBtn = document.getElementById('exportXLSX');
            var originalTabBtn = document.getElementById('originalTabBtn');
            var unmergedTabBtn = document.getElementById('unmergedTabBtn');
            var processedTabBtn = document.getElementById('processedTabBtn');
            var originalTab = document.getElementById('originalTab');
            var unmergedTab = document.getElementById('unmergedTab');
            var processedTab = document.getElementById('processedTab');
            var detectedInfo = document.getElementById('detectedInfo');
            var yearValue = document.getElementById('yearValue');
            var cityValue = document.getElementById('cityValue');
            var generateStructuredDataBtn = document.getElementById('generateStructuredData');
            
            // 存儲解析後的資料
            var originalWorkbook = null;
            var unmergedWorkbook = null;
            var processedData = [];
            var structuredData = [];
            var fileName = '';
            var detectedYear = null;
            var detectedCities = [];
            
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
                    showError('請先選擇檔案');
                    return;
                }
                
                fileName = file.name;
                fileInfo.textContent = '檔案: ' + fileName + ' (' + (file.size / 1024).toFixed(2) + ' KB)';
                
                hideError();
                hideSuccess();
                showProgressBar();
                
                // 檢查檔案類型
                var fileType = file.name.split('.').pop().toLowerCase();
                
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
                        
                        if (fileType === 'csv') {
                            // 解析 CSV 檔案
                            var csvContent = e.target.result;
                            parseCSVFile(csvContent);
                        } else {
                            // 解析 Excel 檔案
                            var data = new Uint8Array(e.target.result);
                            parseExcelFile(data);
                        }
                        
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
                
                if (fileType === 'csv') {
                    reader.readAsText(file); // 使用 readAsText 讀取 CSV 檔案
                } else {
                    reader.readAsArrayBuffer(file); // 使用 readAsArrayBuffer 讀取 Excel 檔案
                }
            });
            
            // 匯出Excel按鈕事件
            exportXLSXBtn.addEventListener('click', function() {
                if (!unmergedWorkbook) {
                    showError('沒有資料可匯出');
                    return;
                }
                
                try {
                        // 如果有整理後的資料，則匯出整理後的資料
                    if (processedData && processedData.length > 0) {
                        // 獲取可用的欄位
                        var availableColumns = identifyAvailableColumns(structuredData);
                        
                        // 創建新的工作簿
                        var wb = XLSX.utils.book_new();
                        
                        // 創建工作表數據
                        var wsData = [];
                        
                        // 添加標題行
                        var headerRow = [];
                        availableColumns.forEach(function(column) {
                            headerRow.push(column.label);
                        });
                        wsData.push(headerRow);
                        
                        // 添加資料行
                        processedData.forEach(function(rowData) {
                            var dataRow = [];
                            availableColumns.forEach(function(column) {
                                dataRow.push(rowData[column.field] !== undefined ? rowData[column.field] : '');
                            });
                            wsData.push(dataRow);
                        });
                        
                        // 將數據轉換為工作表
                        var ws = XLSX.utils.aoa_to_sheet(wsData);
                        
                        // 添加工作表到工作簿
                        XLSX.utils.book_append_sheet(wb, ws, '整理後資料');
                        
                        // 導出為XLSX文件
                        XLSX.writeFile(wb, '整理後資料_' + fileName);
                        showSuccess('已成功匯出整理後資料Excel檔案');
                    } else {
                        // 匯出解除合併後的工作簿
                        XLSX.writeFile(unmergedWorkbook, '解除合併_' + fileName);
                        showSuccess('已成功匯出Excel檔案');
                    }
                } catch (error) {
                    showError('匯出Excel時發生錯誤: ' + error.message);
                    console.error(error);
                }
            });
            
            // 生成整理後資料按鈕事件
            generateStructuredDataBtn.addEventListener('click', function() {
                if (!unmergedWorkbook) {
                    showError('請先上傳並解析檔案');
                    return;
                }
                
                try {
                    // 根據偵測到的年份和縣市生成整理後資料
                    generateStructuredData();
                } catch (error) {
                    showError('生成整理後資料時發生錯誤: ' + error.message);
                    console.error(error);
                }
            });
            
            // 解析Excel函數
            function parseExcelFile(data) {
                // 使用SheetJS讀取Excel
                originalWorkbook = XLSX.read(data, { 
                    type: 'array',
                    cellStyles: true,
                    cellFormulas: true,
                    cellDates: true,
                    cellNF: true,
                    sheetStubs: true
                });
                
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
                
                // 識別標題列並創建處理後的預覽
                displayProcessedPreview(unmergedWorksheet);
                
                // 顯示成功訊息
                showSuccess('檔案解析成功！');
                
                // 顯示結果區段
                resultSection.style.display = 'block';
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
                    
                    // 保存合併區域信息用於報告
                    var mergedCount = 0;
                    if (originalSheet['!merges']) {
                        mergedCount = originalSheet['!merges'].length;
                        
                        // 如果有合併儲存格，執行解除合併
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
                                        };
                                    }
                                }
                            }
                        });
                    }
                    
                    // 添加工作表到新工作簿
                    XLSX.utils.book_append_sheet(newWorkbook, newSheet, sheetName);
                    
                    // 更新結果摘要
                    resultSummary.textContent = '已解除 ' + mergedCount + ' 個合併儲存格';
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
                
                // 顯示所有行數據
                for (var r = range.s.r; r <= range.e.r; r++) {
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
                                
                                // 高亮合併儲存格
                                td.style.backgroundColor = '#e8f4f8';
                                td.style.border = '2px solid #3498db';
                                
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
                
                // 顯示所有行數據
                for (var r = range.s.r; r <= range.e.r; r++) {
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
            }
            
            // 顯示處理後的預覽（識別標題列）
            function displayProcessedPreview(worksheet) {
                // 獲取工作表範圍
                var range = XLSX.utils.decode_range(worksheet['!ref']);
                
                // 清空處理後表格
                processedTable.innerHTML = '';
                
                // 將工作表轉換為JSON
                var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                // 識別可能的標題列（非空且有多個欄位的行）
                var headerRowIndex = -1;
                for (var i = 0; i < Math.min(10, jsonData.length); i++) {
                    var row = jsonData[i];
                    // 檢查行是否有足夠的欄位
                    if (row && row.filter(Boolean).length > 2) {
                        headerRowIndex = i;
                        break;
                    }
                }
                
                // 如果沒找到標題行，假設第一行是標題
                if (headerRowIndex === -1 && jsonData.length > 0) {
                    headerRowIndex = 0;
                }
                
                // 創建表頭
                var thead = document.createElement('thead');
                var headerRow = document.createElement('tr');
                
                if (headerRowIndex !== -1 && jsonData[headerRowIndex]) {
                    // 使用標題行作為表頭
                    jsonData[headerRowIndex].forEach(function(cellValue) {
                        var th = document.createElement('th');
                        th.textContent = cellValue || '';
                        headerRow.appendChild(th);
                    });
                } else {
                    // 如果沒有標題行，使用默認列標籤
                    for (var c = range.s.c; c <= range.e.c; c++) {
                        var th = document.createElement('th');
                        th.textContent = XLSX.utils.encode_col(c);
                        headerRow.appendChild(th);
                    }
                }
                
                thead.appendChild(headerRow);
                processedTable.appendChild(thead);
                
                // 創建表身
                var tbody = document.createElement('tbody');
                
                // 添加數據行，跳過標題行
                for (var i = 0; i < jsonData.length; i++) {
                    if (i === headerRowIndex) continue; // 跳過標題行
                    
                    var dataRow = jsonData[i];
                    if (!dataRow || dataRow.length === 0) continue; // 跳過空行
                    
                    var tr = document.createElement('tr');
                    
                    // 確保所有行具有相同的欄位數
                    var headerCols = headerRowIndex !== -1 ? jsonData[headerRowIndex].length : 0;
                    var maxCols = Math.max(dataRow.length, headerCols);
                    
                    for (var j = 0; j < maxCols; j++) {
                        var td = document.createElement('td');
                        if (j < dataRow.length) {
                            td.textContent = dataRow[j] !== undefined ? dataRow[j] : '';
                        }
                        tr.appendChild(td);
                    }
                    
                    tbody.appendChild(tr);
                }
                
                processedTable.appendChild(tbody);
                
                // 偵測年份和縣市
                detectYearAndCities(jsonData);
            }
            
            // 偵測年份和縣市
            function detectYearAndCities(jsonData) {
                // 台灣縣市列表
            var taiwanCities = [
                '臺北市', '台北市', '新北市', '桃園市', '臺中市', '台中市', 
                '臺南市', '台南市', '高雄市', '基隆市', '新竹市', '嘉義市',
                '新竹縣', '苗栗縣', '彰化縣', '南投縣', '雲林縣', '嘉義縣',
                '屏東縣', '宜蘭縣', '花蓮縣', '臺東縣', '台東縣', '澎湖縣',
                '金門縣', '連江縣'
            ];
            
            // 定義常見欄位對照表 - 用於自動分類
            var fieldMapping = {
                // 年份相關
                year: ['年份', '年度', '民國', '西元', '年', 'year', 'yr'],
                
                // 縣市相關
                city: ['縣市別', '縣市', '城市', '行政區', '縣', '市', 'city', 'county', 'prefecture'],
                
                // 區域/鄉鎮相關
                district: ['鄉鎮別', '鄉鎮市區', '區域', '鄉', '鎮', '區', '行政區', 'district', 'town', 'village'],
                
                // 人口相關
                population: ['人口數', '人口', '總人口', '總數', '人數', 'population', 'pop', 'total population'],
                
                // 戶數相關
                households: ['戶數', '家戶', '戶', '總戶數', 'households', 'household', 'family'],
                
                // 出生相關
                birth: ['出生數', '出生人數', '出生', '新生兒', '出生率', 'birth', 'births', 'birth rate'],
                
                // 死亡相關
                death: ['死亡數', '死亡人數', '死亡', '死亡率', 'death', 'deaths', 'death rate'],
                
                // 性別相關
                gender: ['性別', '男女', '男/女', '男女別', 'gender', 'sex'],
                
                // 男性相關
                male: ['男', '男性', '男性人口', '男性人數', 'male', 'men'],
                
                // 女性相關
                female: ['女', '女性', '女性人口', '女性人數', 'female', 'women']
            };
                
                // 標準化縣市名稱 (處理台/臺的差異)
                function standardizeCity(cityName) {
                    if (!cityName) return '';
                    
                    // 統一使用「臺」取代「台」
                    cityName = cityName.replace('台北市', '臺北市')
                                     .replace('台中市', '臺中市')
                                     .replace('台南市', '臺南市')
                                     .replace('台東縣', '臺東縣');
                    
                    return cityName;
                }
                
                detectedYear = null;
                detectedCities = [];
                
                // 將所有數據轉成字符串進行搜索
                var allText = '';
                for (var i = 0; i < jsonData.length; i++) {
                    if (jsonData[i]) {
                        allText += jsonData[i].join(' ') + ' ';
                    }
                }
                
                // 搜索檔名中的年份
                var yearFromFilename = fileName.match(/(\d+)年/);
                if (yearFromFilename) {
                    var yearNum = parseInt(yearFromFilename[1]);
                    // 判斷是否為民國年
                    if (yearNum < 200) {
                        detectedYear = {
                            original: yearNum,
                            type: '民國',
                            westernYear: yearNum + 1911
                        };
                    } else {
                        detectedYear = {
                            original: yearNum,
                            type: '西元',
                            westernYear: yearNum
                        };
                    }
                }
                
                // 搜索文件內容中的年份
                if (!detectedYear) {
                    // 搜索民國年寫法
                    var rocYearMatch = allText.match(/民國\s*(\d+)\s*年/);
                    if (rocYearMatch) {
                        var rocYear = parseInt(rocYearMatch[1]);
                        detectedYear = {
                            original: rocYear,
                            type: '民國',
                            westernYear: rocYear + 1911
                        };
                    } else {
                        // 搜索西元年寫法 (四位數)
                        var westernYearMatch = allText.match(/(19|20)\d{2}/);
                        if (westernYearMatch) {
                            var westernYear = parseInt(westernYearMatch[0]);
                            detectedYear = {
                                original: westernYear,
                                type: '西元',
                                westernYear: westernYear
                            };
                        }
                    }
                }
                
                // 搜索台灣縣市
                for (var i = 0; i < taiwanCities.length; i++) {
                    var city = taiwanCities[i];
                    if (allText.includes(city) && !detectedCities.includes(city)) {
                        detectedCities.push(city);
                    }
                }
                
                // 更新UI顯示偵測結果
                updateDetectedInfo();
            }
            
            // 更新偵測到的資訊
            function updateDetectedInfo() {
                if (detectedYear) {
                    yearValue.textContent = detectedYear.type + detectedYear.original + '年 (西元' + detectedYear.westernYear + '年)';
                } else {
                    yearValue.textContent = '未偵測到';
                }
                
                if (detectedCities.length > 0) {
                    cityValue.textContent = detectedCities.join(', ');
                } else {
                    cityValue.textContent = '未偵測到';
                }
                
                // 顯示偵測資訊區塊
                detectedInfo.style.display = 'block';
            }
            
            // 生成整理後的資料
            function generateStructuredData() {
                // 獲取第一個工作表
                var firstSheetName = unmergedWorkbook.SheetNames[0];
                var worksheet = unmergedWorkbook.Sheets[firstSheetName];
                
                // 將工作表轉換為JSON
                var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                // 整理後的資料
                structuredData = [];
                
                // 檢查是否有任何標題行包含「縣市別」、「縣市」或相關名稱
                var cityColumnIndex = -1;
                var districtColumnIndex = -1;
                var populationColumnIndex = -1;
                var householdsColumnIndex = -1;
                
                // 搜尋可能的標題行
                for (var i = 0; i < Math.min(10, jsonData.length); i++) {
                    var row = jsonData[i];
                    if (!row || row.length < 2) continue;
                    
                    for (var j = 0; j < row.length; j++) {
                        var cellValue = String(row[j] || '').trim();
                        
                        // 尋找縣市相關欄位
                        if (cellValue.includes('縣市') || cellValue === '縣市別' || 
                            cellValue === '縣市' || cellValue === '城市') {
                            cityColumnIndex = j;
                        }
                        
                        // 尋找鄉鎮市區相關欄位
                        if (cellValue.includes('鄉鎮') || cellValue.includes('區域') || 
                            cellValue === '鄉鎮市區' || cellValue === '行政區') {
                            districtColumnIndex = j;
                        }
                        
                        // 尋找人口數相關欄位
                        if (cellValue.includes('人口') || cellValue === '人數' || 
                            cellValue === '總人口' || cellValue === '總數') {
                            populationColumnIndex = j;
                        }
                        
                        // 尋找戶數相關欄位
                        if (cellValue.includes('戶數') || cellValue === '家戶' || 
                            cellValue === '戶' || cellValue === '總戶數') {
                            householdsColumnIndex = j;
                        }
                    }
                    
                    // 如果至少找到了縣市或鄉鎮欄位，則認為這是有效的標題行
                    if (cityColumnIndex >= 0 || districtColumnIndex >= 0) {
                        break;
                    }
                }
                
                // 如果找到了標題行中的欄位索引，嘗試直接從表格結構提取資料
                if (cityColumnIndex >= 0 || districtColumnIndex >= 0) {
                    // 找到標題行後的資料列
                    var dataStartIndex = i + 1;
                    
                    // 用於追蹤目前處理的縣市
                    var currentCity = null;
                    var currentCityData = null;
                    
                    // 處理資料列
                    for (var j = dataStartIndex; j < jsonData.length; j++) {
                        var dataRow = jsonData[j];
                        if (!dataRow || dataRow.length === 0) continue;
                        
                        // 讀取資料
                        var cityValue = cityColumnIndex >= 0 && dataRow[cityColumnIndex] ? String(dataRow[cityColumnIndex]) : null;
                        var districtValue = districtColumnIndex >= 0 && dataRow[districtColumnIndex] ? String(dataRow[districtColumnIndex]) : null;
                        var populationValue = populationColumnIndex >= 0 ? dataRow[populationColumnIndex] : null;
                        var householdsValue = householdsColumnIndex >= 0 ? dataRow[householdsColumnIndex] : null;
                        
                        // 如果有縣市值，且不同於目前處理的縣市，則建立新的縣市資料
                        if (cityValue && cityValue !== currentCity && 
                            !cityValue.includes('總計') && 
                            !cityValue.includes('合計')) {
                            
                            // 檢查是否為有效的縣市名稱
                            if (taiwanCities.some(c => cityValue.includes(c))) {
                                currentCity = cityValue;
                                currentCityData = {
                                    city: currentCity,
                                    year: detectedYear ? detectedYear.westernYear : null,
                                    districts: []
                                };
                                structuredData.push(currentCityData);
                            }
                        }
                        
                        // 如果目前有處理中的縣市，且有有效的區域值，則添加區域資料
                        if (currentCityData && districtValue && 
                            !districtValue.includes('總計') && 
                            !districtValue.includes('合計')) {
                            
                            currentCityData.districts.push({
                                district: districtValue,
                                households: householdsValue,
                                population: populationValue
                            });
                        }
                    }
                }
                
                // 如果從表格結構無法提取資料，則嘗試用縣市名稱搜尋方式
                if (structuredData.length === 0 && detectedCities.length > 0) {
                    // 嘗試找出每個縣市的資料行
                    for (var c = 0; c < detectedCities.length; c++) {
                        var city = detectedCities[c];
                        
                        // 尋找包含縣市名稱的行
                        for (var i = 0; i < jsonData.length; i++) {
                            var row = jsonData[i];
                            if (row && row.length > 0) {
                                var firstCol = String(row[0] || '');
                                
                                // 如果找到縣市
                                if (firstCol.includes(city)) {
                                    // 建立縣市資料結構
                                    var cityData = {
                                        city: standardizeCity(city),
                                        year: detectedYear ? detectedYear.westernYear : null,
                                        districts: []
                                    };
                                    
                                    // 嘗試找出該縣市下的區域/鄉鎮資料
                                    for (var j = i + 1; j < jsonData.length; j++) {
                                        var districtRow = jsonData[j];
                                        
                                        // 跳過空行
                                        if (!districtRow || districtRow.length === 0) continue;
                                        
                                        var districtName = String(districtRow[0] || '');
                                        
                                        // 如果遇到另一個縣市或結尾標記，結束當前縣市處理
                                        if (taiwanCities.some(c => districtName.includes(c)) || 
                                            districtName.includes('總計') && j > i + 1) {
                                            break;
                                        }
                                        
                                        // 排除不是區域名稱的行（通常區域名稱含有「區」、「鄉」、「鎮」、「市」字）
                                        if (districtName.includes('區') || 
                                            districtName.includes('鄉') || 
                                            districtName.includes('鎮') || 
                                            (districtName.includes('市') && districtName.length <= 4)) {
                                            
                                            // 判斷數據欄位位置，取得適當的數值
                                            var population = null;
                                            var households = null;
                                            
                                            // 假設第2列是戶數，第3/4列是人口數
                                            if (districtRow.length > 1) {
                                                households = districtRow[1];
                                            }
                                            
                                            if (districtRow.length > 2) {
                                                // 假設第3欄是總人口或男性人口，第4欄是女性人口
                                                var totalPopulation = districtRow[2];
                                                
                                                // 如果有第4欄，可能是男女人口分開列的情況
                                                if (districtRow.length > 3) {
                                                    var malePopulation = districtRow[2];
                                                    var femalePopulation = districtRow[3];
                                                    
                                                    // 檢查是否為數字，如果是則相加得到總人口
                                                    if (!isNaN(malePopulation) && !isNaN(femalePopulation)) {
                                                        totalPopulation = malePopulation + femalePopulation;
                                                    }
                                                }
                                                
                                                population = totalPopulation;
                                            }
                                            
                                            // 添加區域資料
                                            cityData.districts.push({
                                                district: districtName,
                                                households: households,
                                                population: population
                                            });
                                        }
                                    }
                                    
                                    // 添加縣市資料到結構化資料中
                                    structuredData.push(cityData);
                                    break; // 找到縣市後跳出循環，進行下一個縣市的搜索
                                }
                            }
                        }
                    }
                }
                
                // 創建並顯示整理好的資料表格
                displayStructuredData();
            }
            
            // 將資料表轉換為標準格式（一列一列的格式，每欄都有標準化的標題）
            function displayStructuredData() {
                // 清空處理後表格
                processedTable.innerHTML = '';
                
                if (structuredData.length === 0) {
                    processedTable.innerHTML = '<div class="info-panel" style="background-color: #fadbd8;">未能生成整理後資料，請確認檔案包含縣市及區域資訊。</div>';
                    return;
                }
                
                // 根據資料中的欄位，動態生成標題列
                var availableColumns = identifyAvailableColumns(structuredData);
                
                // 創建表頭
                var thead = document.createElement('thead');
                var headerRow = document.createElement('tr');
                
                // 添加標題列
                availableColumns.forEach(function(column) {
                    var th = document.createElement('th');
                    th.textContent = column.label;
                    headerRow.appendChild(th);
                });
                
                thead.appendChild(headerRow);
                processedTable.appendChild(thead);
                
                // 創建表身
                var tbody = document.createElement('tbody');
                
                // 將每個縣市和區域的資料展開為單獨的行
                var flattenedData = [];
                
                for (var i = 0; i < structuredData.length; i++) {
                    var cityData = structuredData[i];
                    
                    // 處理縣市總計行
                    var totalHouseholds = 0;
                    var totalPopulation = 0;
                    var totalBirth = 0;
                    var totalDeath = 0;
                    
                    cityData.districts.forEach(function(district) {
                        if (!isNaN(district.households)) totalHouseholds += parseFloat(district.households);
                        if (!isNaN(district.population)) totalPopulation += parseFloat(district.population);
                        if (!isNaN(district.birth)) totalBirth += parseFloat(district.birth);
                        if (!isNaN(district.death)) totalDeath += parseFloat(district.death);
                        
                        // 添加每個鄉鎮市區的數據
                        var rowData = {
                            year: cityData.year,
                            city: standardizeCity(cityData.city),
                            district: district.district,
                            households: district.households,
                            population: district.population,
                            birth: district.birth,
                            death: district.death
                        };
                        
                        // 如果有性別資料，添加性別相關欄位
                        if (district.gender) rowData.gender = district.gender;
                        if (district.male) rowData.male = district.male;
                        if (district.female) rowData.female = district.female;
                        
                        flattenedData.push(rowData);
                    });
                    
                    // 添加縣市總計行
                    var totalRowData = {
                        year: cityData.year,
                        city: standardizeCity(cityData.city),
                        district: '總計',
                        households: totalHouseholds,
                        population: totalPopulation,
                        birth: totalBirth,
                        death: totalDeath
                    };
                    
                    flattenedData.push(totalRowData);
                }
                
                // 添加數據行
                flattenedData.forEach(function(rowData) {
                    var tr = document.createElement('tr');
                    
                    // 標記總計行的樣式
                    if (rowData.district === '總計') {
                        tr.style.backgroundColor = '#eaf2f8';
                        tr.style.fontWeight = 'bold';
                    }
                    
                    // 添加每個欄位的數據
                    availableColumns.forEach(function(column) {
                        var td = document.createElement('td');
                        var value = rowData[column.field];
                        
                        // 根據欄位類型進行格式化
                        if (value !== null && value !== undefined) {
                            if (column.type === 'number' && !isNaN(value)) {
                                td.textContent = parseFloat(value).toLocaleString();
                            } else {
                                td.textContent = value;
                            }
                        } else {
                            td.textContent = '-';
                        }
                        
                        tr.appendChild(td);
                    });
                    
                    tbody.appendChild(tr);
                });
                
                processedTable.appendChild(tbody);
                
                // 保存轉換後的數據，用於匯出
                processedData = flattenedData;
                
                // 顯示成功訊息
                showSuccess('已成功生成整理後資料，共 ' + flattenedData.length + ' 筆記錄');
            }
            
            // 識別資料中有哪些可用的欄位
            function identifyAvailableColumns(data) {
                var columns = [];
                var hasHouseholds = false;
                var hasPopulation = false;
                var hasBirth = false;
                var hasDeath = false;
                var hasGender = false;
                var hasMale = false;
                var hasFemale = false;
                
                // 檢查每個縣市的資料
                for (var i = 0; i < data.length; i++) {
                    var cityData = data[i];
                    
                    // 檢查每個區域的資料
                    for (var j = 0; j < cityData.districts.length; j++) {
                        var district = cityData.districts[j];
                        
                        if (district.households !== null && district.households !== undefined) hasHouseholds = true;
                        if (district.population !== null && district.population !== undefined) hasPopulation = true;
                        if (district.birth !== null && district.birth !== undefined) hasBirth = true;
                        if (district.death !== null && district.death !== undefined) hasDeath = true;
                        if (district.gender !== null && district.gender !== undefined) hasGender = true;
                        if (district.male !== null && district.male !== undefined) hasMale = true;
                        if (district.female !== null && district.female !== undefined) hasFemale = true;
                    }
                }
                
                // 添加必要的欄位
                columns.push({ field: 'year', label: '年份', type: 'text' });
                columns.push({ field: 'city', label: '縣市別', type: 'text' });
                columns.push({ field: 'district', label: '鄉鎮市區', type: 'text' });
                
                // 根據實際資料添加其他欄位
                if (hasHouseholds) columns.push({ field: 'households', label: '戶數', type: 'number' });
                if (hasPopulation) columns.push({ field: 'population', label: '人口數', type: 'number' });
                if (hasBirth) columns.push({ field: 'birth', label: '出生數', type: 'number' });
                if (hasDeath) columns.push({ field: 'death', label: '死亡數', type: 'number' });
                if (hasGender) columns.push({ field: 'gender', label: '性別', type: 'text' });
                if (hasMale) columns.push({ field: 'male', label: '男性人口', type: 'number' });
                if (hasFemale) columns.push({ field: 'female', label: '女性人口', type: 'number' });
                
                return columns;
            }
            
            // 解析 CSV 檔案
            function parseCSVFile(csvContent) {
                // 使用 Papa Parse 解析 CSV
                Papa.parse(csvContent, {
                    header: false,
                    dynamicTyping: true,
                    skipEmptyLines: true,
                    complete: function(results) {
                        // 獲取解析後的資料
                        var jsonData = results.data;
                        
                        // 從CSV建立虛擬工作簿
                        var workbook = XLSX.utils.book_new();
                        var worksheet = XLSX.utils.aoa_to_sheet(jsonData);
                        XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
                        
                        // 保存為原始工作簿
                        originalWorkbook = workbook;
                        
                        // 顯示原始預覽
                        displayOriginalPreview(worksheet);
                        
                        // 解析後的CSV已經沒有合併儲存格，所以解除合併後和原始相同
                        unmergedWorkbook = workbook;
                        displayUnmergedPreview(worksheet);
                        
                        // 識別標題列並處理
                        displayProcessedPreview(worksheet);
                        
                        // 顯示成功訊息
                        showSuccess('CSV檔案解析成功！');
                        
                        // 顯示結果區段
                        resultSection.style.display = 'block';
                    },
                    error: function(error) {
                        showError('解析 CSV 時發生錯誤: ' + error.message);
                    }
                });
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