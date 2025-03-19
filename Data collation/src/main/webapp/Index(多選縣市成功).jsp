<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.sql.*"%>
<jsp:useBean id='objDBConfig' scope='application' class='hitstd.group.tool.database.DBConfig' />

<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="icon" href="img/documentation.png" type="image/png">
    <title>縣市鄉鎮人口統計資料處理器</title>
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
            background-color: #1A76D1;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            margin-right: 5px;
        }
        button:hover {
        	color: white;
            background-color: #76ace3;
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
            background-color: #629ecc;
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
            background-color: #c4dff5;
            cursor: pointer;
            margin-right: 5px;
            border-radius: 4px 4px 0 0;
        }
        .tab-button.active {
            background-color: #1A76D1;
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
            color: #1c4e80;
            background-color: #dcebf7;
            padding: 10px;
            border-radius: 4px;
            margin: 10px 0;
            display: none;
        }
        .info-panel {
            background-color: #dcebf7;
            border-left: 4px solid #1A76D1;
            padding: 10px 15px;
            margin: 15px 0;
            font-size: 14px;
        }
        #debugPanel {
            background-color: #f0f0f0; 
            padding: 10px; 
            margin-top: 20px; 
            border-radius: 5px;
            display: none;
        }
        .log-item {
            border-bottom: 1px solid #eee;
            padding: 5px 0;
        }
        .error-item {
            border-bottom: 1px solid #eee;
            padding: 5px 0;
            color: #e74c3c;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>縣市鄉鎮人口統計資料處理器</h1>
        
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
        <!-- 工作表選擇區域 -->
		<div id="sheetSelector" class="sheet-selector" style="display: none; margin: 15px 0; padding: 15px; background-color: #f0f7fb; border-radius: 4px; border-left: 4px solid #1A76D1;">
		    <h2><span class="step-number" style="display: inline-block; width: 24px; height: 24px; background-color: #1A76D1; color: white; border-radius: 50%; text-align: center; line-height: 24px; margin-right: 8px;">2</span>選擇工作表</h2>
		    <div class="info-panel" style="background-color: #dcebf7; border-left: 4px solid #1A76D1; padding: 10px 15px; margin: 15px 0; font-size: 14px;">
		        此Excel檔案包含多個工作表，請選擇要處理的工作表（可多選）。
		    </div>
		    <button id="selectAllSheets" class="select-all-btn" style="margin-right: 10px; margin-bottom: 10px; padding: 5px 10px; background-color: #76ace3; color: white; border: none; border-radius: 4px; cursor: pointer;">全選</button>
		    <div id="sheetList" class="sheet-list" style="margin: 10px 0; display: flex; flex-wrap: wrap; gap: 10px;"></div>
		    <button id="processSelectedSheet" class="sheet-process-btn" style="display: none; margin-top: 10px; padding: 8px 15px; background-color: #1A76D1; color: white; border: none; border-radius: 4px; cursor: pointer;">處理選取的工作表</button>
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
					            <th>工作表名稱</th>
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
        // 引入 Papa Parse 庫 (用於解析 CSV)
        var papaParseScript = document.createElement('script');
        papaParseScript.src = 'https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.3.2/papaparse.min.js';
        document.head.appendChild(papaParseScript);
        
        // 移除字串中的特殊符號「※」
        function removeSpecialSymbol(str) {
            if (typeof str !== 'string') return str;
            return str.replace(/※/g, '');
        }
        
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
        
        // 工作表選擇相關元素 - 需要在HTML中添加這些元素
        var sheetSelector = document.getElementById('sheetSelector') || createSheetSelector();
        var sheetList = document.getElementById('sheetList') || document.querySelector('.sheet-list');
        var processSelectedSheetBtn = document.getElementById('processSelectedSheet') || document.querySelector('.sheet-process-btn');
        var selectAllSheetsBtn = document.getElementById('selectAllSheets') || document.querySelector('.select-all-btn');
        
        // 創建工作表選擇器（如果不存在）
        function createSheetSelector() {
            var selector = document.createElement('div');
            selector.id = 'sheetSelector';
            selector.className = 'sheet-selector';
            selector.style.display = 'none';
            selector.style.margin = '15px 0';
            selector.style.padding = '15px';
            selector.style.backgroundColor = '#f0f7fb';
            selector.style.borderRadius = '4px';
            selector.style.borderLeft = '4px solid #1A76D1';
            
            // 添加標題
            var title = document.createElement('h2');
            title.innerHTML = '<span class="step-number">2</span>選擇工作表';
            selector.appendChild(title);
            
            // 添加信息面板
            var infoPanel = document.createElement('div');
            infoPanel.className = 'info-panel';
            infoPanel.textContent = '此Excel檔案包含多個工作表，請選擇要處理的工作表（可多選）。';
            selector.appendChild(infoPanel);
            
            // 添加全選按鈕
            var selectAllBtn = document.createElement('button');
            selectAllBtn.id = 'selectAllSheets';
            selectAllBtn.className = 'select-all-btn';
            selectAllBtn.textContent = '全選';
            selectAllBtn.style.marginRight = '10px';
            selectAllBtn.style.marginBottom = '10px';
            selectAllBtn.style.padding = '5px 10px';
            selectAllBtn.style.backgroundColor = '#76ace3';
            selectAllBtn.style.color = 'white';
            selectAllBtn.style.border = 'none';
            selectAllBtn.style.borderRadius = '4px';
            selectAllBtn.style.cursor = 'pointer';
            selector.appendChild(selectAllBtn);
            
            // 添加工作表列表
            var list = document.createElement('div');
            list.id = 'sheetList';
            list.className = 'sheet-list';
            list.style.margin = '10px 0';
            list.style.display = 'flex';
            list.style.flexWrap = 'wrap';
            list.style.gap = '10px';
            selector.appendChild(list);
            
            // 添加處理按鈕
            var processBtn = document.createElement('button');
            processBtn.id = 'processSelectedSheet';
            processBtn.className = 'sheet-process-btn';
            processBtn.textContent = '處理選取的工作表';
            processBtn.style.display = 'none';
            processBtn.style.marginTop = '10px';
            processBtn.style.padding = '8px 15px';
            processBtn.style.backgroundColor = '#1A76D1';
            processBtn.style.color = 'white';
            processBtn.style.border = 'none';
            processBtn.style.borderRadius = '4px';
            processBtn.style.cursor = 'pointer';
            selector.appendChild(processBtn);
            
            // 將選擇器插入到適當位置
            var uploadSection = document.querySelector('.upload-section');
            if (uploadSection) {
                uploadSection.parentNode.insertBefore(selector, uploadSection.nextSibling);
            } else {
                document.querySelector('.container').appendChild(selector);
            }
            
            return selector;
        }
        
        // 存儲解析後的資料
        var originalWorkbook = null;
        var unmergedWorkbook = null;
        var parsedData = [];
        var fileName = '';
        var fileYear = '';
        var fileMonth = ''; 
        var selectedSheets = []; // 存儲選中的工作表名稱
        
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
        
        // 在解析前預處理 Excel 數據，確保數值欄位為數值格式
        function preprocessDataFormat(workbook) {
            console.log('開始預處理數值欄位格式...');
            // 對每個工作表進行處理
            workbook.SheetNames.forEach(function(sheetName) {
                var worksheet = workbook.Sheets[sheetName];
                var range = XLSX.utils.decode_range(worksheet['!ref']);
                
                // 尋找對應的數值欄位 (通常是第2欄、第4欄和第5欄)
                for (var r = 0; r <= range.e.r; r++) {
                    // 處理第2欄 (戶數)
                    var cellAddress2 = XLSX.utils.encode_cell({r: r, c: 1});
                    if (worksheet[cellAddress2] && typeof worksheet[cellAddress2].v === 'string') {
                        var numValue = parseFloat(worksheet[cellAddress2].v.replace(/,/g, '').trim());
                        if (!isNaN(numValue)) {
                            console.log('將第2欄儲存格從文字轉為數值:', worksheet[cellAddress2].v, '->', numValue);
                            worksheet[cellAddress2].v = numValue;
                            worksheet[cellAddress2].t = 'n'; // 設定為數值類型
                        }
                    }
                    
                    // 處理第4欄 (男性人口)
                    var cellAddress4 = XLSX.utils.encode_cell({r: r, c: 3});
                    if (worksheet[cellAddress4] && typeof worksheet[cellAddress4].v === 'string') {
                        var numValue = parseFloat(worksheet[cellAddress4].v.replace(/,/g, '').trim());
                        if (!isNaN(numValue)) {
                            console.log('將第4欄儲存格從文字轉為數值:', worksheet[cellAddress4].v, '->', numValue);
                            worksheet[cellAddress4].v = numValue;
                            worksheet[cellAddress4].t = 'n'; // 設定為數值類型
                        }
                    }
                    
                    // 處理第5欄 (女性人口)
                    var cellAddress5 = XLSX.utils.encode_cell({r: r, c: 4});
                    if (worksheet[cellAddress5] && typeof worksheet[cellAddress5].v === 'string') {
                        var numValue = parseFloat(worksheet[cellAddress5].v.replace(/,/g, '').trim());
                        if (!isNaN(numValue)) {
                            console.log('將第5欄儲存格從文字轉為數值:', worksheet[cellAddress5].v, '->', numValue);
                            worksheet[cellAddress5].v = numValue;
                            worksheet[cellAddress5].t = 'n'; // 設定為數值類型
                        }
                    }
                }
            });
            
            console.log('數值欄位預處理完成');
            return workbook;
        }
        
        // 解析按鈕點擊事件
        parseBtn.addEventListener('click', function() {
            var file = excelFileInput.files[0];
            if (!file) {
                showError('請先選擇檔案');
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
            
            // 隱藏先前的結果
            resultSection.style.display = 'none';
            sheetSelector.style.display = 'none';
            
            // 重置選中的工作表
            selectedSheets = [];
            
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
                        
                        // 讀取Excel檔案
                        originalWorkbook = XLSX.read(data, { type: 'array' });
                        
                        // 預處理數值欄位格式
                        originalWorkbook = preprocessDataFormat(originalWorkbook);
                        
                        // 檢查工作表數量
                        if (originalWorkbook.SheetNames.length > 1) {
                            // 如果有多個工作表，顯示工作表選擇界面
                            displaySheetSelector(originalWorkbook.SheetNames);
                        } else {
                            // 如果只有一個工作表，直接處理
                            selectedSheets = [originalWorkbook.SheetNames[0]];
                            processExcelSheets(selectedSheets);
                        }
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
        
        // 顯示工作表選擇器
        function displaySheetSelector(sheetNames) {
            // 清空工作表列表
            sheetList.innerHTML = '';
            
            // 添加每個工作表選項
            sheetNames.forEach(function(sheetName) {
                var sheetOption = document.createElement('div');
                sheetOption.className = 'sheet-option';
                sheetOption.textContent = sheetName;
                sheetOption.dataset.sheetName = sheetName;
                
                // 為工作表選項添加樣式
                sheetOption.style.display = 'inline-block';
                sheetOption.style.padding = '8px 15px';
                sheetOption.style.backgroundColor = '#e0eefa';
                sheetOption.style.border = '1px solid #1A76D1';
                sheetOption.style.borderRadius = '4px';
                sheetOption.style.cursor = 'pointer';
                sheetOption.style.transition = 'all 0.2s';
                
                // 點擊工作表選項時的處理
                sheetOption.addEventListener('click', function() {
                    // 切換選中狀態
                    if (this.classList.contains('selected')) {
                        // 取消選中
                        this.classList.remove('selected');
                        this.style.backgroundColor = '#e0eefa';
                        this.style.color = '';
                        
                        // 從選中列表中移除
                        var index = selectedSheets.indexOf(this.dataset.sheetName);
                        if (index !== -1) {
                            selectedSheets.splice(index, 1);
                        }
                    } else {
                        // 選中
                        this.classList.add('selected');
                        this.style.backgroundColor = '#1A76D1';
                        this.style.color = 'white';
                        
                        // 添加到選中列表
                        selectedSheets.push(this.dataset.sheetName);
                    }
                    
                    // 更新處理按鈕的顯示狀態
                    if (selectedSheets.length > 0) {
                        processSelectedSheetBtn.style.display = 'inline-block';
                        processSelectedSheetBtn.textContent = '處理選取的 ' + selectedSheets.length + ' 個工作表';
                    } else {
                        processSelectedSheetBtn.style.display = 'none';
                    }
                });
                
                sheetList.appendChild(sheetOption);
            });
            
            // 顯示工作表選擇區域
            sheetSelector.style.display = 'block';
            
            // 清空處理按鈕的現有事件
            var newProcessBtn = processSelectedSheetBtn.cloneNode(true);
            processSelectedSheetBtn.parentNode.replaceChild(newProcessBtn, processSelectedSheetBtn);
            processSelectedSheetBtn = newProcessBtn;
            
            // 全選按鈕事件
            if (selectAllSheetsBtn) {
                // 清空全選按鈕的現有事件
                var newSelectAllBtn = selectAllSheetsBtn.cloneNode(true);
                selectAllSheetsBtn.parentNode.replaceChild(newSelectAllBtn, selectAllSheetsBtn);
                selectAllSheetsBtn = newSelectAllBtn;
                
                selectAllSheetsBtn.addEventListener('click', function() {
                    var allSelected = (selectedSheets.length === sheetNames.length);
                    
                    if (allSelected) {
                        // 如果全部都已經選中，則全部取消選中
                        selectedSheets = [];
                        var options = document.querySelectorAll('.sheet-option');
                        options.forEach(function(option) {
                            option.classList.remove('selected');
                            option.style.backgroundColor = '#e0eefa';
                            option.style.color = '';
                        });
                        this.textContent = '全選';
                        processSelectedSheetBtn.style.display = 'none';
                    } else {
                        // 否則全部選中
                        selectedSheets = [...sheetNames];
                        var options = document.querySelectorAll('.sheet-option');
                        options.forEach(function(option) {
                            option.classList.add('selected');
                            option.style.backgroundColor = '#1A76D1';
                            option.style.color = 'white';
                        });
                        this.textContent = '取消全選';
                        processSelectedSheetBtn.style.display = 'inline-block';
                        processSelectedSheetBtn.textContent = '處理選取的 ' + selectedSheets.length + ' 個工作表';
                    }
                });
            }
            
            // 添加處理選中工作表的按鈕事件
            processSelectedSheetBtn.addEventListener('click', function() {
                if (selectedSheets.length === 0) {
                    showError('請至少選擇一個工作表');
                    return;
                }
                
                // 處理選中的工作表
                processExcelSheets(selectedSheets);
            });
        }
        
        // 處理多個Excel工作表
        function processExcelSheets(sheetNames) {
            try {
                console.log('處理工作表:', sheetNames);
                
                // 重置已解析的數據
                parsedData = [];
                
                // 處理每個選中的工作表
                for (var i = 0; i < sheetNames.length; i++) {
                    var sheetName = sheetNames[i];
                    
                    // 更新進度
                    updateProgressBar(33 + (i / sheetNames.length * 33));
                    
                    // 獲取選定的工作表
                    var worksheet = originalWorkbook.Sheets[sheetName];
                    
                    // 如果是第一個工作表，顯示原始Excel預覽
                    if (i === 0) {
                        displayOriginalPreview(worksheet);
                    }
                    
                    // 解除合併儲存格
                    if (i === 0) {
                        updateProgressBar(66);
                        unmergedWorkbook = unmergeCells(originalWorkbook);
                    }
                    var unmergedWorksheet = unmergedWorkbook.Sheets[sheetName];
                    
                    // 如果是第一個工作表，顯示解除合併後的Excel預覽
                    if (i === 0) {
                        displayUnmergedPreview(unmergedWorksheet);
                    }
                    
                    // 從解除合併儲存格後的工作表提取資料
                    var jsonData = XLSX.utils.sheet_to_json(unmergedWorksheet, { header: 1 });
                    
                    // 提取資料 (根據明確指示)
                    var extractedData = extractDataFromSpecificFormat(jsonData, sheetName);
                    
                    // 添加到總數據
                    parsedData = parsedData.concat(extractedData);
                }
                
                // 顯示處理後資料
                displayProcessedData(parsedData);
                
                // 隱藏工作表選擇區域
                sheetSelector.style.display = 'none';
                
                // 顯示結果區域
                resultSection.style.display = 'block';
                
                showSuccess('成功處理 ' + sheetNames.length + ' 個工作表，共 ' + parsedData.length + ' 筆記錄');
            } catch (error) {
                showError('處理工作表時發生錯誤: ' + error.message);
                console.error(error);
            }
        }
        
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
                var wsData = [['年分', '月份', '縣市別', '鄉鎮別', '性別', '戶數', '人口數', '工作表名稱']];
                
                // 添加資料行
                parsedData.forEach(function(row) {
                    wsData.push([
                        row.year,
                        row.month,
                        row.city,
                        row.district,
                        row.gender,
                        row.households,
                        row.population,
                        row.sheetName || ''
                    ]);
                });
                
                // 將數據轉換為工作表
                var ws = XLSX.utils.aoa_to_sheet(wsData);
                
                // 添加工作表到工作簿
                XLSX.utils.book_append_sheet(wb, ws, '鄉鎮人口統計');
                
                // 導出為XLSX文件
                var exportFileName = '人口統計';
                if (fileYear && fileMonth) {
                    exportFileName += '_' + fileYear + '年' + fileMonth + '月';
                }
                if (selectedSheets.length === 1) {
                    exportFileName += '_' + selectedSheets[0];
                } else if (selectedSheets.length > 1) {
                    exportFileName += '_多表合併';
                }
                exportFileName += '.xlsx';
                
                XLSX.writeFile(wb, exportFileName);
                showSuccess('已成功匯出Excel檔案: ' + exportFileName);
            } catch (error) {
                showError('匯出Excel時發生錯誤: ' + error.message);
                console.error(error);
            }
        });
        
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
            
            // 顯示所有行數據，不限制行數
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
            
            // 顯示所有行數據，不限制行數
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
        
        // 判斷是否為備註或說明行的函數
        function isNoteOrDescription(row) {
            // 檢查是否為空行
            if (!row || row.length === 0) return false;
            
            // 檢查第一格是否包含常見的備註標記
            var firstCell = row[0];
            if (firstCell === null || firstCell === undefined) return false;
            
            // 轉換為字串以進行檢查
            if (typeof firstCell !== 'string') {
                firstCell = String(firstCell);
            }
            
            // 檢查是否為備註或說明行
            if (firstCell.includes('備註') || 
                firstCell.includes('說明') || 
                firstCell.includes('註') || 
                firstCell.includes('附註') || 
                firstCell.includes('注意') || 
                firstCell.includes('說') || 
                firstCell.includes('資料來源')) {
                return true;
            }
            
            // 檢查是否整行僅有文字說明，沒有數值資料
            var hasNumericData = false;
            for (var i = 1; i < row.length; i++) {
                if (row[i] !== null && row[i] !== undefined) {
                    // 檢查是否為數值
                    if (typeof row[i] === 'number' || 
                        (typeof row[i] === 'string' && !isNaN(parseFloat(row[i].replace(/,/g, ''))))) {
                        hasNumericData = true;
                        break;
                    }
                }
            }
            
            // 如果沒有數值資料，且有文字資料，可能是說明行
            return !hasNumericData && firstCell.length > 0;
        }
        
        // 根據特定格式提取資料 - 支持多工作表版本
        function extractDataFromSpecificFormat(jsonData, sheetName) {
            var result = [];
            
            try {
                // 尋找標題行中包含「中華民國」的行以確定資料開始的位置
                var titleRowIndex = -1;
                for (var i = 0; i < Math.min(10, jsonData.length); i++) {
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
                
                // 如果沒找到標題行，嘗試從第三列找中華民國的資訊
                if (titleRowIndex === -1) {
                    for (var i = 0; i < Math.min(10, jsonData.length); i++) {
                        var row = jsonData[i];
                        if (row && row[2] && String(row[2]).includes('中華民國')) {
                            titleRowIndex = i;
                            
                            // 嘗試從標題中提取年月
                            if (!fileYear || !fileMonth) {
                                var titleText = String(row[2]);
                                var yearMonthMatch = titleText.match(/(\d+)年(\d+)月/);
                                if (yearMonthMatch) {
                                    fileYear = yearMonthMatch[1];
                                    fileMonth = yearMonthMatch[2];
                                }
                            }
                            break;
                        }
                    }
                }
                
                // 如果仍然沒找到標題行，就假設資料從第一行開始
                if (titleRowIndex === -1) {
                    titleRowIndex = 0;
                }
                
                // 找尋縣市資料列的索引
                var cityRowIndex = -1;
                var cityName = '';
                
                // 定義所有可能的縣市名稱
                const cityNames = [
                    '新北市', '臺北市', '桃園市', '臺中市', '臺南市', '高雄市', 
                    '宜蘭縣', '新竹縣', '苗栗縣', '彰化縣', '南投縣', '雲林縣', 
                    '嘉義縣', '屏東縣', '臺東縣', '花蓮縣', '澎湖縣', '基隆市', 
                    '新竹市', '嘉義市', '金門縣', '連江縣'
                ];
                
                // 尋找縣市列
                for (var i = 0; i < jsonData.length; i++) {
                    var row = jsonData[i];
                    if (!row || !row[0]) continue;
                    
                    var cellValue = String(row[0]).trim();
                    
                    // 檢查是否為縣市名稱
                    for (var j = 0; j < cityNames.length; j++) {
                        if (cellValue.includes(cityNames[j])) {
                            cityRowIndex = i;
                            cityName = cityNames[j];
                            console.log('找到縣市:', cityName, '在行:', i, '工作表:', sheetName);
                            break;
                        }
                    }
                    
                    if (cityRowIndex !== -1) break;
                }
                
                if (cityRowIndex === -1) {
                    console.log('工作表 ' + sheetName + ' 中無法找到縣市資料行，嘗試使用工作表名稱作為縣市');
                    
                    // 嘗試從工作表名稱中找出縣市名稱
                    for (var j = 0; j < cityNames.length; j++) {
                        if (sheetName.includes(cityNames[j])) {
                            cityName = cityNames[j];
                            console.log('從工作表名稱找到縣市:', cityName);
                            
                            // 尋找包含戶數和人口數的行
                            for (var i = 0; i < jsonData.length; i++) {
                                var row = jsonData[i];
                                if (row && row.length >= 5) {
                                    // 檢查第2欄、第4欄和第5欄是否為數值
                                    if (typeof row[1] === 'number' && typeof row[3] === 'number' && typeof row[4] === 'number') {
                                        cityRowIndex = i;
                                        console.log('找到可能的縣市資料行:', i);
                                        break;
                                    }
                                }
                            }
                            break;
                        }
                    }
                    
                    if (cityRowIndex === -1 || !cityName) {
                        throw new Error('工作表 ' + sheetName + ' 中無法找到縣市資料行');
                    }
                }
                
                console.log('縣市行:', cityRowIndex);
                console.log('縣市名稱:', cityName);
                console.log('該行數據:', jsonData[cityRowIndex]);
                
                // 找尋列標題行 (通常在縣市行之前)
                var headerRowIndex = -1;
                for (var i = Math.max(0, cityRowIndex - 5); i < cityRowIndex; i++) {
                    var row = jsonData[i];
                    if (row && row.length >= 5) {
                        // 檢查是否為包含「縣市別」或「鄉鎮市區」或「戶數」或「人口數」的列
                        var hasRelevantHeader = false;
                        for (var j = 0; j < row.length; j++) {
                            if (row[j] && (
                                String(row[j]).includes('縣市') || 
                                String(row[j]).includes('鄉鎮') || 
                                String(row[j]).includes('戶數') || 
                                String(row[j]).includes('人口')
                            )) {
                                hasRelevantHeader = true;
                                break;
                            }
                        }
                        if (hasRelevantHeader) {
                            headerRowIndex = i;
                            console.log('找到標題行:', i, '內容:', row);
                            break;
                        }
                    }
                }
                
                // 處理縣市總計行
                var cityRow = jsonData[cityRowIndex];
                if (cityRow[1] !== undefined && cityRow[3] !== undefined && cityRow[4] !== undefined) {
                    // 移除特殊符號「※」
                    var cleanCityName = removeSpecialSymbol(cityName);
                    
                    // 確保數值有效
                    var households = isNaN(parseInt(cityRow[1])) ? 0 : parseInt(cityRow[1]);
                    var malePop = isNaN(parseInt(cityRow[3])) ? 0 : parseInt(cityRow[3]);
                    var femalePop = isNaN(parseInt(cityRow[4])) ? 0 : parseInt(cityRow[4]);
                    
                    // 男性資料
                    result.push({
                        year: fileYear,
                        month: fileMonth,
                        city: cleanCityName,
                        district: '總計',
                        gender: '男',
                        households: households,
                        population: malePop,
                        sheetName: sheetName
                    });
                    
                    // 女性資料
                    result.push({
                        year: fileYear,
                        month: fileMonth,
                        city: cleanCityName,
                        district: '總計',
                        gender: '女',
                        households: households,
                        population: femalePop,
                        sheetName: sheetName
                    });
                }
                
                // 從縣市行之後開始處理鄉鎮資料
                var nextCityFound = false;
                for (var i = cityRowIndex + 1; i < jsonData.length && !nextCityFound; i++) {
                    var row = jsonData[i];
                    
                    // 確保行有效且有資料
                    if (!row || !row[0]) continue;
                    
                    var districtName = String(row[0]).trim();
                    
                    // 跳過空行
                    if (!districtName || districtName === '') continue;
                    
                    // 檢查是否為備註或說明行，如果是則跳過
                    if (isNoteOrDescription(row)) {
                        console.log('跳過備註或說明行:', row);
                        continue;
                    }
                    
                    // 檢查是否已經到了下一個縣市（如果有）
                    for (var j = 0; j < cityNames.length; j++) {
                        if (districtName.includes(cityNames[j]) && cityNames[j] !== cityName) {
                            nextCityFound = true;
                            console.log('找到下一個縣市:', cityNames[j], '，停止處理當前縣市資料');
                            break;
                        }
                    }
                    if (nextCityFound) break;
                    
                    // 檢查是否為鄉鎮名稱
                    if (!(districtName.includes('區') || 
                         districtName.includes('鄉') || 
                         districtName.includes('鎮') || 
                         (districtName.includes('市') && districtName.length <= 3))) {
                        // 不是鄉鎮名稱，可能是備註，跳過
                        console.log('不是有效的鄉鎮名稱，跳過:', districtName);
                        continue;
                    }
                    
                    console.log('處理鄉鎮:', districtName, '行數據:', row);
                    
                    // 確保有人口數資料
                    if (row[1] === undefined || row[3] === undefined || row[4] === undefined) {
                        console.log('跳過無效資料行:', row);
                        continue;
                    }
                    
                    // 確保數值有效
                    var households = isNaN(parseInt(row[1])) ? 0 : parseInt(row[1]);
                    var malePop = isNaN(parseInt(row[3])) ? 0 : parseInt(row[3]);
                    var femalePop = isNaN(parseInt(row[4])) ? 0 : parseInt(row[4]);
                    
                    // 移除特殊符號「※」
                    var cleanDistrictName = removeSpecialSymbol(districtName);
                    var cleanCityName = removeSpecialSymbol(cityName);
                    
                    // 男性資料
                    result.push({
                        year: fileYear,
                        month: fileMonth,
                        city: cleanCityName,
                        district: cleanDistrictName,
                        gender: '男',
                        households: households,
                        population: malePop,
                        sheetName: sheetName
                    });
                    
                    // 女性資料
                    result.push({
                        year: fileYear,
                        month: fileMonth,
                        city: cleanCityName,
                        district: cleanDistrictName,
                        gender: '女',
                        households: households,
                        population: femalePop,
                        sheetName: sheetName
                    });
                }
                
                if (result.length === 0) {
                    console.warn('工作表 ' + sheetName + ' 中未能提取到任何有效資料');
                } else {
                    console.log('從工作表 ' + sheetName + ' 中成功提取 ' + result.length + ' 筆資料');
                }
            } catch (error) {
                console.error('提取資料錯誤 (工作表 ' + sheetName + '):', error);
            }
            
            return result;
        }
        
     // 顯示處理後資料
        function displayProcessedData(data) {
            parsedData = data;
            
            // 清空表格
            processedBody.innerHTML = '';
            
            // 更新摘要 - 添加已選擇的工作表名稱
            var summaryText = '共解析出 ' + data.length + ' 筆記錄 (來自 ' + selectedSheets.length + ' 個工作表)';
            
            // 添加已選擇的工作表列表
            if (selectedSheets.length > 0) {
                summaryText += '<div style="margin-top: 10px; font-weight: normal;">已處理的工作表: ';
                var sheetList = selectedSheets.map(function(sheet) {
                    return '<span style="display: inline-block; margin: 2px 5px; padding: 3px 8px; background-color: #e0eefa; border: 1px solid #1A76D1; border-radius: 4px;">' + sheet + '</span>';
                }).join(' ');
                summaryText += sheetList + '</div>';
            }
            
            resultSummary.innerHTML = summaryText;
            
            // 檢查是否需要顯示工作表列
            var showSheetColumn = selectedSheets.length > 1;
            
            // 更新表頭
            var headerRow = document.querySelector('#processedTable thead tr');
            if (headerRow) {
                // 檢查是否已經有工作表列
                var hasSheetColumn = false;
                for (var i = 0; i < headerRow.cells.length; i++) {
                    if (headerRow.cells[i].textContent === '工作表名稱') {
                        hasSheetColumn = true;
                        break;
                    }
                }
                
                // 如果需要顯示工作表列但尚未有
                if (showSheetColumn && !hasSheetColumn) {
                    var sheetTh = document.createElement('th');
                    sheetTh.textContent = '工作表名稱';
                    headerRow.appendChild(sheetTh);
                }
                // 如果不需要顯示工作表列但已存在
                else if (!showSheetColumn && hasSheetColumn) {
                    for (var i = 0; i < headerRow.cells.length; i++) {
                        if (headerRow.cells[i].textContent === '工作表名稱') {
                            headerRow.removeChild(headerRow.cells[i]);
                            break;
                        }
                    }
                }
            }
            
            // 添加所有資料到表格
            for (var i = 0; i < data.length; i++) {
                var item = data[i];
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
                
                // 如果有多個工作表，添加工作表名稱列
                if (showSheetColumn) {
                    var sheetCell = document.createElement('td');
                    sheetCell.textContent = item.sheetName || '';
                    row.appendChild(sheetCell);
                }
                
                processedBody.appendChild(row);
            }
            
            // 顯示結果區段
            resultSection.style.display = 'block';
        }
        
        // 解析 CSV 檔案
        function parseCSVFile(csvContent) {
            // 使用 Papa Parse 解析 CSV
            Papa.parse(csvContent, {
                header: false, // 假設沒有標題列
                dynamicTyping: true, // 自動轉換數字
                skipEmptyLines: true,
                complete: function(results) {
                    // 獲取解析後的資料
                    var jsonData = results.data;
                    
                    // 預處理數值欄位
                    jsonData = preprocessCSVData(jsonData);
                    
                    // 從解析後的 CSV 中顯示原始預覽
                    displayCSVOriginalPreview(jsonData);
                    
                    // 從解析後的 CSV 中顯示解除合併後的預覽 (對於 CSV，與原始預覽相同)
                    displayCSVUnmergedPreview(jsonData);
                    
                    // 提取資料
                    var extractedData = extractDataFromSpecificFormat(jsonData, "CSV檔案");
                    
                    // 顯示處理後資料
                    displayProcessedData(extractedData);
                },
                error: function(error) {
                    showError('解析 CSV 時發生錯誤: ' + error.message);
                }
            });
        }
        
        // 顯示 CSV 原始預覽
        function displayCSVOriginalPreview(data) {
            // 清空原始表格
            originalTable.innerHTML = '';
            
            // 建立表頭
            var thead = document.createElement('thead');
            var headerRow = document.createElement('tr');
            
            // 獲取最大列數
            var maxColumns = 0;
            for (var i = 0; i < data.length; i++) {
                if (data[i].length > maxColumns) {
                    maxColumns = data[i].length;
                }
            }
            
            // 建立表頭標籤 (A, B, C, ...)
            for (var i = 0; i < maxColumns; i++) {
                var th = document.createElement('th');
                th.textContent = String.fromCharCode(65 + i); // ASCII A-Z
                headerRow.appendChild(th);
            }
            
            thead.appendChild(headerRow);
            originalTable.appendChild(thead);
            
            // 建立表身
            var tbody = document.createElement('tbody');
            
            // 顯示所有行資料
            for (var i = 0; i < data.length; i++) {
                var tr = document.createElement('tr');
                
                for (var j = 0; j < maxColumns; j++) {
                    var td = document.createElement('td');
                    if (j < data[i].length && data[i][j] !== null) {
                        td.textContent = data[i][j];
                    }
                    tr.appendChild(td);
                }
                
                tbody.appendChild(tr);
            }
            
            originalTable.appendChild(tbody);
        }
        
        // 顯示 CSV 解除合併預覽 (與原始預覽相同，CSV 本身沒有合併儲存格)
        function displayCSVUnmergedPreview(data) {
            // 清空解除合併表格
            unmergedTable.innerHTML = '';
            
            // 建立表頭
            var thead = document.createElement('thead');
            var headerRow = document.createElement('tr');
            
            // 獲取最大列數
            var maxColumns = 0;
            for (var i = 0; i < data.length; i++) {
                if (data[i].length > maxColumns) {
                    maxColumns = data[i].length;
                }
            }
            
            // 建立表頭標籤 (A, B, C, ...)
            for (var i = 0; i < maxColumns; i++) {
                var th = document.createElement('th');
                th.textContent = String.fromCharCode(65 + i); // ASCII A-Z
                headerRow.appendChild(th);
            }
            
            thead.appendChild(headerRow);
            unmergedTable.appendChild(thead);
            
            // 建立表身
            var tbody = document.createElement('tbody');
            
            // 顯示所有行資料
            for (var i = 0; i < data.length; i++) {
                var tr = document.createElement('tr');
                
                for (var j = 0; j < maxColumns; j++) {
                    var td = document.createElement('td');
                    if (j < data[i].length && data[i][j] !== null) {
                        td.textContent = data[i][j];
                    }
                    tr.appendChild(td);
                }
                
                tbody.appendChild(tr);
            }
            
            unmergedTable.appendChild(tbody);
        }
        
        // 預處理CSV數據中的數值欄位
        function preprocessCSVData(jsonData) {
            console.log('開始預處理CSV數值欄位...');
            
            for (var i = 0; i < jsonData.length; i++) {
                var row = jsonData[i];
                if (!row) continue;
                
                // 處理第2欄 (戶數)
                if (row[1] && typeof row[1] === 'string') {
                    var numValue = parseFloat(row[1].replace(/,/g, '').trim());
                    if (!isNaN(numValue)) {
                        console.log('將CSV第2欄從文字轉為數值:', row[1], '->', numValue);
                        row[1] = numValue;
                    }
                }
                
                // 處理第4欄 (男性人口)
                if (row[3] && typeof row[3] === 'string') {
                    var numValue = parseFloat(row[3].replace(/,/g, '').trim());
                    if (!isNaN(numValue)) {
                        console.log('將CSV第4欄從文字轉為數值:', row[3], '->', numValue);
                        row[3] = numValue;
                    }
                }
                
                // 處理第5欄 (女性人口)
                if (row[4] && typeof row[4] === 'string') {
                    var numValue = parseFloat(row[4].replace(/,/g, '').trim());
                    if (!isNaN(numValue)) {
                        console.log('將CSV第5欄從文字轉為數值:', row[4], '->', numValue);
                        row[4] = numValue;
                    }
                }
            }
            
            console.log('CSV數值欄位預處理完成');
            return jsonData;
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