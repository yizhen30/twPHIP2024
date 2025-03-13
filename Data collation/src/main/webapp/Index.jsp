<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.sql.*"%>
<jsp:useBean id='objDBConfig' scope='application' class='hitstd.group.tool.database.DBConfig' />
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel檔案解析器</title>
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
    </style>
</head>
<body>
    <div class="container">
        <h1>鄉鎮人口統計 Excel 檔案解析器</h1>
        
        <div class="upload-section">
            <h2>1. 選擇Excel檔案</h2>
            <input type="file" id="excelFile" accept=".xls,.xlsx" />
            <button id="parseBtn">解析檔案</button>
            <div id="fileInfo"></div>
            <div class="progress-container" id="progressContainer">
                <div class="progress-bar" id="progressBar"></div>
            </div>
        </div>
        
        <div class="result-section" id="resultSection" style="display: none;">
            <h2>2. 解析結果</h2>
            <div id="resultSummary"></div>
            <button id="exportCSV" class="export-btn">匯出 CSV</button>
            <div class="table-container">
                <table id="resultTable">
                    <thead>
                        <tr>
                            <th>年分</th>
                            <th>月份</th>
                            <th>縣市別</th>
                            <th>鄉鎮別</th>
                            <th>性別</th>
                            <th>人口數</th>
                        </tr>
                    </thead>
                    <tbody id="resultBody"></tbody>
                </table>
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
            var resultSection = document.getElementById('resultSection');
            var resultSummary = document.getElementById('resultSummary');
            var resultBody = document.getElementById('resultBody');
            var exportCSVBtn = document.getElementById('exportCSV');
            
            // 存儲解析後的資料
            var parsedData = [];
            var fileName = '';
            var fileYear = '';
            var fileMonth = '';
            
            // 解析按鈕點擊事件
            parseBtn.addEventListener('click', function() {
                var file = excelFileInput.files[0];
                if (!file) {
                    alert('請先選擇Excel檔案');
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
                
                // 顯示進度條
                progressContainer.style.display = 'block';
                progressBar.style.width = '0%';
                
                // 讀取檔案
                var reader = new FileReader();
                
                reader.onprogress = function(e) {
                    if (e.lengthComputable) {
                        var percentLoaded = Math.round((e.loaded / e.total) * 100);
                        progressBar.style.width = percentLoaded + '%';
                    }
                };
                
                reader.onload = function(e) {
                    try {
                        // 開始解析
                        progressBar.style.width = '50%';
                        var data = new Uint8Array(e.target.result);
                        parseExcelFile(data);
                        
                        // 完成
                        progressBar.style.width = '100%';
                        setTimeout(function() {
                            progressContainer.style.display = 'none';
                        }, 500);
                    } catch (error) {
                        alert('解析檔案時發生錯誤: ' + error.message);
                        console.error(error);
                        progressContainer.style.display = 'none';
                    }
                };
                
                reader.onerror = function() {
                    alert('讀取檔案時發生錯誤');
                    progressContainer.style.display = 'none';
                };
                
                reader.readAsArrayBuffer(file);
            });
            
            // 匯出 CSV 按鈕事件
            exportCSVBtn.addEventListener('click', function() {
                if (parsedData.length === 0) {
                    alert('沒有資料可匯出');
                    return;
                }
                
                // 建立 CSV 內容
                var csvContent = '年分,月份,縣市別,鄉鎮別,性別,人口數\n';
                
                for (var i = 0; i < parsedData.length; i++) {
                    var row = parsedData[i];
                    csvContent += row.year + ',' + row.month + ',' + row.city + ',' + 
                                 row.district + ',"' + row.gender + '",' + row.population + '\n';
                }
                
                // 建立下載連結
                var blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
                var url = URL.createObjectURL(blob);
                var link = document.createElement('a');
                link.setAttribute('href', url);
                link.setAttribute('download', '人口統計_' + fileYear + '年' + fileMonth + '月.csv');
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
            });
            
            // 解析 Excel 函數
            function parseExcelFile(data) {
                // 使用 SheetJS 讀取 Excel
                var workbook = XLSX.read(data, { type: 'array' });
                var firstSheetName = workbook.SheetNames[0];
                var worksheet = workbook.Sheets[firstSheetName];
                
                // 轉換成 JSON
                var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                // 分析結構
                var structureResult = analyzeStructure(jsonData);
                var headers = structureResult.headers;
                var dataStartRow = structureResult.dataStartRow;
                
                // 提取資料
                var extractedData = extractData(jsonData, headers, dataStartRow);
                
                // 顯示資料
                displayData(extractedData);
            }
            
            // 分析 Excel 結構
            function analyzeStructure(jsonData) {
                // 尋找標題行
                var headerRow = -1;
                var headers = {
                    city: -1,
                    district: -1,
                    malePop: -1,
                    femalePop: -1
                };
                
                // 搜索前10行找標題行
                for (var i = 0; i < Math.min(10, jsonData.length); i++) {
                    var row = jsonData[i];
                    if (!row) continue;
                    
                    // 尋找可能的標題
                    for (var j = 0; j < row.length; j++) {
                        var cell = String(row[j] || '');
                        
                        if (cell.includes('縣') || cell.includes('市') || cell.includes('縣市')) {
                            headers.city = j;
                        }
                        if (cell.includes('鄉') || cell.includes('鎮') || cell.includes('區') || 
                            cell.includes('鄉鎮') || cell.includes('鄉鎮市區')) {
                            headers.district = j;
                        }
                        if ((cell.includes('男') && cell.includes('口')) || 
                            (cell.includes('男性') && cell.includes('人口'))) {
                            headers.malePop = j;
                        }
                        if ((cell.includes('女') && cell.includes('口')) || 
                            (cell.includes('女性') && cell.includes('人口'))) {
                            headers.femalePop = j;
                        }
                    }
                    
                    // 如果找到主要欄位就設定標題行
                    if (headers.city >= 0 || headers.district >= 0) {
                        headerRow = i;
                        break;
                    }
                }
                
                // 找資料開始行 (通常在標題行之後)
                var dataStartRow = headerRow + 1;
                
                // 如果未找到標題，假設第一行是標題
                if (headerRow === -1) {
                    headerRow = 0;
                    dataStartRow = 1;
                }
                
                // 再次嘗試找欄位位置
                if (headers.city === -1 || headers.district === -1 || 
                    headers.malePop === -1 || headers.femalePop === -1) {
                    // 查看標題行和下一行，有時候欄位名稱分成多行
                    var headerRowData = jsonData[headerRow] || [];
                    var nextRowData = jsonData[headerRow + 1] || [];
                    
                    for (var j = 0; j < Math.max(headerRowData.length, nextRowData.length); j++) {
                        var headerCell = String(headerRowData[j] || '');
                        var nextCell = String(nextRowData[j] || '');
                        var combinedHeader = headerCell + ' ' + nextCell;
                        
                        if (headers.city === -1 && 
                            (combinedHeader.includes('縣') || combinedHeader.includes('市') || 
                             combinedHeader.includes('縣市'))) {
                            headers.city = j;
                        }
                        if (headers.district === -1 && 
                            (combinedHeader.includes('鄉') || combinedHeader.includes('鎮') || 
                             combinedHeader.includes('區') || combinedHeader.includes('鄉鎮'))) {
                            headers.district = j;
                        }
                        if (headers.malePop === -1 && 
                            ((combinedHeader.includes('男') && combinedHeader.includes('口')) || 
                             (combinedHeader.includes('男性') && combinedHeader.includes('人口')))) {
                            headers.malePop = j;
                        }
                        if (headers.femalePop === -1 && 
                            ((combinedHeader.includes('女') && combinedHeader.includes('口')) || 
                             (combinedHeader.includes('女性') && combinedHeader.includes('人口')))) {
                            headers.femalePop = j;
                        }
                    }
                    
                    // 如果標題在兩行，資料開始行可能要再往下一行
                    if (headers.malePop === -1 || headers.femalePop === -1) {
                        dataStartRow = headerRow + 2;
                    }
                }
                
                console.log('找到的欄位位置:', headers);
                console.log('資料開始行:', dataStartRow);
                
                return { headers: headers, dataStartRow: dataStartRow };
            }
            
            // 提取資料
            function extractData(jsonData, headers, dataStartRow) {
                var result = [];
                var currentCity = '';
                
                // 從資料開始行開始處理
                for (var i = dataStartRow; i < jsonData.length; i++) {
                    var row = jsonData[i];
                    if (!row || row.length === 0) continue;
                    
                    // 獲取欄位值
                    var cityValue = row[headers.city];
                    var districtValue = headers.district >= 0 ? row[headers.district] : null;
                    var malePopValue = headers.malePop >= 0 ? row[headers.malePop] : null;
                    var femalePopValue = headers.femalePop >= 0 ? row[headers.femalePop] : null;
                    
                    // 檢查是否有縣市值
                    if (cityValue && typeof cityValue === 'string' && 
                        (cityValue.includes('縣') || cityValue.includes('市'))) {
                        currentCity = cityValue.trim();
                        
                        // 如果這是縣市總計行，加入資料
                        if (malePopValue !== null && femalePopValue !== null) {
                            // 男性資料
                            result.push({
                                year: fileYear,
                                month: fileMonth,
                                city: currentCity,
                                district: '總計',
                                gender: '男',
                                population: malePopValue
                            });
                            
                            // 女性資料
                            result.push({
                                year: fileYear,
                                month: fileMonth,
                                city: currentCity,
                                district: '總計',
                                gender: '女',
                                population: femalePopValue
                            });
                        }
                    }
                    // 檢查是否有鄉鎮值
                    else if (currentCity && districtValue) {
                        // 男性資料
                        if (malePopValue !== null) {
                            result.push({
                                year: fileYear,
                                month: fileMonth,
                                city: currentCity,
                                district: districtValue.trim(),
                                gender: '男',
                                population: malePopValue
                            });
                        }
                        
                        // 女性資料
                        if (femalePopValue !== null) {
                            result.push({
                                year: fileYear,
                                month: fileMonth,
                                city: currentCity,
                                district: districtValue.trim(),
                                gender: '女',
                                population: femalePopValue
                            });
                        }
                    }
                }
                
                return result;
            }
            
            // 顯示資料
            function displayData(data) {
                parsedData = data;
                
                // 清空表格
                resultBody.innerHTML = '';
                
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
                    
                    var populationCell = document.createElement('td');
                    populationCell.textContent = item.population;
                    
                    row.appendChild(yearCell);
                    row.appendChild(monthCell);
                    row.appendChild(cityCell);
                    row.appendChild(districtCell);
                    row.appendChild(genderCell);
                    row.appendChild(populationCell);
                    
                    resultBody.appendChild(row);
                }
                
                // 顯示結果區段
                resultSection.style.display = 'block';
                
                // 如果有 100 筆以上記錄，添加注意文字
                if (data.length > 100) {
                    resultSummary.textContent += ' (只顯示前 100 筆)';
                }
            }
        });
    </script>
</body>
</html>