<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.sql.*"%>
<jsp:useBean id='objDBConfig' scope='application' class='hitstd.group.tool.database.DBConfig' />
<!DOCTYPE html>
<html lang="zh-TW">
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
        h1, h2, h3 {
            color: #2c3e50;
        }
        .card {
            background-color: #fff;
            border-radius: 5px;
            padding: 20px;
            margin-bottom: 20px;
            box-shadow: 0 2px 5px rgba(0, 0, 0, 0.1);
        }
        .input-group {
            margin-bottom: 15px;
        }
        label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
        }
        input[type="file"] {
            display: block;
            width: 100%;
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 4px;
            margin-bottom: 10px;
        }
        button {
            background-color: #3498db;
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 16px;
        }
        button:hover {
            background-color: #2980b9;
        }
        button:disabled {
            background-color: #95a5a6;
            cursor: not-allowed;
        }
        .progress-container {
            height: 20px;
            background-color: #ecf0f1;
            border-radius: 10px;
            margin: 15px 0;
            overflow: hidden;
        }
        .progress-bar {
            height: 100%;
            background-color: #2ecc71;
            border-radius: 10px;
            width: 0;
            transition: width 0.3s ease;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 15px 0;
            font-size: 14px;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
            font-weight: bold;
        }
        tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        .result-summary {
            background-color: #edf7ff;
            padding: 10px 15px;
            border-radius: 4px;
            margin-bottom: 15px;
            font-weight: bold;
        }
        .action-buttons {
            display: flex;
            gap: 10px;
            margin: 15px 0;
        }
        .error {
            color: #e74c3c;
            background-color: #fadbd8;
            padding: 10px;
            border-radius: 4px;
            margin: 10px 0;
        }
        .success {
            color: #27ae60;
            background-color: #d4efdf;
            padding: 10px;
            border-radius: 4px;
            margin: 10px 0;
        }
        #loadingText {
            text-align: center;
            margin: 5px 0;
            font-style: italic;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>鄉鎮人口統計Excel解析器</h1>
        
        <div class="card">
            <h2>選擇Excel檔案</h2>
            <div class="input-group">
                <label for="excelFile">選擇包含鄉鎮人口統計資料的Excel檔案 (.xls 或 .xlsx)</label>
                <input type="file" id="excelFile" accept=".xls,.xlsx">
            </div>
            <button id="parseBtn">解析檔案</button>
            
            <div id="progressContainer" class="progress-container" style="display: none;">
                <div id="progressBar" class="progress-bar"></div>
            </div>
            <div id="loadingText" style="display: none;">正在解析檔案，請稍候...</div>
            <div id="errorMessage" class="error" style="display: none;"></div>
        </div>
        
        <div id="resultContainer" class="card" style="display: none;">
            <h2>解析結果</h2>
            <div id="fileInfo" class="result-summary"></div>
            
            <h3>資料預覽</h3>
            <div style="overflow-x: auto;">
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
            
            <div class="action-buttons">
                <button id="exportBtn">匯出CSV</button>
                <button id="dbBtn">準備資料庫格式</button>
            </div>
            
            <div id="successMessage" class="success" style="display: none;"></div>
        </div>
        
        <div id="dbFormatContainer" class="card" style="display: none;">
            <h2>資料庫匯入準備</h2>
            <p>下列資料已經整理成適合資料庫匯入的格式：</p>
            
            <h3>1. 縣市資料表</h3>
            <div style="overflow-x: auto;">
                <table id="cityTable">
                    <thead>
                        <tr>
                            <th>CityCode</th>
                            <th>CityName</th>
                        </tr>
                    </thead>
                    <tbody id="cityBody"></tbody>
                </table>
            </div>
            
            <h3>2. 鄉鎮資料表</h3>
            <div style="overflow-x: auto;">
                <table id="districtTable">
                    <thead>
                        <tr>
                            <th>DistrictCode</th>
                            <th>CityCode</th>
                            <th>DistrictName</th>
                        </tr>
                    </thead>
                    <tbody id="districtBody"></tbody>
                </table>
            </div>
            
            <h3>3. 人口統計資料</h3>
            <div style="overflow-x: auto;">
                <table id="populationTable">
                    <thead>
                        <tr>
                            <th>YearID (民國年)</th>
                            <th>MonthID</th>
                            <th>CityCode</th>
                            <th>DistrictCode</th>
                            <th>Gender</th>
                            <th>Population</th>
                        </tr>
                    </thead>
                    <tbody id="populationBody"></tbody>
                </table>
            </div>
            
            <div class="action-buttons">
                <button id="exportSqlBtn">匯出SQL指令</button>
                <button id="backBtn">返回資料預覽</button>
            </div>
        </div>
    </div>

    <script>
        // DOM元素
        const excelFileInput = document.getElementById('excelFile');
        const parseBtn = document.getElementById('parseBtn');
        const progressContainer = document.getElementById('progressContainer');
        const progressBar = document.getElementById('progressBar');
        const loadingText = document.getElementById('loadingText');
        const errorMessage = document.getElementById('errorMessage');
        const resultContainer = document.getElementById('resultContainer');
        const fileInfo = document.getElementById('fileInfo');
        const resultBody = document.getElementById('resultBody');
        const exportBtn = document.getElementById('exportBtn');
        const dbBtn = document.getElementById('dbBtn');
        const successMessage = document.getElementById('successMessage');
        const dbFormatContainer = document.getElementById('dbFormatContainer');
        const cityBody = document.getElementById('cityBody');
        const districtBody = document.getElementById('districtBody');
        const populationBody = document.getElementById('populationBody');
        const exportSqlBtn = document.getElementById('exportSqlBtn');
        const backBtn = document.getElementById('backBtn');

        // 存儲解析後的資料
        let parsedData = [];
        let fileName = '';
        let fileYear = '';
        let fileMonth = '';
        let cityData = [];
        let districtData = [];
        
        // 設定按鈕事件
        parseBtn.addEventListener('click', parseExcelFile);
        exportBtn.addEventListener('click', exportToCsv);
        dbBtn.addEventListener('click', showDbFormat);
        backBtn.addEventListener('click', showResultView);
        exportSqlBtn.addEventListener('click', exportSqlCommands);
        
        // 解析Excel文件
        function parseExcelFile() {
            const file = excelFileInput.files[0];
            if (!file) {
                showError('請選擇Excel檔案');
                return;
            }
            
            // 檢查檔案類型
            const fileExt = file.name.split('.').pop().toLowerCase();
            if (fileExt !== 'xls' && fileExt !== 'xlsx') {
                showError('請選擇正確的Excel檔案 (.xls 或 .xlsx)');
                return;
            }
            
            // 開始解析
            fileName = file.name;
            showProgressBar();
            hideError();
            
            const reader = new FileReader();
            
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    // 更新進度
                    updateProgress(50);
                    
                    // 解析第一個工作表
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    
                    // 解析標題和年月
                    const docInfo = parseDocumentInfo(worksheet);
                    fileYear = docInfo.year;
                    fileMonth = docInfo.month;
                    
                    // 解析資料
                    parsedData = parseWorksheet(worksheet, docInfo);
                    
                    // 完成
                    updateProgress(100);
                    
                    // 顯示結果
                    displayResults(parsedData, docInfo);
                    
                } catch (error) {
                    console.error('解析Excel時發生錯誤:', error);
                    showError('解析檔案時發生錯誤: ' + error.message);
                    hideProgressBar();
                }
            };
            
            reader.onerror = function() {
                showError('讀取檔案時發生錯誤');
                hideProgressBar();
            };
            
            reader.readAsArrayBuffer(file);
        }
        
        // 解析文件標題及年月資訊
        function parseDocumentInfo(worksheet) {
            const range = XLSX.utils.decode_range(worksheet['!ref']);
            let title = '';
            let year = '';
            let month = '';
            
            // 檢查第一行可能的標題
            for (let c = range.s.c; c <= range.e.c; c++) {
                const cellAddress = XLSX.utils.encode_cell({ r: 0, c });
                const cell = worksheet[cellAddress];
                if (cell && cell.v) {
                    title = cell.v.toString().trim();
                    
                    // 嘗試從標題中提取年月
                    const yearMonthMatch = title.match(/(\d+)年(\d+)月/);
                    if (yearMonthMatch) {
                        year = yearMonthMatch[1];  // 民國年
                        month = yearMonthMatch[2];
                    }
                    
                    break;
                }
            }
            
            // 如果標題中沒有找到年月，嘗試從文件名中提取
            if (!year || !month) {
                const fileNameMatch = fileName.match(/(\d+)年(\d+)月/);
                if (fileNameMatch) {
                    year = fileNameMatch[1];
                    month = fileNameMatch[2];
                }
            }
            
            return {
                title: title,
                year: year,
                month: month
            };
        }
        
        // 解析工作表
        function parseWorksheet(worksheet, docInfo) {
            const result = [];
            const range = XLSX.utils.decode_range(worksheet['!ref']);
            let currentCity = '';
            
            // 識別資料欄位位置
            const columns = findColumns(worksheet, range);
            
            // 從第4行開始處理資料 (基於截圖中的資料結構)
            for (let r = 3; r <= range.e.r; r++) {
                // 讀取區域別（縣市或鄉鎮）
                const areaCellAddress = XLSX.utils.encode_cell({ r, c: columns.area });
                const areaCell = worksheet[areaCellAddress];
                
                if (!areaCell) continue;
                
                const areaValue = areaCell.v.toString().trim();
                
                // 檢查是否為縣市名稱
                if (areaValue.includes('市') || areaValue.includes('縣')) {
                    currentCity = areaValue;
                    
                    // 讀取縣市層級的人口數
                    const totalCellAddress = XLSX.utils.encode_cell({ r, c: columns.total });
                    const maleCellAddress = XLSX.utils.encode_cell({ r, c: columns.male });
                    const femaleCellAddress = XLSX.utils.encode_cell({ r, c: columns.female });
                    
                    const totalCell = worksheet[totalCellAddress];
                    const maleCell = worksheet[maleCellAddress];
                    const femaleCell = worksheet[femaleCellAddress];
                    
                    // 只有在人口數存在時才添加
                    if (maleCell && femaleCell) {
                        // 添加男性資料
                        result.push({
                            year: docInfo.year,
                            month: docInfo.month,
                            city: currentCity,
                            district: '總計',
                            gender: '男',
                            population: maleCell.v
                        });
                        
                        // 添加女性資料
                        result.push({
                            year: docInfo.year,
                            month: docInfo.month,
                            city: currentCity,
                            district: '總計',
                            gender: '女',
                            population: femaleCell.v
                        });
                    }
                }
                // 處理鄉鎮層級
                else if (currentCity && areaValue.includes('區')) {
                    // 讀取鄉鎮層級的人口數
                    const totalCellAddress = XLSX.utils.encode_cell({ r, c: columns.total });
                    const maleCellAddress = XLSX.utils.encode_cell({ r, c: columns.male });
                    const femaleCellAddress = XLSX.utils.encode_cell({ r, c: columns.female });
                    
                    const totalCell = worksheet[totalCellAddress];
                    const maleCell = worksheet[maleCellAddress];
                    const femaleCell = worksheet[femaleCellAddress];
                    
                    // 只有在人口數存在時才添加
                    if (maleCell && femaleCell) {
                        // 添加男性資料
                        result.push({
                            year: docInfo.year,
                            month: docInfo.month,
                            city: currentCity,
                            district: areaValue,
                            gender: '男',
                            population: maleCell.v
                        });
                        
                        // 添加女性資料
                        result.push({
                            year: docInfo.year,
                            month: docInfo.month,
                            city: currentCity,
                            district: areaValue,
                            gender: '女',
                            population: femaleCell.v
                        });
                    }
                }
            }
            
            return result;
        }
        
        // 查找各欄位位置
        function findColumns(worksheet, range) {
            let area = 0;  // 預設區域別在A欄
            let total = -1;
            let male = -1;
            let female = -1;
            
            // 搜索第3行標題
            for (let c = range.s.c; c <= range.e.c; c++) {
                const cellAddress = XLSX.utils.encode_cell({ r: 2, c });
                const cell = worksheet[cellAddress];
                
                if (cell && cell.v) {
                    const value = cell.v.toString().trim();
                    if (value === '計') {
                        total = c;
                    } else if (value === '男') {
                        male = c;
                    } else if (value === '女') {
                        female = c;
                    }
                }
            }
            
            return {
                area: area,
                total: total,
                male: male,
                female: female
            };
        }
        
        // 顯示解析結果
        function displayResults(data, docInfo) {
            // 更新檔案資訊
            fileInfo.innerHTML = `
                <div>檔案名稱: ${fileName}</div>
                <div>資料時間: 民國 ${docInfo.year} 年 ${docInfo.month} 月</div>
                <div>資料總筆數: ${data.length} 筆</div>
            `;
            
            // 清空結果表格
            resultBody.innerHTML = '';
            
            // 顯示前50筆資料
            const previewData = data.slice(0, 50);
            for (let i = 0; i < previewData.length; i++) {
                const item = previewData[i];
                const row = document.createElement('tr');
                
                row.innerHTML = `
                    <td>${item.year}</td>
                    <td>${item.month}</td>
                    <td>${item.city}</td>
                    <td>${item.district}</td>
                    <td>${item.gender}</td>
                    <td>${item.population}</td>
                `;
                
                resultBody.appendChild(row);
            }
            
            // 顯示結果容器
            hideProgressBar();
            resultContainer.style.display = 'block';
            
            // 如果資料超過50筆，顯示提示
            if (data.length > 50) {
                showSuccess(`僅顯示前50筆資料，實際共有 ${data.length} 筆資料`);
            } else {
                hideSuccess();
            }
        }
        
        // 匯出CSV文件
        function exportToCsv() {
            if (parsedData.length === 0) {
                showError('沒有資料可匯出');
                return;
            }
            
            // 準備CSV內容
            let csvContent = '年分,月份,縣市別,鄉鎮別,性別,人口數\n';
            
            parsedData.forEach(function(item) {
                csvContent += `${item.year},${item.month},${item.city},${item.district},"${item.gender}",${item.population}\n`;
            });
            
            // 建立下載連結
            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.setAttribute('href', url);
            link.setAttribute('download', `人口統計_${fileYear}年${fileMonth}月.csv`);
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            
            showSuccess('CSV檔案已成功匯出');
        }
        
        // 顯示資料庫格式
        function showDbFormat() {
            if (parsedData.length === 0) {
                showError('沒有資料可處理');
                return;
            }
            
            // 準備資料庫資料
            prepareDbData();
            
            // 填充資料表
            fillDbTables();
            
            // 隱藏結果視圖，顯示資料庫格式視圖
            resultContainer.style.display = 'none';
            dbFormatContainer.style.display = 'block';
        }
        
        // 返回資料預覽視圖
        function showResultView() {
            dbFormatContainer.style.display = 'none';
            resultContainer.style.display = 'block';
        }
        
        // 準備資料庫格式資料
        function prepareDbData() {
            // 整理縣市資料
            cityData = [];
            const cityMap = new Map();
            
            parsedData.forEach(function(item) {
                if (!cityMap.has(item.city)) {
                    // 簡單的城市代碼生成，實際應用請使用標準代碼
                    const cityCode = generateCityCode(item.city);
                    cityMap.set(item.city, cityCode);
                    
                    cityData.push({
                        cityCode: cityCode,
                        cityName: item.city
                    });
                }
            });
            
            // 整理鄉鎮資料
            districtData = [];
            const districtMap = new Map();
            
            parsedData.forEach(function(item) {
                // 跳過總計行
                if (item.district === '總計') return;
                
                const key = item.city + '-' + item.district;
                if (!districtMap.has(key)) {
                    const cityCode = cityMap.get(item.city);
                    const districtCode = generateDistrictCode(cityCode, item.district);
                    districtMap.set(key, districtCode);
                    
                    districtData.push({
                        districtCode: districtCode,
                        cityCode: cityCode,
                        districtName: item.district
                    });
                }
            });
        }
        
        // 生成城市代碼
        function generateCityCode(cityName) {
            // 在實際應用中，應使用標準的城市代碼
            // 這裡簡單實現，使用拼音首字母或其他邏輯
            const codeMap = {
                '臺北市': 'A',
                '臺中市': 'B',
                '基隆市': 'C',
                '臺南市': 'D',
                '高雄市': 'E',
                '新北市': 'F',
                '宜蘭縣': 'G',
                '桃園市': 'H',
                '嘉義市': 'I',
                '新竹縣': 'J',
                '苗栗縣': 'K',
                '南投縣': 'M',
                '彰化縣': 'N',
                '新竹市': 'O',
                '雲林縣': 'P',
                '嘉義縣': 'Q',
                '屏東縣': 'T',
                '花蓮縣': 'U',
                '臺東縣': 'V',
                '金門縣': 'W',
                '澎湖縣': 'X',
                '連江縣': 'Z'
            };
            
            return codeMap[cityName] || cityName.charCodeAt(0).toString(16).toUpperCase();
        }
        
        // 生成鄉鎮代碼
        function generateDistrictCode(cityCode, districtName) {
            // 在實際應用中，應使用標準的鄉鎮代碼
            // 這裡簡單實現，使用城市代碼加序號
            return cityCode + districtName.slice(0, 2);
        }
        
        // 填充資料庫表格
        function fillDbTables() {
            // 清空表格
            cityBody.innerHTML = '';
            districtBody.innerHTML = '';
            populationBody.innerHTML = '';
            
            // 填充縣市表格
            cityData.forEach(function(city) {
                const row = document.createElement('tr');
                row.innerHTML = `
                    <td>${city.cityCode}</td>
                    <td>${city.cityName}</td>
                `;
                cityBody.appendChild(row);
            });
            
            // 填充鄉鎮表格
            districtData.forEach(function(district) {
                const row = document.createElement('tr');
                row.innerHTML = `
                    <td>${district.districtCode}</td>
                    <td>${district.cityCode}</td>
                    <td>${district.districtName}</td>
                `;
                districtBody.appendChild(row);
            });
            
            // 填充人口統計表格（僅顯示前20筆）
            const cityCodeMap = new Map();
            cityData.forEach(city => cityCodeMap.set(city.cityName, city.cityCode));
            
            const districtCodeMap = new Map();
            districtData.forEach(district => {
                districtCodeMap.set(district.cityCode + '-' + district.districtName, district.districtCode);
            });
            
            // 只顯示前20筆做預覽
            const previewPopulation = parsedData.slice(0, 20);
            
            previewPopulation.forEach(function(item) {
                const row = document.createElement('tr');
                const cityCode = cityCodeMap.get(item.city);
                let districtCode = null;
                
                if (item.district !== '總計') {
                    districtCode = districtCodeMap.get(cityCode + '-' + item.district);
                }
                
                row.innerHTML = `
                    <td>${item.year}</td>
                    <td>${item.month}</td>
                    <td>${cityCode}</td>
                    <td>${districtCode || 'NULL'}</td>
                    <td>${item.gender}</td>
                    <td>${item.population}</td>
                `;
                populationBody.appendChild(row);
            });
        }
        
        // 匯出SQL指令
        function exportSqlCommands() {
            if (parsedData.length === 0 || cityData.length === 0) {
                showError('沒有資料可匯出');
                return;
            }
            
            // 準備SQL指令
            let sqlContent = '';
            
            // 年份資料
            sqlContent += "-- 年份資料\n";
            sqlContent += `INSERT INTO Year (Year, ROCYear) VALUES (${parseInt(fileYear) + 1911}, ${fileYear});\n\n`;
            
            // 月份資料
            sqlContent += "-- 月份資料\n";
            sqlContent += `INSERT INTO Month (Month, MonthName) VALUES (${fileMonth}, '${getMonthName(fileMonth)}');\n\n`;
            
            // 縣市資料
            sqlContent += "-- 縣市資料\n";
            cityData.forEach(function(city) {
                sqlContent += `INSERT INTO City (CityCode, CityName) VALUES ('${city.cityCode}', '${city.cityName}');\n`;
            });
            sqlContent += "\n";
            
            // 鄉鎮資料
            sqlContent += "-- 鄉鎮資料\n";
            districtData.forEach(function(district) {
                sqlContent += `INSERT INTO District (DistrictCode, CityCode, DistrictName) VALUES ('${district.districtCode}', '${district.cityCode}', '${district.districtName}');\n`;
            });
            sqlContent += "\n";
            
            // 人口統計資料（僅顯示前20筆做預覽）
            sqlContent += "-- 人口統計資料 (僅顯示前20筆示例)\n";
            
            const cityCodeMap = new Map();
            cityData.forEach(city => cityCodeMap.set(city.cityName, city.cityCode));
            
            const districtCodeMap = new Map();
            districtData.forEach(district => {
                district