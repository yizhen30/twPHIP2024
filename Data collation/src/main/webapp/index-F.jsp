<%@ page import="java.io.*, java.util.*, com.opencsv.*, com.google.gson.*, ca.uhn.fhir.context.*, ca.uhn.fhir.rest.client.api.*, org.hl7.fhir.r4.model.*" %>
<%@ page contentType="text/html; charset=UTF-8" %>
<!DOCTYPE html>
<html>
<head>
    <title>CSV 轉換 FHIR JSON</title>
</head>
<body>
    <h2>上傳 CSV 轉換為 FHIR JSON</h2>
    <form method="post" enctype="multipart/form-data">
        <input type="file" name="file" accept=".csv" required>
        <button type="submit">上傳並轉換</button>
    </form>

    <%
        if ("POST".equalsIgnoreCase(request.getMethod())) {
            Part filePart = request.getPart("file");
            if (filePart != null) {
                InputStream fileContent = filePart.getInputStream();
                
                // 解析 CSV 檔案
                List<Map<String, String>> csvData = new ArrayList<>();
                try (Reader reader = new InputStreamReader(fileContent);
                     CSVReader csvReader = new CSVReader(reader)) {
                    String[] headers = csvReader.readNext(); // 讀取標題行
                    String[] row;
                    while ((row = csvReader.readNext()) != null) {
                        Map<String, String> dataMap = new HashMap<>();
                        for (int i = 0; i < headers.length; i++) {
                            dataMap.put(headers[i], row[i]);
                        }
                        csvData.add(dataMap);
                    }
                }

                // 轉換為 FHIR JSON
                FhirContext ctx = FhirContext.forR4();
                Bundle bundle = new Bundle();
                bundle.setType(Bundle.BundleType.TRANSACTION);
                
                for (Map<String, String> row : csvData) {
                    Observation obs = new Observation();
                    obs.setId(row.get("id"));  // 根據 CSV 欄位設定
                    obs.setStatus(Observation.ObservationStatus.FINAL);
                    obs.setCode(new org.hl7.fhir.r4.model.CodeableConcept().setText(row.get("category")));
                    obs.setValue(new org.hl7.fhir.r4.model.StringType(row.get("value")));

                    Bundle.BundleEntryComponent entry = new Bundle.BundleEntryComponent();
                    entry.setFullUrl("urn:uuid:" + UUID.randomUUID());
                    entry.setResource(obs);
                    entry.getRequest().setMethod(Bundle.HTTPVerb.POST).setUrl("Observation");

                    bundle.addEntry(entry);
                }

                String fhirJson = ctx.newJsonParser().setPrettyPrint(true).encodeResourceToString(bundle);

                // 上傳到 FHIR Server
                IGenericClient client = ctx.newRestfulGenericClient("http://localhost:8080/fhir");
                client.registerInterceptor(new ca.uhn.fhir.rest.client.interceptor.LoggingInterceptor(true));
                
                String uploadResult;
                try {
                    client.transaction().withBundle(bundle).execute();
                    uploadResult = "FHIR JSON 上傳成功！";
                } catch (Exception e) {
                    uploadResult = "上傳失敗：" + e.getMessage();
                }
    %>

    <h3>FHIR JSON 結果：</h3>
    <pre><%= fhirJson %></pre>
    <h3>上傳結果：</h3>
    <p><%= uploadResult %></p>

    <%
            }
        }
    %>
</body>
</html>
