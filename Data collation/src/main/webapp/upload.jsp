<%@ page import="java.io.*, java.util.*, org.apache.poi.ss.usermodel.*, org.apache.poi.xssf.usermodel.XSSFWorkbook, org.apache.poi.hssf.usermodel.HSSFWorkbook" %>
<%@ page import="org.apache.commons.fileupload.servlet.ServletFileUpload, org.apache.commons.fileupload.FileItem, org.apache.commons.fileupload.disk.DiskFileItemFactory" %>
<%@ page import="org.json.JSONArray, org.json.JSONObject" %>

<%
    response.setContentType("application/json");
    response.setCharacterEncoding("UTF-8");
    PrintWriter out = response.getWriter();
    JSONArray resultArray = new JSONArray();

    try {
        if (!ServletFileUpload.isMultipartContent(request)) {
            out.println("{\"error\":\"請上傳檔案！\"}");
            return;
        }

        ServletFileUpload upload = new ServletFileUpload(new DiskFileItemFactory());
        List<FileItem> items = upload.parseRequest(request);

        for (FileItem item : items) {
            if (!item.isFormField()) {
                InputStream fileContent = item.getInputStream();
                String fileName = item.getName();
                String extension = fileName.substring(fileName.lastIndexOf(".") + 1).toLowerCase();

                if (extension.equals("xls") || extension.equals("xlsx")) {
                    Workbook workbook = (extension.equals("xls")) ? new HSSFWorkbook(fileContent) : new XSSFWorkbook(fileContent);
                    Sheet sheet = workbook.getSheetAt(0);

                    for (Row row : sheet) {
                        JSONArray rowArray = new JSONArray();
                        for (Cell cell : row) {
                            switch (cell.getCellType()) {
                                case STRING:
                                    rowArray.put(cell.getStringCellValue().trim());
                                    break;
                                case NUMERIC:
                                    rowArray.put(cell.getNumericCellValue());
                                    break;
                                default:
                                    rowArray.put("");
                            }
                        }
                        resultArray.put(rowArray);
                    }
                    workbook.close();
                }
                fileContent.close();
            }
        }

        out.println(resultArray.toString());
    } catch (Exception e) {
        out.println("{\"error\":\"發生錯誤：" + e.getMessage() + "\"}");
        e.printStackTrace();
    } finally {
        out.flush();
        out.close();
    }
%>
