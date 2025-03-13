<%@ page import="java.io.*, org.json.JSONArray, org.json.JSONObject" %>

<%
    response.setContentType("application/octet-stream");

    String originalFileName = (String) session.getAttribute("originalFileName");
    if (originalFileName == null) {
        response.getWriter().println("沒有可下載的檔案！");
        return;
    }
    response.setHeader("Content-Disposition", "attachment; filename=" + originalFileName);

    JSONArray processedData = new JSONArray((String) session.getAttribute("processedData"));
    PrintWriter writer = response.getWriter();

    for (int i = 0; i < processedData.length(); i++) {
        JSONArray row = processedData.getJSONArray(i);
        for (int j = 0; j < row.length(); j++) {
            writer.print(row.get(j));
            if (j < row.length() - 1) writer.print(",");
        }
        writer.println();
    }

    writer.flush();
    writer.close();
%>
