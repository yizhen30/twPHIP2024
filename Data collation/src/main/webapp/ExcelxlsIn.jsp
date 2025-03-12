<html lang="en">
<%@ page language="java" import="java.util.*" contentType="text/html;charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.sql.*"%>
<%@ page import="javax.servlet.*,java.text.*" %>
<%@ page import="java.io.*,java.util.*"%>
<%@ page import="com.oreilly.servlet.MultipartRequest"%>
<%@ page import="java.text.SimpleDateFormat" %>
<%@ page import="java.text.DateFormat" %>
<%@ page import="org.apache.poi.ss.usermodel.Cell"%>
<%@ page import="org.apache.poi.ss.usermodel.Row"%>
<%@ page import="org.apache.poi.ss.usermodel.Sheet"%>
<%@ page import="org.apache.poi.ss.usermodel.Workbook"%>
<%@ page import="org.apache.poi.hssf.usermodel.HSSFSheet"%>
<%@ page import="org.apache.poi.hssf.usermodel.HSSFWorkbook"%>
<%@ page import="org.apache.poi.hssf.usermodel.HSSFRow"%>
<%@ page import="org.apache.poi.hssf.usermodel.HSSFCell"%>
<jsp:useBean id='objDBConfig' scope='application' class='hitstd.group.tool.database.DBConfig' />
  <head> 
      <title>ExcelXllsIn</title>
    </head>   
<body>
<%      
        //sql
        request.setCharacterEncoding("utf-8");
		Connection conn=null;
		Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
		conn = DriverManager.getConnection("jdbc:ucanaccess://" + objDBConfig.FilePath() + ";");
		Statement smt= conn.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE,ResultSet.CONCUR_READ_ONLY);
	
		String fileName = new String(request.getParameter("fileName1").getBytes("ISO-8859-1"),"UTF-8");
		String sql;
		String pathFile =objDBConfig.FilePath()+fileName;
		
		//Excel read
		FileInputStream fis = new FileInputStream(pathFile);
		HSSFWorkbook wb = new HSSFWorkbook(fis);
		HSSFSheet sheet= wb.getSheet("Concepts");
		
		try{
 		//讀取每一列的值傳到資料庫	
           	for(int r= 1;r<=sheet.getLastRowNum();r++)
        	{	
        		
        		HSSFRow row=sheet.getRow(r);
        		if (row == null) {
            	    row = sheet.createRow(r);
            	    }
        		
        		String Code = row.getCell(0).getStringCellValue();
        		String CHIDisplay = row.getCell(1).getStringCellValue();
        		String ENGDisplay = row.getCell(2).getStringCellValue();
        		String Definition = row.getCell(3).getStringCellValue();
        		
        		sql="INSERT INTO CodeSystem (Code, CHIDisplay, ENGDisplay, Definition) VALUES ('"+Code+"', '"+CHIDisplay+"', '"+ENGDisplay+"', '"+Definition+"') ";
 
        		smt.execute(sql);      
        		}
        			
        	}catch(Exception e)  
			{  
        		e.printStackTrace();  
        	}  
		
			   out.println("<script>");
			   out.println("location='ExcelImport.jsp'");
			   out.println("</script>");

 %> 


</body>

</html>

