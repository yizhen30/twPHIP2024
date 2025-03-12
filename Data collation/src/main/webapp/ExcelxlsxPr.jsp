<!DOCTYPE html>
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
<%@ page import="org.apache.poi.xssf.usermodel.XSSFSheet"%>
<%@ page import="org.apache.poi.xssf.usermodel.XSSFWorkbook"%>
<%@ page import="org.apache.poi.xssf.usermodel.XSSFRow"%>
<%@ page import="org.apache.poi.xssf.usermodel.XSSFCell"%>
<jsp:useBean id='objDBConfig' scope='application' class='hitstd.group.tool.database.DBConfig' />
<head>
      <title>Excel.xlsx檔</title>
</head>
<body>
<%
		
        //檔案匯入
        MultipartRequest theMultipartRequest = new MultipartRequest(request,objDBConfig.FilePath(),100*1024*1024,"UTF-8") ;		
		Enumeration theEnumeration = theMultipartRequest.getFileNames() ;
		String fileName="";
		while (theEnumeration.hasMoreElements()){
		String fieldName = (String)theEnumeration.nextElement () ;
		fileName =theMultipartRequest.getFilesystemName (fieldName);
// 		fileName=new String(fileName.getBytes("iso-8859-1"), "utf-8");
		String contentType = theMultipartRequest.getContentType (fieldName) ;
		}
		String FilePath = objDBConfig.FilePath()+fileName;
		
		FileInputStream fis = new FileInputStream(FilePath);   //obtaining bytes from the file  
		//創建新的 .xlsx 文件的工作簿實例    
		XSSFWorkbook wb = new XSSFWorkbook(fis);
        
        //創建工作表對像以檢索對象
		XSSFSheet sheet = wb.getSheet("Concepts");  //wb.getSheetAt(0);  
        
		XSSFRow row=null;
		//宣告一列
		XSSFCell cell=null;
		//宣告一個儲存格
		
%>	        			
<div >
         <form  method="post" name="form1" action="ExcelxlsxIn.jsp"> 
         <div>
             <input type="submit" class="btn-theme2 btn" value="上傳" >
         </div>                  
             <input type="text" name="fileName1" size="50" value=<%out.print(fileName);%> style="display: none">
         </form>
 </div>  		
<div>     
<table style=";margin-left:auto;margin-right:auto;"; border="1" id="tableAdd" class="table table-bordered">
   <thead>
         <tr>   
           <%
           		short r=0;
     	  		short t=0;
     	  		for (r=0;r<=0;r++)
     	  		{
     	  			row=sheet.getRow(r);
     	  			for (t=0;t<row.getLastCellNum();t++)
     	  			{
     	  		       cell=row.getCell(t);
     	  		       out.print("<th>");
     	  		    	out.print(cell);
     	  		    	out.println("</th>");
     	  			}
     	  		}
           %>
         </tr>
     </thead>
     <tbody>
 <% 
 try {
	  short i=0;
	  short y=0;
	  //以巢狀迴圈讀取所有儲存格資料
	  for (i=1;i<=sheet.getLastRowNum();i++)//sheet.getLastRowNum()
	  {
	    out.println("<tr>");
	    row=sheet.getRow(i);
	    if (row == null) {
		    row = sheet.createRow(i);
		    }
	    for (y=0;y<row.getLastCellNum();y++)
	    {
	       cell=row.getCell(y);
	       out.print("<td>");
	      
	       //判斷儲存格的格式
	       switch ( cell.getCellType() )
	       {
	           case XSSFCell.CELL_TYPE_NUMERIC:
	               out.print(cell.getNumericCellValue());
	               //getNumericCellValue()會回傳double值，若不希望出現小數點，請自行轉型為int
	               break;
	           case XSSFCell.CELL_TYPE_STRING:
	               out.print(cell.getStringCellValue());
	               break;
	           case XSSFCell.CELL_TYPE_FORMULA:
	               out.print(cell.getNumericCellValue());
	               //讀出公式儲存格計算後的值
	               //若要讀出公式內容，可用cell.getCellFormula()
	               break;
	           default:
	               out.print("");
	               break;
	       }
	       out.println("</td>");
	    }
	    out.println("</tr>");
	  }
	}catch(Exception e)  
	{  
	e.printStackTrace();  
}  
	%> 
	</tbody>
	
	</table>
	</div>
</body>
</html>
