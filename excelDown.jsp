<%@page contentType= "text/html;charset = UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.net.*,java.io.*,java.sql.*,java.util.*,javax.sql.*,javax.naming.*,javax.sql.DataSource" %>
<%@ page import="org.apache.poi.poifs.filesystem.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.POIFSFileSystem.*" %>
<%@ page import="org.apache.poi.poifs.dev.*" %>
<%@ page import="org.apache.poi.xssf.*" %>
<%@ page import="org.apache.poi.xssf.model.*" %>
<%@ page import="org.apache.poi.xssf.usermodel.*" %>
<%@ page import="org.apache.poi.xssf.util.*" %>
<%@ page import="org.apache.poi.ss.usermodel.*" %>
<%!
  public static String kor2Db(String sToDb) throws UnsupportedEncodingException{
        if (sToDb == null) return null;
        else return new String(sToDb.getBytes("8859_1"), "KSC5601");
     }
 public static String dbTxt2kor(String sFromDb) throws UnsupportedEncodingException{
        if (sFromDb == null) return null;
        else return new String(sFromDb.getBytes("KSC5601"), "8859_1");
     }
 %>
 <%!
public static String fileName(String name) {
	String filename="";
	switch (name) {
	case "borrow_book":
		filename="borrow_book.xlsx";
		break;
	case "use_avm":
		filename="use_avm.xlsx";
		break;
	case "entry_lib":
		filename="entry_lib.xlsx";
		break;
	case "ticket_seat":
		filename="ticket_seat.xlsx";
		break;
	case "edu_user":
		filename="edu_user.xlsx";
		break;
	case "review_book":
		filename="review_book.xlsx";
		break;
	case "e_learning":
		filename="e_learning.xlsx";
		break;
	case "part_lib":
		filename="part_lib.xlsx";
		break;
	case "e_book":
		filename="e_book.xlsx";
		break;
	default:
		filename="total.xlsx";
	}
	return filename;
}
 
 public static boolean isNumber(String str) {

     try {
         //Double.parseDouble(str);
         //학번 제외
         if(str.length()==9 && str.startsWith("2")) {
        	 return false;
         } else {
         Integer.parseInt(str);
         return true;
         }
     } catch (Exception e) {
         return false;
     }
 }
%>
<%@ include file="./try.jsp" %>
<%
request.setCharacterEncoding("UTF-8");
response.reset(); 
String selName = request.getParameter("selName");

String listSql="select stuNum, num, score from "+ selName;

if(selName.equals("total"))
{
	listSql="select stuNum, bb, ua, enl, ts, eu, rb, el, pl, eb from " + selName;
}
Workbook wb = null;
Sheet sheet = null;
Row row = null;
Cell cell = null;
int rsCnt=1;


//String filepath="C:/tmj/";
String filename = fileName(selName);

            //쿼크북 객체 생성
            wb = new XSSFWorkbook();
			//새로운 시크생성
            sheet = wb.createSheet("new sheet");
			/*셀 서식 */
            CellStyle style = wb.createCellStyle();
       		
            
    

            rs = stmt.executeQuery(listSql);
			//메타데이터생성(테이블의 컬럼명,속성등의 정보를 사지고 온다.
            ResultSetMetaData rsmd =  rs.getMetaData();

			//컬럼수 카운트
            int numberOfColumns = rsmd.getColumnCount();

			//엑셀파일의 타이들을 만들기 위한 배열값
			String[] tbl_title   = {"학번","횟수","점수"};
			String[] total_title = {"학번","대출","시청각 자료 이용","도서관 출입", "열람실 좌석발권", "도서관 특강", "서평쓰기", "이러닝강좌 수강 실적", "도서관 행사 참여", "전자책 우수 이용"};

			//컬럼의 값을 받을 배열변수선언
			String[] ColumnsName=new String[numberOfColumns];  

			//컬럼의 타잎정보(즉,String , Integer, Decimal등 자료형을 정보
			int[] ColumnsType = new int[numberOfColumns];     

			row = sheet.createRow((short)0);

            for(int i=0 ; i<numberOfColumns;i++){
			   ColumnsType[i]=rsmd.getColumnType(i+1);
			   cell = row.createCell((short)(i));
			   if(selName.equals("total")) {
				   cell.setCellValue(total_title[i]);
			   } else {
			   cell.setCellValue(tbl_title[i]);
			   }


            }
           

		   //데이터베이스의 내용을 출력
            while(rs.next()){
                    row = sheet.createRow((short)rsCnt);

                    
                    for(int k=0 ; k<numberOfColumns;k++){
                            cell = row.createCell((short)k);
                            String cellValue= kor2Db(rs.getString(k+1));
                            cell.setCellValue(kor2Db(rs.getString(k+1)));
                            if(isNumber(cellValue)) {
                            	int i = Integer.parseInt(cellValue);
                            	cell.setCellValue(i);
                            } else {
                            	cell.setCellValue(cellValue);
                            }
                    }
                
                    rsCnt++;
            }
         //파일스트림으로 파일을 읽어 들인다.   
        FileOutputStream fileOut  = new FileOutputStream(filename);
       // FileOutputStream fileOut  = new FileOutputStream(filepath+filename);
        wb.write(fileOut);
        fileOut.close();
        
 //여기부터 화일 다운로드 창이 자동으로 뜨게 하기 위한 코딩(임시화일을 스트림으로 저장)
 //해당 경로의 파일 객체를 만든다. 
// File file = new File (filepath+filename);
 File file = new File (filename);
 
 //파일 스트림을 저장하기 위한 바이트 배열 생성. 
 byte[] bytestream = new byte[(int)file.length()]; 
 
//파일 객체를 스트림으로 불러온다. 
 FileInputStream filestream = new FileInputStream(file);

 //파일 스트림을 바이트 배열에 넣는다. 
 int i = 0;
 int j = 0;   
 while((i = filestream.read()) != -1) { 
  bytestream[j] = (byte)i; 
  j++; 
 }
 filestream.close();   //FileInputStream을 닫아줘야 file이 삭제된다.

 try{
  //화일을 생성과 동시에 byte[]배열에 입력후 화일은 삭제
  boolean  success = file.delete();
  if(!success) System.out.println("<script>alert('not success')</script>"); 
 } catch(IllegalArgumentException e){ 
  System.err.println(e.getMessage()); 
 } 

//Content-Disposition 헤더에 파일 이름 세팅.
 response.setHeader("Content-Disposition","attachment; filename="+filename); 

// 응답 스트림 객체를 생성한다. 
 OutputStream outStream = response.getOutputStream();  

// 응답 스트림에 파일 바이트 배열을 쓴다. 
 outStream.write(bytestream);  
 outStream.close();
%>
<%@ include file="./finally.jsp" %>
