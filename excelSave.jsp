<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.net.*,java.io.*,java.sql.*,java.util.*,javax.sql.*,javax.naming.*,javax.sql.DataSource, org.apache.poi.hssf.usermodel.*" %>
<%@ page import="com.oreilly.servlet.MultipartRequest"%>
<%@ page import="com.oreilly.servlet.multipart.DefaultFileRenamePolicy"%>
<%@ page import="org.apache.poi.xssf.usermodel.*" %>
<%!
 //null체크를 위한 function으로 nullpoint에러를 잡는다.
 public static String nullcheck(String str) throws Exception {

        if (str == null){
            return "";
        }else{
     	   	return str;     // 넘어온 Stirng 을 그대로 다시 return
	     }
     }
public static String dbTxt2kor(String sFromDb) throws UnsupportedEncodingException{
    if (sFromDb == null) return null;
    else return new String(sFromDb.getBytes("KSC5601"), "8859_1");
 }

public static String kor2Db(String sToDb) throws UnsupportedEncodingException{
	if(sToDb == null) return null;
	else return new String(sToDb.getBytes("8859_1"),"KSC5601");
  }

public static String type_check(String typeName) {
	String tbname = "";
	switch(typeName) {
	case "대출":
		tbname="borrow_book";
		break;
	case "시청각 자료 이용":
		tbname="use_avm";
		break;
	case "도서관 출입":
		tbname="entry_lib";
		break;
	case "열람실 좌석발권":
		tbname="ticket_seat";
		break;
	case "도서관 교육":
		tbname="edu_user";
		break;
	case "서평쓰기":
		tbname="review_book";
		break;
	case "이러닝강좌 수강 실적":
		tbname="e_learning";
		break;
	case "도서관 행사 참여":
		tbname="part_lib";
		break;
	case "전자책 우수 이용":
		tbname="e_book";
		break;
	case "total":
		tbname="total";
		break;
	default:
	}
	return tbname;
}
%>
<%!
public static String del_dup(String tbName) {
	String sql =
			"delete p "+
			"from " + tbName +" as p "+
			"join ( "+
				"select min(seq) as min_seq,stuNum "+
				"from "+tbName+" as pp "+
				"group by stuNum "+
				"having count(*)>1 "+
				") as q "+
				"on q.stuNum = p.stuNum "+
				"where p.seq > q.min_seq;";
	return sql;
}
public static String up_num(String tbName) {
	String sql = 
			"update " + tbName +" as p "+
			"join ( "+
				"select stuNum, sum(num) as sum_num "+
				"from " + tbName +
				" group by stuNum "+
				"having count(*)>1 "+
			") AS j " +
			"ON j.stuNum = p.stuNum "+
			"SET p.num = j.sum_num;";
	return sql;
}
  
  public static String cal_score(String tbName) {
	  String sql;
	  switch (tbName) {
	  case "borrow_book":
		  sql =
			  "update borrow_book set score = "+
				"case "+
				"when num>=50 " +
					"then 5 "+
				"when num>=30 "+
					"then 3 "+
				"when num>=10 "+
					"then 1 "+
				"else 0 "+
				"end";
		  break;
	  case "use_avm":
		  sql =
		  "update use_avm set score = "+
			"case "+
			"when num>=10 " +
				"then 3 "+
			"when num>=5 "+
				"then 2 "+
			"when num>=3 "+
				"then 1 "+
			"else 0 "+
			"end";
		  break;
	  case "entry_lib":
		  sql =
		  "update entry_lib set score = "+
			"case "+
			"when num>=80 " +
				"then 6 "+
			"when num>=50 "+
				"then 3 "+
			"when num>=30 "+
				"then 1 "+
			"else 0 "+
			"end";
		  break;
	  case "ticket_seat":
		  sql =
		  "update ticket_seat set score = "+
			"case "+
			"when num>=50 " +
				"then 3 "+
			"when num>=30 "+
				"then 2 "+
			"when num>=20 "+
				"then 1 "+
			"else 0 "+
			"end";
		  break;
	  case "edu_user":
		  sql =
		  "update edu_user set score = "+
			"case "+
			"when num>=1 " +
				"then 2 "+
			"else 0 "+
			"end";
		  break;
	  case "review_book":
		  sql =
		  "update review_book set score = "+
			"case "+
			"when num>=7 " +
				"then 5 "+
			"when num>=5 "+
				"then 3 "+
			"when num>=3 "+
				"then 1 "+
			"else 0 "+
			"end";
		  break;
	  case "e_learning":
		  sql =
		  "update e_learning set score = "+
			"case "+
			"when num>=5 " +
				"then 8 "+
			"when num>=3 "+
				"then 5 "+
			"when num>=2 "+
				"then 3 "+
			"else 0 "+
			"end";
		  break;
	  case "part_lib":
		  sql = "update part_lib set score = num*3";
		  break;
	  case "e_book":
		  sql =
		  "update e_book set score = "+
			"case "+
			"when num>=30 " +
				"then 3 "+
			"when num>=20 "+
				"then 2 "+
			"when num>=10 "+
				"then 1 "+
			"else 0 "+
			"end";
		  break;
	  default:
		  sql = "";
	  }
	  return sql;
  }
  
  public static String sum_student() {
	  String sql="";
	  sql = 	"insert into student (stuNum) "+
				"select stuNum from borrow_book "+
			    "union "+
			    "select stuNum from use_avm "+
			    "union "+
			    "select stuNum from entry_lib "+
			    "union "+
			    "select stuNum from ticket_seat "+
			    "union "+
			    "select stuNum from edu_user "+
			    "union "+
			    "select stuNum from review_book "+
			    "union "+
			    "select stuNum from e_learning "+
			    "union "+
			    "select stuNum from part_lib "+
			    "union "+
			    "select stuNum from e_book;";
	  return sql;
  }
  
  public static String del_sum_student() {
	  String sql="delete from total;";
	  return sql;
  }
  public static String init_sum_student() {
	  String sql="alter table total auto_increment=1;";
	  return sql;
  }
  public static String insert_sum_student() {
	  String sql="insert into total (stuNum) "+
			  "select stuNum from student;";
	  return sql;
  }
  
  public static String update_total(String tbName) {
		String sql= "update total as a "+
				"join ( "+
				"select * "+
				"from " + tbName + " " +
				") as b "+
				"ON a.stuNum = b.stuNum ";
		
		if(tbName.equals("borrow_book")) {
			sql += "a.bb = b.score;";
		}else if(tbName.equals("use_avm")) {
			sql += "a.ua = b.score;";
		}else if(tbName.equals("entry_lib")) {
			sql += "a.enl = b.score;";
		}else if(tbName.equals("ticket_seat")) {
			sql += "a.ts = b.score;";
		}else if(tbName.equals("edu_user")) {
			sql += "a.eu = b.score;";
		}else if(tbName.equals("review_book")) {
			sql += "a.rb = b.score;";
		}else if(tbName.equals("e_learning")) {
			sql += "a.el = b.score;";
		}else if(tbName.equals("part_lib")) {
			sql += "a.pl = b.score;";
		}else if(tbName.equals("e_book")) {
			sql += "a.eb = b.score;";
		}else {
			
		}
		return sql;

  }

  
%>

<%
PreparedStatement ps = null;

String savePath=System.getProperty("java.io.tmpdir"); 
int sizeLimit=1024*1024*5;
String str_cell = null;
String sql="";
String tbName = request.getParameter("tbName");
String xlsfilename="";

MultipartRequest multi = new MultipartRequest(request,savePath,sizeLimit,new DefaultFileRenamePolicy());
String fileName1 = nullcheck(multi.getFilesystemName("addfile"));

XSSFWorkbook xworkBook = null;
XSSFSheet xsheet = null;
XSSFRow xrow = null;
XSSFCell xcell = null;


//엑셀을 위한 워크북 및 쉬트 등의 변수 선언
//HSSFWorkbook workbook = null;
//HSSFSheet sheet = null;
//HSSFRow row = null;
//HSSFCell cell = null;
//업로드된 경로와 파일의 이름을 얻는다.
xlsfilename=savePath+fileName1;

sql="insert into "+tbName+"(type,stuNum,num) values (?,?,?)";
 try
    {
			//엑�W파일이 업로드된 디렉토에서 해당 파일을 FileInputStream를 이용하여 읽어들이고
			//HSSFWorkbook의 객체를 새로 생성하여 파이의 내용을 세팅한다.
			FileInputStream fis = new FileInputStream(new File(xlsfilename));
			xworkBook = new XSSFWorkbook(fis);
			//workbook = new HSSFWorkbook(fis);		
			fis.close();
			if (xworkBook != null) {
				//불러들일 sheet
				xsheet = xworkBook.getSheetAt(0);

				if (xsheet != null) {

					//기록물철의 경우 실제 데이터가 시작되는 Row지정
					int nRowStartIndex = 1;
					//기록물철의 경우 실제 데이터가 끝 Row지정
					int nRowEndIndex = xsheet.getLastRowNum();

					//기록물철의 경우 실제 데이터가 시작되는 Column지정
					int nColumnStartIndex =0;
					//기록물철의 경우 실제 데이터가 끝나는 Column지정
					int nColumnEndIndex = xsheet.getRow(0).getLastCellNum();

					//해당 컬럼의 값을 받을 변수 선언
					String szValue = "";
%>
<%@ include file="./try.jsp" %>
<%
						ps = conn.prepareStatement(sql);
						
						for (int i = nRowStartIndex; i <= nRowEndIndex; i++) 
						{
							 xrow = xsheet.getRow(i);

								for (int nColumn = nColumnStartIndex; nColumn <= nColumnEndIndex; nColumn++)
								{
									
									 xcell = xrow.getCell((short) nColumn);
									 int score = 0;
									 
									 if(nColumn==nColumnStartIndex) {
										 String typeValue = xcell.getStringCellValue();
										 boolean valid = type_check(typeValue).equals(tbName);
										 if(!valid) {
											 int a = 3;
											 int b = 0;
											 int c = a/b;
										 }
									 }
									 //빈셀의 경우 
						
										 if (xcell == null) 
										 {
												continue;
										  }
										else
										   {
												switch(xcell.getCellType())
												{
												case XSSFCell.CELL_TYPE_FORMULA:
													szValue = xcell.getCellFormula().toString();
													ps.setString(nColumn+1,szValue);
													break;
												case XSSFCell.CELL_TYPE_NUMERIC:
													int a = (int)xcell.getNumericCellValue();
													ps.setInt(nColumn+1, a);
													break;
												case XSSFCell.CELL_TYPE_STRING:
													szValue = xcell.getStringCellValue();
													ps.setString(nColumn+1,szValue);
													break;
												default:
													szValue = "";
													ps.setString(nColumn+1,szValue);
												}
											}
								}
								ps.executeUpdate();	 
						}
						ps.close();
						stmt.executeUpdate(up_num(tbName));
						stmt.executeUpdate(del_dup(tbName));
						stmt.executeUpdate(cal_score(tbName));
						stmt.executeUpdate("delete from student;");
						stmt.executeUpdate("alter table student auto_increment=1;");
						stmt.executeUpdate(sum_student());
						stmt.executeUpdate(del_sum_student());
						stmt.executeUpdate(init_sum_student());
						stmt.executeUpdate(insert_sum_student());
						stmt.executeUpdate(update_total(tbName));
						stmt.close();
%>
<%@ include file="./finally.jsp" %>						
<%
			   }
    	}

    }
    catch( Exception e)
    {
        e.printStackTrace();
		out.println(e);
%>
<form name="formm" method="post" action="./mileage_admin2.jsp">
</form>
<script language="javascript">
	alert('오류가 발생하였습니다. 알맞은 파일인지 확인하여 주십시오.');
	document.formm.submit();
</script>
<%
    }
    finally
    {
       
%>
<form name="formm" method="post" action="./mileage_admin2.jsp">
</form>
<script language="javascript">
	alert('정상적으로 저장되었습니다.');
	document.formm.submit();
</script>
<%
    }
%>