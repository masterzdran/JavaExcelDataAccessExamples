package pt.nunocancelo.examples.excel;

import java.io.File;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

public class JavaExcelDataAccessExample01 {
//https://www.microsoft.com/en-us/download/details.aspx?id=13255&mnui=1
	public static void main(String[] args) {
		Connection con = null;
		try {

			if ((new File("C:/dev/github/JavaExcelDataAccessExamples/data/sample-data.xls")).exists()) {
				System.out.println("O ficheiro existe.");
			}

		    String driver = "sun.jdbc.odbc.JdbcOdbcDriver";
		    String url = "jdbc:odbc:Driver={Microsoft Excel Driver (*.xls)};;DBQ=C:/dev/github/JavaExcelDataAccessExamples/data/sample-data.xls;DriverID=22;READONLY=false";
		    String username = "yourName";
		    String password = "yourPass";
		    Class.forName(driver);
		    con =  DriverManager.getConnection(url, username, password);

		    
		    Class.forName( "sun.jdbc.odbc.JdbcOdbcDriver" );
		    
		  //using DSN-less connection in 32 bit env
		 // con = DriverManager.getConnection( "jdbc:odbc:Driver={Microsoft Excel Driver (*.xls)};DBQ=C:/Abhinay/test.xls")
		   
		  //using DSN-less connection in 64 bit env
		   
		 // con = DriverManager.getConnection( "jdbc:odbc:DRIVER={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=C:/Abhinay/test.xls");
		   
		    
		    
		    
		    Statement stmt = null;
		    ResultSet rs = null;
		    PreparedStatement ps = null;
		    
		    //String createStm = "create table [Teste] (DEPT_ID NUMBER,DEPT_NAME STRING)";
		    // stmt.executeUpdate(createStm);
		    //stmt.close();
		    //Numeric, String, Boolean
		    //String excelQuery = "insert into [Teste$](DEPT_ID, DEPT_NAME) values(?, ?)";
		    //ps = con.prepareStatement(excelQuery);
		    //for (int i =0;i<10;i++)
		    //{
		    //	ps.setInt(1,i);
		    //	ps.setString(2, "Example"+i);
		    //	ps.addBatch();
		    	
		    //}
		    // ps.executeBatch();
		    
		    
//		    String selectStm = " select DEPT_ID, DEPT_NAME from [Teste$] where DEPT_ID = ?";
//		    ps = con.prepareStatement(selectStm);
//		    ps.setInt(1, 5);
//		    
//		    rs = ps.executeQuery();
//		    
//		    while (rs.next())
//		    {
//		    	System.out.println(rs.getInt(1) + " --- "+rs.getString(2));
//		    	
//		    }
//		    
//		    String updateStm = "update [Teste$] set DEPT_NAME = ? WHERE  DEPT_ID = ?";
//		    
//		    ps = con.prepareStatement(updateStm);
//		    ps.setInt(2, 5);
//		    ps.setString(1, "Changed Excel Again");
//		    int res = ps.executeUpdate();
		    

//		    ps = con.prepareStatement(selectStm);
//		    ps.setInt(1, 5);
//		    
//		    rs = ps.executeQuery();
//		    
//		    while (rs.next())
//		    {
//		    	System.out.println(rs.getInt(1) + " --- "+rs.getString(2));
//		    	
//		    }    
		    
		    //java.sql.SQLException: [Microsoft][ODBC Excel Driver] Deleting data in a linked table is not supported by this ISAM.
//		    String deleteStm = "delete from [Teste$]  WHERE  DEPT_ID = ?";
//		    ps = con.prepareStatement(deleteStm);
//		    ps.setInt(1, 2);
//		     ps.executeUpdate();
//		    String selectStm = " select DEPT_ID, DEPT_NAME from [Teste$] ";
//		    ps = con.prepareStatement(selectStm);
//		    rs = ps.executeQuery();
//		    while (rs.next())
//		    {
//		    	System.out.println(rs.getInt(1) + " --- "+rs.getString(2));
//		    }
		    
		    //NO work
//		    String createStm = "drop table [Delete$]";
//		    stmt = con.createStatement();
//		    int x = stmt.executeUpdate(createStm);
//		    System.out.println(x);
//		    stmt.close();
		    
		    
		    //rs.close();
		    //ps.close();
		    con.close();
		    
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			try {
				if (con != null && !con.isClosed()) {
					con.close();
				}
			} catch (Exception e) {
			}
		}

	}

}
