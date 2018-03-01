package main;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;

class JDBCUtil {

    static Connection getConnection() {
        Connection conn = null;
        String driver = "com.mysql.jdbc.Driver";
        String url = "jdbc:mysql://localhost:3306/citys";
        String user = "root";
        String password = "admin";
        try {
            Class.forName(driver);
            conn = DriverManager.getConnection(url, user, password);
        } catch (Exception e) {
            System.out.println("连接数据库异常" + e.getMessage());
        }
        return conn;
    }

    static void closeConnection(Connection conn, Statement st, ResultSet rs) {
        try {
            if (rs != null) {
                rs.close();
            }
            if (st != null) {
                st.close();
            }
            if (conn != null) {
                conn.close();
            }
        } catch (Exception e) {
            System.out.println("关闭连接Exception:" + e.getStackTrace());
        }
    }
}
