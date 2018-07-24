package main;


import org.apache.log4j.Logger;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;

class JDBCUtil {

    private static Logger logger = Logger.getLogger(JDBCUtil.class.getName());

    static Connection getConnection() {
        Connection conn = null;
        String driver = "com.mysql.jdbc.Driver";
        String url = "jdbc:mysql://localhost:3306/citys?useUnicode=true&characterEncoding=utf-8";
        String user = "root";
        String password = "admin";
        try {
            Class.forName(driver);
            conn = DriverManager.getConnection(url, user, password);
        } catch (Exception e) {
            e.printStackTrace();
            logger.error("连接数据库异常" + e.getMessage());
        }
        return conn;
    }

    static Connection getAccessConnection() {
        Connection conn = null;
        String driver = "net.ucanaccess.jdbc.UcanaccessDriver";
        String url = "jdbc:ucanaccess://citys.mdb";
        String user = "";
        String password = "";
        try {
            Class.forName(driver);
            conn = DriverManager.getConnection(url, user, password);
        } catch (Exception e) {
            logger.error("连接数据库异常" + e.getMessage());
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
            logger.error("关闭连接Exception:" + e.getStackTrace());
        }
    }
}
