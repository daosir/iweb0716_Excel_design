package com.iweb.DBUtil;

import java.sql.Connection;
import java.sql.DriverManager;
import java.util.ArrayList;
import java.util.List;

/**
 * @author ASUS
 * @Date 2023/7/15 16:24
 * @Version 1.8
 */
public class DBConnectionPool {

    /**
     * 定义连接池的长度
     */
    int size;

    /**
     * 容器
     */
    List<Connection> list = new ArrayList<>();

    public DBConnectionPool(int size) {
        this.size = size;
        init();
    }
    private void init() {
        try {
            Class.forName("com.mysql.jdbc.Driver");
            for (int i = 0; i < size; i++) {
                Connection c = DriverManager.getConnection
                        ("jdbc:mysql://localhost:3306/test?characterEncoding=utf8", "root", "a12345");
                list.add(c);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * @return 获取连接池中的连接
     */
    public synchronized Connection getConnection() {
        while(list.isEmpty()){
            try {
                this.wait();
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }
        Connection c = list.remove(0);
        return c;
    }

    public synchronized void returnConnection(Connection c ){
        list.add(c);
        this.notifyAll();
    }


}
