package com.iweb.DAO.Impl;

import com.iweb.DAO.BillDAO;
import com.iweb.DBUtil.DBConnectionPool;
import com.iweb.data.Bill;
import org.apache.poi.ss.formula.functions.T;

import java.sql.Connection;
import java.sql.Date;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

/**
 * @author ASUS
 * @Date 2023/7/15 19:35
 * @Version 1.8
 */
public class BillDAOImpl implements BillDAO {
    /**
     * @param bill 用于通过bill的字段和注解，在本地mysql的test表空间中创建对应的表来存储数据
     */
    @Override
    public void create(Bill bill) throws Exception {
        String sql = "CREATE TABLE Bills (\n" +
                "bid BIGINT PRIMARY KEY ,\n" +
                "bamount DECIMAL (10,2) ,\n" +
                "bdate DATE ,\n" +
                "bcreator VARCHAR(50) );";
        try (
                Connection c = new DBConnectionPool(1).getConnection();
                PreparedStatement ps = c.prepareStatement(sql);
        ) {
            ps.execute();
        } catch (Exception e) {
            System.out.println("表已经存在");
            throw  new Exception();
        }
    }

    /**
     * @param list 将获得的Bill集合导入导入数据库
     */
    @Override
    public void insertAll(List<Bill> list) throws Exception {

//        写出将所有对象内容插入数据库的sql语句
        String sql = "INSERT INTO bills (bid,bamount,bdate,bcreator) VALUES (?,?,?,?) ";
//        遍历从excel表得来的bill集合，将数据传入数据库
        for (Bill bill :list) {
            try (
                    Connection c = new DBConnectionPool(1).getConnection();
                    PreparedStatement ps = c.prepareStatement(sql)
            ) {
                ps.setLong(1,bill.getId() );
                ps.setBigDecimal(2,bill.getMoney());
//                这里将java.util.Date转换成java.sql.date
                java.sql.Date time = new Date(bill.getCreateTime().getTime());
                ps.setDate(3, time);
                ps.setString(4,bill.getCreateUser());
                ps.execute();
            } catch (Exception e) {
                System.out.println("未能成功插入");
                throw new Exception();
            }

        }
    }

    /**
     * @return 将数据库的内容读出来，返回一个对象集合
     */
    @Override
    public Collection<Bill> collectionAll() {
//        书写sql查语句
        String sql = "SELECT * FROM bills;\n";
        List<Bill> list = new ArrayList<>();
        try(
                Connection c = new  DBConnectionPool(1).getConnection();
                PreparedStatement ps = c.prepareStatement(sql)
                ){
//            将查到的内容放入集合，进而放入list
            ResultSet rs = ps.executeQuery();
            while (rs.next()){
                Bill bill = new Bill();
                bill.setId(rs.getLong("bid"));
                bill.setMoney(rs.getBigDecimal("bamount"));
                bill.setCreateTime(rs.getTimestamp("bdate"));
                bill.setCreateUser(rs.getString("bcreator"));
                list.add(bill);
            }
        }catch (Exception e ){
            e.printStackTrace();
        }
        return (list.isEmpty()?null:list);
    }

    /**
     * 在需要情况下可以完成导出excel之后清除数据库内容
     */
    @Override
    public void drop() {
        String sql= "DROP TABLE bills;\n";
        try(
                Connection c = new DBConnectionPool(1).getConnection();
                PreparedStatement ps = c.prepareStatement(sql)
                ){
            ps.execute();
        }catch (Exception e ){
            e.printStackTrace();
        }
    }
}
