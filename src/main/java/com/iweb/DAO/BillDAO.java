package com.iweb.DAO;

import com.iweb.data.Bill;
import org.apache.poi.ss.formula.functions.T;

import java.util.Collection;
import java.util.List;

/**
 * @author ASUS
 * @Date 2023/7/15 15:58
 * @Version 1.8
 */
public interface BillDAO {

    /**
     * @param bill 用于通过bill的字段和注解，在本地mysql的test表空间中创建对应的表来存储数据
     */
    void create(Bill bill) throws Exception;


    /**
     * @param list 将获得的Bill集合导入导入数据库
     */
     void insertAll(List<Bill> list) throws Exception;


    /**
     * @return 将数据库的内容读出来，返回一个对象集合
     */
     Collection<Bill> collectionAll();

    /**
     * 在需要情况下可以完成导出excel之后清除数据库内容
     */
     void drop();



}
