package com.iweb;

import static org.junit.Assert.assertTrue;

import com.iweb.DAO.BillDAO;
import com.iweb.DAO.Impl.BillDAOImpl;
import com.iweb.DBUtil.DBConnectionPool;
import com.iweb.data.Bill;
import com.iweb.excelAnnotation.Excel;
import com.iweb.port.ImportExcel;
import com.iweb.port.PortUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import javax.sound.sampled.Port;
import java.awt.*;
import java.io.File;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * Unit test for simple App.
 */
public class AppTest {
    /**
     * Rigorous Test :-)
     */
    @Test
    public void shouldAnswerWithTrue() {
        assertTrue(true);
    }

    public static void main(String[] args) {


//        获得Bill集合实验

        PortUtils util  = new PortUtils();
        List<Bill> list = new ArrayList<>();
        list = util.exportExcelDB();
        for (Bill bill:list){
            System.out.print(bill.getId());
            System.out.print(bill.getCreateTime());
            System.out.print(bill.getMoney());
            System.out.print(bill.getCreateUser());
            System.out.println(bill.getVersion());
        }
        util.exportExcel(list, HSSFWorkbook.class,"F:\\Javafile_test\\12345.xlsx",5);



//PortUtils utils = new PortUtils();


//        List<Bill> billList  = new ArrayList<>();
//        Bill bill = new Bill();
//        bill.setId((long) 1);;
//        bill.set

//        导入实验
//
//        File file = new File("F:\\Javafile_test\\123.xlsx");
//        PortUtils utils = new PortUtils();
//        List<Bill> bills = utils.importExcel(file, XSSFWorkbook.class);
//        System.out.println(utils.importExcelDB(bills));

//        List<Bill> list = (List<Bill>) new BillDAOImpl().collectionAll();
//        for (Bill bill :list){
//                        System.out.print(bill.getId());
//            System.out.print(bill.getCreateTime());
//            System.out.print(bill.getMoney());
//            System.out.print(bill.getCreateUser());
//            System.out.println(bill.getVersion());
//        }



//        BillDAO b = new BillDAOImpl();
//        b.insertAll(bills);

//        b.create(new Bill());



//        for (Bill bill : bills) {
//            System.out.print(bill.getId());
//            System.out.print(bill.getCreateTime());
//            System.out.print(bill.getMoney());
//            System.out.print(bill.getCreateUser());
//            System.out.println(bill.getVersion());

//        Bill bill= new Bill();
//        Class billClass = bill.getClass();
//        Field [] fields = billClass.getDeclaredFields();
//        System.out.println(Excel.class);
//        for (Field f: fields){
//            System.out.println(f.toString());
//            System.out.println(f.getAnnotation(Excel.class));
//        }


//        }
    }
}

