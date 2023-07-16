package com.iweb;

import com.iweb.data.Bill;
import com.iweb.port.PortUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.util.List;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
//       导入excel功能，在包中包含了一个xlsx格式的表格去，格式与注解要求一致
        PortUtils util = new PortUtils();

//        导入excel成为一个Bill集合(具体文件位置需要自己修改)
        List<Bill> bills = util.importExcel(new File("F:\\Javafile_test\\designExcel_maven\\123.xlsx"), XSSFWorkbook.class);
//        将bills导入数据库（数据库连接的是localhost的test表空间）（封装了自动创建bills表的功能）
        System.out.println(util.importExcelDB(bills));

//        从数据库中导出功能（导出之后会数据库的bills表会自动删除）
        List<Bill>newBills = util.exportExcelDB();

//        导出excel（测试表格内容较少，为了展示分页功能，将pagesize设置的较小）
        util.exportExcel(newBills,HSSFWorkbook.class,"F:\\Javafile_test\\designExcel_maven\\1234.xls",5);

//        在完成之后，会在包里看到一个和123.xlsx内容一致的1234.xls表格




    }
}
