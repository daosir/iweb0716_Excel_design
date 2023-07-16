package com.iweb.port;


import com.iweb.DAO.BillDAO;
import com.iweb.DAO.Impl.BillDAOImpl;
import com.iweb.data.Bill;
import com.iweb.excelAnnotation.Excel;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.sql.Date;
import java.sql.Time;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.List;

/**
 * @author ASUS
 * @Date 2023/7/15 10:51
 * @Version 1.8
 */
public class PortUtils implements ExportExcel, ImportExcel {


    /**
     * 将数据库的数据放入Bill对象集合再整合成excel文件输出
     *
     * @param exportBills 从数据库中取出的Bill对象集合
     * @param type        判断xls或xlsx格式的excel文件输出
     * @param writePath   将文件写到的本地地址
     * @param pageSize    每一页的行数限制
     * @return 返回一个excel文件
     */
    @Override
    public File exportExcel(List<Bill> exportBills, Class<?> type, String writePath, int pageSize) {
        if (type == HSSFWorkbook.class) {
//            基于poi创建excel表
            HSSFWorkbook workbook = new HSSFWorkbook();
  //            创建单元格日期格式
            CellStyle style = workbook.createCellStyle();
            CreationHelper creationHelper = workbook.getCreationHelper();
            style.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-MM-dd"));
//            通过设置的pageSize和bill集合长度，来设置sheet的数量
            int sheets = 1;
            if (exportBills.size() > pageSize) {
                int count = exportBills.size();
                while (true) {
                    count -= pageSize;
                    sheets++;
                    if (count <= pageSize) {
                        break;
                    }
                }
            }
//            根据表的数量创建相应的内容
            for (int k = 0; k < sheets; k++) {
//                换到下一页的时候需要将list的相应前面部分删除，来完成分页连续读取数据的功能
                List<Bill> removeList = exportBills.subList(0,pageSize);
                if (k>0){
                    exportBills.removeAll(removeList);
                }
//            创建工作表
                HSSFSheet sheet = workbook.createSheet("sheet"+(k+1));
                sheet.setRowBreak(pageSize);
//           创建表头，即Bill的注解内容
                HSSFRow firstRow = sheet.createRow(0);
//            获取Bill的class类，获取field字段，通过字段获取注解
                Class billClass = new Bill().getClass();
                Field[] fields = billClass.getDeclaredFields();
                for (int i = 0; i < fields.length; i++) {
                    fields[i].setAccessible(true);
//                创建单元格，来获取注解(cell.setCellValue();)
                    HSSFCell cell = firstRow.createCell(i);
//                将注解作为excel表的首段输入
                    String head = fields[i].getAnnotation(Excel.class).name();
                    cell.setCellValue(head);
                }
//            根据集合长度，根据对象数量创建行
                for (int i = 1; i <= (pageSize>=exportBills.size()?exportBills.size():pageSize); i++) {
                    HSSFRow dataRow = sheet.createRow(i);
//                每一行逐个创建单元格，将内容具体的字段的值赋值进去
                    for (int j = 0; j < fields.length; j++) {
                        HSSFCell cell = dataRow.createCell(j);
//            先利用字段的get方法，获取字段存储的值，再通过比对注解，使得cell里的格式与字段的一致
                        setCellValue(fields[j], exportBills.get(i - 1), cell);
                        if (fields[j].getAnnotation(Excel.class).name().equals("账单发生时间")){
                            cell.setCellStyle(style);
                        }
                    }
                }
            }
            //            创建文件输出流 将已经做好的excel文件写入预先设定的地址
            File excelFile = new File(writePath);
            try (
                    FileOutputStream fos = new FileOutputStream(excelFile)
            ) {
                workbook.write(fos);
            } catch (Exception e) {
                e.printStackTrace();
            }
            return excelFile;
        } else if (type == XSSFWorkbook.class) {
//            基于poi创建excel表
            XSSFWorkbook workbook = new XSSFWorkbook();
//            创建单元格日期格式
            CellStyle style = workbook.createCellStyle();
            CreationHelper creationHelper = workbook.getCreationHelper();
            style.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-MM-dd"));
//            通过设置的pageSize和bill集合长度，来设置sheet的数量
            int sheets = 1;
            if (exportBills.size() > pageSize) {
                int count = exportBills.size();
                while (true) {
                    count -= pageSize;
                    sheets++;
                    if (count <= pageSize) {
                        break;
                    }
                }
            }
//            根据表的数量创建相应的内容
            for (int k = 0; k < sheets; k++) {
//                换到下一页的时候需要将list的相应前面部分删除，来完成分页连续读取数据的功能
                List<Bill> removeList = exportBills.subList(0,pageSize);
                if (k>0){
                    exportBills.removeAll(removeList);
                }
//            创建工作表
                XSSFSheet sheet = workbook.createSheet("sheet"+(k+1));
                sheet.setRowBreak(pageSize);
//           创建表头，即Bill的注解内容
                XSSFRow firstRow = sheet.createRow(0);
//            获取Bill的class类，获取field字段，通过字段获取注解
                Class billClass = new Bill().getClass();
                Field[] fields = billClass.getDeclaredFields();
                for (int i = 0; i < fields.length; i++) {
                    fields[i].setAccessible(true);
//                创建单元格，来获取注解(cell.setCellValue();)
                    XSSFCell cell = firstRow.createCell(i);
//                将注解作为excel表的首段输入
                    String head = fields[i].getAnnotation(Excel.class).name();
                    cell.setCellValue(head);
                }
//            根据集合长度，根据对象数量创建行
                for (int i = 1; i <= (pageSize>=exportBills.size()?exportBills.size():pageSize); i++) {
                    XSSFRow dataRow = sheet.createRow(i);
//                每一行逐个创建单元格，将内容具体的字段的值赋值进去
                    for (int j = 0; j < fields.length; j++) {
                        XSSFCell cell = dataRow.createCell(j);
//            先利用字段的get方法，获取字段存储的值，再通过比对注解，使得cell里的格式与字段的一致
                        setCellValue(fields[j], exportBills.get(i - 1), cell);
                        if (fields[j].getAnnotation(Excel.class).name().equals("账单发生时间")){
                            cell.setCellStyle(style);
                        }
                    }
                }
            }
            //            创建文件输出流 将已经做好的excel文件写入预先设定的地址
            File excelFile = new File(writePath);
            try (
                    FileOutputStream fos = new FileOutputStream(excelFile)
            ) {
                workbook.write(fos);
            } catch (Exception e) {
                e.printStackTrace();
            }
            return excelFile;
        } else {
            System.out.println("你的输出类型有误，请重新输入");
            return null;
        }
    }

    /**
     * 导入excel（xls/xlsx）表
     *
     * @param excelFile 需要导入的表文件
     * @param type      选择导入的表文件的格式
     * @return 返回一个bill对象的集合 再对数据库进行存入
     */
    @Override
    public List<Bill> importExcel(File excelFile, Class<?> type) {

        List<Bill> list = new ArrayList<>();
//        判断版本
        if (type == HSSFWorkbook.class) {
//            文件流输入
            try (FileInputStream fis = new FileInputStream(excelFile)
            ) {
//                基于poi创建excel的workbook
                Workbook workbook = WorkbookFactory.create(fis);
                HSSFWorkbook hb = (HSSFWorkbook) workbook;
//                获取表（sheet）的数量
                int sheets = hb.getNumberOfSheets();
                for (int i = 0; i < sheets; i++) {
//                    获取表
                    HSSFSheet sheet = hb.getSheetAt(i);
                    if (sheet == null) {
                        continue;
                    }
//                    获取每个表里行的数量
                    int rows = sheet.getLastRowNum();
//                    通过注解判断导入的excel是否是按照注解的形式导入
                    HSSFRow firstRow = sheet.getRow(0);
//                    获得excel的第一行，来获取对应的列名
//                    并进行判断，如果与所需的Bill账单不同，那么直接返回
                    if (firstRow == null) {
                        continue;
                    } else {
                        Bill neededBill = new Bill();
                        Class neededBillClass = neededBill.getClass();
                        Field[] neededFields = neededBillClass.getDeclaredFields();
//                        对第一行遍历注意比对
                        for (int j = 0; j < neededFields.length; j++) {
                            HSSFCell cell = firstRow.getCell(j);
                            if (neededFields[j].getAnnotation(Excel.class).ignore() == true) {
                                continue;
                            } else if (cell == null) {
                                System.out.println("该表不符合账单要求，请重新导入");
                                return null;
                            } else if (neededFields[j].getAnnotation(Excel.class).name().equals(cell.getStringCellValue())) {
                                continue;
                            } else {
                                System.out.println("该表不符合账单要求，请重新导入");
                                return null;
                            }
                        }
                    }
                    for (int j = 1; j <= rows; j++) {
//                        获取行
                        HSSFRow row = sheet.getRow(j);
//                        每一行对应了一个bill对象
                        Bill bill = new Bill();
                        Class billClass = bill.getClass();
                        Field[] fields = billClass.getDeclaredFields();
                        if (row != null) {
//                            获得每个单元格
                            for (int k = 0; k < fields.length; k++) {
                                fields[k].setAccessible(true);
                                HSSFCell cell = row.getCell(k);
                                if (cell == null) {
                                    fields[k].set(bill, null);
                                }
                                fields[k].set(bill, getXlsValue(cell, fields[k].getAnnotation(Excel.class)));
                            }
                        }
                        list.add(bill);
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        } else if (type == XSSFWorkbook.class) {
            //            文件流输入
            try (FileInputStream fis = new FileInputStream(excelFile)
            ) {
//                基于poi创建excel的workbook
                Workbook workbook = WorkbookFactory.create(fis);
                XSSFWorkbook hb = (XSSFWorkbook) workbook;
//                获取表（sheet）的数量
                int sheets = hb.getNumberOfSheets();
                for (int i = 0; i < sheets; i++) {
//                    获取表
                    XSSFSheet sheet = hb.getSheetAt(i);
                    if (sheet == null) {
                        continue;
                    }
//                    获取每个表里行的数量
                    int rows = sheet.getLastRowNum();
//                    通过注解判断导入的excel是否是按照注解的形式导入
                    XSSFRow firstRow = sheet.getRow(0);
                    if (firstRow == null) {
                        continue;
                    } else {
                        Bill neededBill = new Bill();
                        Class neededBillClass = neededBill.getClass();
                        Field[] neededFields = neededBillClass.getDeclaredFields();
                        for (int j = 0; j < neededFields.length; j++) {
                            XSSFCell cell = firstRow.getCell(j);
                            if (neededFields[j].getAnnotation(Excel.class).ignore() == true) {
                                continue;
                            } else if (cell == null) {
                                System.out.println("该表不符合账单要求，请重新导入");
                                return null;
                            } else if (neededFields[j].getAnnotation(Excel.class).name().equals(cell.getStringCellValue())) {
                                continue;
                            } else {
                                System.out.println("该表不符合账单要求，请重新导入");
                                return null;
                            }
                        }
                    }
                    for (int j = 1; j <= rows; j++) {
//                        获取行
                        XSSFRow row = sheet.getRow(j);
//                        每一行对应了一个bill对象
                        Bill bill = new Bill();
                        Class billClass = bill.getClass();
                        Field[] fields = billClass.getDeclaredFields();

                        if (row != null) {
//                            获得每个单元格
                            for (int k = 0; k < fields.length; k++) {
                                fields[k].setAccessible(true);
                                XSSFCell cell = row.getCell(k);
                                if (cell == null) {
                                    fields[k].set(bill, null);
                                }
                                fields[k].set(bill, getXlsxValue(cell, fields[k].getAnnotation(Excel.class)));
                            }
                        }
                        list.add(bill);
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        } else {
            return null;
        }
        return list;
    }

    /**
     * 用于获取单元格内的内容
     *
     * @param cell 单元格
     * @return 将各种格式通过注解的值（这里认为excel表中的列对应好了Bill的各项属性）
     */
    public Object getXlsValue(HSSFCell cell, Excel excel) {
//        布尔值的时候
        if (cell.getCellType() == CellType.BOOLEAN) {
            return cell.getBooleanCellValue();
//            如果单元格内是数字
        } else if (cell.getCellType() == CellType.NUMERIC) {
//            由于Bill类中的数字属性有两种种存在模式
//            分成两种方式依次判断
            if (DateUtil.isCellDateFormatted(cell)) {
                return cell.getDateCellValue();
//                利用bill的注解，来进行判断
            } else {
                if (excel.name().equals("账单金额")) {
                    return new BigDecimal(cell.getNumericCellValue());
                } else if (excel.name().equals("账单编号")) {
                    double temp = Math.floor(cell.getNumericCellValue());
                    return (long) temp;
                } else {
                    return (int) cell.getNumericCellValue();
                }
            }
        } else {
//            其他情况均返回字符串
            return cell.getStringCellValue();
        }

    }

    public Object getXlsxValue(XSSFCell cell, Excel excel) {
//        布尔值的时候
        if (cell.getCellType() == CellType.BOOLEAN) {
            return cell.getBooleanCellValue();
//            如果单元格内是数字
        } else if (cell.getCellType() == CellType.NUMERIC) {
//            由于Bill类中的数字属性有两种种存在模式
//            分成两种方式依次判断
            if (DateUtil.isCellDateFormatted(cell)) {
                return cell.getDateCellValue();
//                利用bill的注解，来进行判断
            } else {
                if (excel.name().equals("账单金额")) {
                    return new BigDecimal(cell.getNumericCellValue());
                } else if (excel.name().equals("账单编号")) {
                    double temp = Math.floor(cell.getNumericCellValue());
                    return (long) temp;
                } else {
                    return (int) cell.getNumericCellValue();
                }
            }
        } else {
//            其他情况均返回字符串
            return cell.getStringCellValue();
        }

    }

    /**
     * 将导入拆分为两部分，这一步将已经导入到Java的Bill集合导入到数据库
     * 方法中调用了sql语句，会根据Bill字段的属性创建表，并且将内容导入
     *
     * @param list 装着excel数据的集合
     * @return 返回一个是否成功装入数据库的布尔值
     */
    public boolean importExcelDB(List<Bill> list) {

        if (list.isEmpty()){
            return false;
        }
        try {
            BillDAO DAO = new BillDAOImpl();
            DAO.create(new Bill());
            DAO.insertAll(list);
        } catch (Exception e) {
            return false;
        }
        return true;
    }


    /**
     * 将已经在数据库的账单信息导出，成为一个对象集合，并将数据库的表进行删除
     *
     * @return 返回包含着所有账单对象信息的集合
     */
    public List<Bill> exportExcelDB() {

        BillDAO DAO = new BillDAOImpl();
        List<Bill> list = (List<Bill>) DAO.collectionAll();
        DAO.drop();
        if (list == null) {
            System.out.println("你的数据库中没有信息");
            return null;
        } else {
            return list;
        }

    }

    /**
     * 创建方法将field内的属性根据对应格式放入cell单元格
     *
     * @param field 存放内容的字段
     * @param bill  单元格对应的账单
     * @param cell  单元格
     */
    public void setCellValue(Field field, Bill bill, Cell cell) {
        if ("账单编号".equals(field.getAnnotation(Excel.class).name())) {
            try {
                cell.setCellValue((long) field.get(bill));
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
        } else if ("账单金额".equals(field.getAnnotation(Excel.class).name())) {
            try {
                BigDecimal money = (BigDecimal) field.get(bill);
                cell.setCellValue(money.doubleValue());
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
        } else if ("账单发生时间".equals(field.getAnnotation(Excel.class).name())) {
            try {
                Timestamp timestamp = (Timestamp) field.get(bill);
                java.util.Date date = new java.util.Date(timestamp.getTime());
                cell.setCellValue(date);
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
        } else if ("账单创建人".equals(field.getAnnotation(Excel.class).name())) {
            try {
                cell.setCellValue((String) field.get(bill));
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
        }
    }


}
