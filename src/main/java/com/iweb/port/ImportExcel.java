package com.iweb.port;

import com.iweb.data.Bill;

import java.io.File;
import java.util.List;

/**
 * @author ASUS
 * @Date 2023/7/15 10:50
 * @Version 1.8
 */
public interface ImportExcel {

    public List<Bill> importExcel(File excelFile, Class<?> type);

}
