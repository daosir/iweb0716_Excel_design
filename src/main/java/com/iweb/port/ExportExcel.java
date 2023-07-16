package com.iweb.port;

import com.iweb.data.Bill;

import java.io.File;
import java.util.List;

/**
 * @author ASUS
 * @Date 2023/7/15 10:50
 * @Version 1.8
 */
public interface ExportExcel {
    public File exportExcel(List<Bill> exportBills, Class<?> type, String writePath, int pageSize);
}
