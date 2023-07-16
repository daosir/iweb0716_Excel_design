package com.iweb.data;

import com.iweb.excelAnnotation.Excel;
import lombok.Data;

import java.math.BigDecimal;
import java.util.Date;

@Data


/**
 * @author ASUS
 * @Date 2023/7/15 10:49
 * @Version 1.8
 */
public class Bill {
    @Excel(name="账单编号")
    private Long id;

    @Excel(name="账单金额")
    private BigDecimal money;

    @Excel(name="账单发生时间")
    private Date createTime;

    @Excel(name="账单创建人")
    private String createUser;

    @Excel(ignore = true)
    private int version;

}
