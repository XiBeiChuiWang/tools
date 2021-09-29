package com.rainbow.excel.entity;


import com.rainbow.excel.handel.Actuator;
import lombok.Getter;


@Getter
public class ReadExcelSheet {

    //封装实体类，实体类注解见文档
    private Class clazz;

    //如excel很大，将十分占用内存，因此每读取 onceCount 条数据执行一次操作（当flag为true时生效）
    private int onceCount;

    //执行操作，lambda表达式
    private Actuator actuator;


    //是否分批次执行操作
    private Boolean flag;

    //默认flag为true
    public ReadExcelSheet(Class clazz, int onceCount, Actuator actuator) {
        this.clazz = clazz;
        this.onceCount = onceCount;
        this.actuator = actuator;
        this.flag = true;
    }

    public ReadExcelSheet(Class clazz, Actuator actuator) {
        this.clazz = clazz;
        this.actuator = actuator;
        this.flag = false;
    }
}
