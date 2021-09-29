package com.rainbow.excel.entity;

import lombok.Getter;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
public class WriteExcelSheet<T> {

    //封装实体类
    private Class clazz;

    //数据
    private List<T> list;

    //sheet名
    private String sheetName;

    public WriteExcelSheet(Class clazz, List<T> list, String sheetName) {
        this.clazz = clazz;
        this.list = list;
        this.sheetName = sheetName;
    }
}
