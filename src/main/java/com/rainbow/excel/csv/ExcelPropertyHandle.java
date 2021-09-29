package com.rainbow.excel.csv;


import com.alibaba.excel.annotation.ExcelProperty;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;

public class ExcelPropertyHandle {

    public static CsvContext handle(Class clazz) {
        if (clazz == null) {
            throw new RuntimeException("传入参数不能为空");
        }
        CsvContext csvContext = new CsvContext();
        csvContext.setHeadClass(clazz);
        csvContext.setHead(new ArrayList<Field>());
        Field[] fields = clazz.getDeclaredFields();
        int max = -1;
        ArrayList<String> strings = new ArrayList<>();
        for (Field field : fields) {
            field.setAccessible(true);
            ExcelProperty annotation = field.getAnnotation(ExcelProperty.class);
            if (annotation != null) {
                String[] value1 = annotation.value();
                if (value1.length == 0) {
                    strings.add(field.getName());
                } else {
                    strings.add(value1[0]);
                }
                max = Math.max(value1.length, max);
                csvContext.getHead().add(field);
                csvContext.setCol(csvContext.getCol() + 1);
            }
        }

        csvContext.setMaxHead(max);
        csvContext.setExcelHeadName(strings);

        HashMap dateTimeFormat = DateTimeFormatHandle.handleDateTimeFormat(clazz);
        csvContext.setDateFormatHashMap(dateTimeFormat);

        HashMap handle = AllowNullHandle.handle(clazz);
        csvContext.setAllowNull(handle);
        return csvContext;
    }
}
