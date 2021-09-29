package com.rainbow.excel.csv;



import com.alibaba.excel.annotation.format.DateTimeFormat;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.HashMap;

public class DateTimeFormatHandle {

    public static HashMap handleDateTimeFormat(Class clazz){

        Field[] declaredFields = clazz.getDeclaredFields();
        HashMap<Integer, SimpleDateFormat> map = new HashMap<>();
        for (int i = 0;i < declaredFields.length;i++){
            DateTimeFormat annotation = declaredFields[i].getAnnotation(DateTimeFormat.class);
            if (annotation != null){
                map.put(i, new SimpleDateFormat(annotation.value()));
            }
        }
        return map;
    }
}
