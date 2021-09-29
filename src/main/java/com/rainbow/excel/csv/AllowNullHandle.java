package com.rainbow.excel.csv;

import java.lang.reflect.Field;
import java.util.HashMap;

public class AllowNullHandle {
    public static HashMap handle(Class clazz){

        Field[] declaredFields = clazz.getDeclaredFields();
        HashMap<Integer, Boolean> map = new HashMap<>();
        for (int i = 0;i < declaredFields.length;i++){
            AllowNull annotation = declaredFields[i].getAnnotation(AllowNull.class);
            if (annotation != null){
                map.put(i, annotation.value());
            }else {
                map.put(i,true);
            }
        }
        return map;
    }
}
