package com.rainbow.excel.csv;


import com.opencsv.CSVReader;
import com.rainbow.excel.entity.ResultList;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.tuple.ImmutablePair;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.Reader;
import java.lang.reflect.Field;
import java.lang.reflect.Type;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

@Slf4j
public class CsvReader {

    private Reader reader;

    private CsvContext csvContext;

    private CSVReader csvReader;

    public CsvReader(Reader reader, Class clazz) {
        this.csvContext = ExcelPropertyHandle.handle(clazz);
        this.csvReader = new CSVReader(new BufferedReader(reader));
        this.reader = reader;
    }

    public CsvReader(Reader reader) {
        this.csvReader = new CSVReader(new BufferedReader(reader));
        this.reader = reader;
    }

    public CsvReader() {
    }

    public List<String[]> readCsvToStringListWithoutHead(){
        try {
            List<String[]> list = csvReader.readAll();
            int maxHead = csvContext.getMaxHead();
            for (int i = 0;i<maxHead;i++){
                list.remove(i);
            }
            return list;
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    public List<String[]> readCsvToStringListWithHead(){
        try {
            List<String[]> list = csvReader.readAll();
            return list;
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    public ResultList readModelList() {
        try {
            List<String[]> list = csvReader.readAll();
            int maxHead = csvContext.getMaxHead();

            if (list.size() < maxHead)
                throw new RuntimeException("表头错误");

            csvContext.setCurrentRow(csvContext.getMaxHead() + 1);
            for (int i = maxHead; i < list.size(); i++) {
                handleColl(list.get(i));
                csvContext.setCurrentRow(csvContext.getCurrentRow() + 1);
            }
            return csvContext.getResultList();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return csvContext.getResultList();
    }

    private void handleColl(String[] split) {
        List<Field> head = csvContext.getHead();
        Class headClass = csvContext.getHeadClass();
        HashMap<Integer, Boolean> allowNull = csvContext.getAllowNull();
        Object o = null;
        try {
            o = headClass.newInstance();
        } catch (Exception e) {
            e.printStackTrace();
        }

        for (int i = 0; i < csvContext.getCol(); i++) {
            Field field = csvContext.getHead().get(i);
            field.setAccessible(true);
            Type type = field.getType();
            String value = split[i];

            if (value == null ||value.length() < 1){
                if (allowNull.get(i)){
                    try {
                        field.set(o,null);
                        continue;
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }
                }else {
                    List errorList = csvContext.getResultList().getErrorList();
                    errorList.add(new ImmutablePair<Integer, Integer>(csvContext.getCurrentRow(), i));
                    return;
                }
            }

            Object value1 = null;
            try {
                switch (type.toString()) {
                    case "class java.lang.String":
                        value1 = value;
                        break;
                    case "int":
                        value1 = Integer.valueOf(value);
                        break;
                    case "class java.lang.Integer":
                        value1 = Integer.valueOf((String) value);
                        break;
                    case "boolean":
                        value1 = Boolean.valueOf((String) value);
                        break;
                    case "class java.lang.Boolean":
                        value1 = Boolean.valueOf((String) value);
                        break;
                    case "class java.lang.Double":
                        value1 = Double.valueOf((String) value);
                        break;
                    case "double":
                        value1 = Double.valueOf((String) value);
                        break;
                    case "long":
                        value1 = Long.valueOf((String) value);
                        break;
                    case "class java.lang.Long":
                        value1 = Long.valueOf((String) value);
                        break;
                    case "char":
                        if (((String) value).length() > 1) {
                            throw new RuntimeException(field + "属性转换异常");
                        }
                        value1 = ((String) value).charAt(0);
                        break;
                    case "class java.lang.Character":
                        if (((String) value).length() > 1) {
                            throw new RuntimeException(field + "属性转换异常");
                        }
                        value1 = ((String) value).charAt(0);
                        break;
                    case "short":
                        value1 = Short.valueOf((String) value);
                        break;
                    case "class java.lang.Short":
                        value1 = Short.valueOf((String) value);
                        break;
                    case "byte":
                        value1 = Byte.valueOf((String) value);
                        break;
                    case "class java.lang.Byte":
                        value1 = Byte.valueOf((String) value);
                        break;
                    case "class java.util.Date":
                        value1 = dateHandle((String) value, csvContext.getDateFormatHashMap().get(i));
                        if (value1 == null) {
                            throw new RuntimeException("错误");
                        }
                        break;
                }
                field.set(o, value1);
            } catch (Exception e) {
                List errorList = csvContext.getResultList().getErrorList();
                errorList.add(new ImmutablePair<Integer, Integer>(csvContext.getCurrentRow(), i));
                return;
            }
        }
        List rightList = csvContext.getResultList().getRightList();
        rightList.add(o);
    }

    private Date dateHandle(String value, SimpleDateFormat simpleDateFormat) throws ParseException {
        Date parse = null;

        if (simpleDateFormat != null) {
            parse = simpleDateFormat.parse(value);
            return parse;
        } else {
            if (value.contains("-")) {
                try {
                    SimpleDateFormat simpleDateFormat1 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                    parse = simpleDateFormat1.parse(value);
                    return parse;
                } catch (ParseException e) {
                    SimpleDateFormat simpleDateFormat1 = new SimpleDateFormat("yyyy-MM-dd");
                    parse = simpleDateFormat1.parse(value);
                    return parse;
                }
            } else {
                try {
                    SimpleDateFormat simpleDateFormat1 = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
                    parse = simpleDateFormat1.parse(value);
                    return parse;
                } catch (ParseException e) {
                    SimpleDateFormat simpleDateFormat1 = new SimpleDateFormat("yyyy/MM/dd");
                    parse = simpleDateFormat1.parse(value);
                    return parse;
                }
            }
        }
    }
}