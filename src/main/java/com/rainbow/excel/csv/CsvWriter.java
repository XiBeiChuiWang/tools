package com.rainbow.excel.csv;


import com.opencsv.CSVWriter;

import java.io.IOException;
import java.io.Writer;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

public class CsvWriter {

    private CsvContext csvContext;

    private CSVWriter csvWriter;

    private Writer writer;

    public CsvWriter(Writer writer) {
        this.csvWriter = new CSVWriter(writer);
        this.writer = writer;
    }

    public CsvWriter(Writer writer, Class headClass) {
        this.csvContext = ExcelPropertyHandle.handle(headClass);
        this.csvWriter = new CSVWriter(writer);
        this.writer = writer;
        csvWriter.writeNext(csvContext.getExcelHeadName().toArray(new String[csvContext.getExcelHeadName().size()]));
    }

    public CsvWriter writeCsv(List data){
        HashMap dateTimeFormatMap = csvContext.getDateFormatHashMap();
        try {
            for (Object o : data) {
                if (!csvContext.getHeadClass().getName().equals(o.getClass().getName()))
                    throw new RuntimeException("类型不一致");
                List<Field> head = csvContext.getHead();
                ArrayList<String> strings1 = new ArrayList<>();
                for (int i = 0; i < head.size(); i++) {
                    Field field = head.get(i);
                    field.setAccessible(true);
                    if (!field.getType().toString().equals("class java.util.Date")) {
                        if (field.get(o)== null){
                            strings1.add(null);
                        }else {
                            strings1.add(field.get(o).toString());
                        }
                    } else {
                        SimpleDateFormat simpleDateFormat = (SimpleDateFormat) dateTimeFormatMap.get(i);
                        if (simpleDateFormat == null) {
                            simpleDateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
                        }
                        String format = simpleDateFormat.format((Date) field.get(o));
                        strings1.add(format);
                    }
                }
                writeListCsv(strings1);
            }
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        return this;
    }

    public synchronized CsvWriter writeListCsv(List<String> list){
        csvWriter.writeNext(list.toArray(new String[list.size()]));
        return this;
    }

    public Writer build(){
        try {
            writer.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return writer;
    }

}
