package com.rainbow.excel.csv;

import com.rainbow.excel.entity.ResultList;
import lombok.Data;
import lombok.Getter;
import lombok.Setter;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.List;

@Getter
@Setter
@Data
public class CsvContext {

    private Class headClass;

    private int maxHead;

    private int col;

    private List<Field> head;

    private ResultList resultList= new ResultList();

    private int currentRow = 0;

    private HashMap<Integer, SimpleDateFormat> dateFormatHashMap;

    private List<String> ExcelHeadName;

    private HashMap<Integer,Boolean> allowNull;

}
