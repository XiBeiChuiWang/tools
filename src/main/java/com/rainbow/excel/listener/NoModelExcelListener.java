package com.rainbow.excel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.exception.ExcelDataConvertException;


import com.rainbow.excel.handel.NoModelActuator;
import org.apache.commons.lang3.tuple.ImmutablePair;

import java.util.ArrayList;
import java.util.Map;

public class NoModelExcelListener extends AnalysisEventListener<Map<Integer,String>> {

    private static final int DEFAULT_BATCH_COUNT = 3000;

    private int maxCount;

    private ArrayList headList = new ArrayList();

    private ArrayList rightList = new ArrayList();

    private ArrayList<ImmutablePair<Integer,Integer>> errorList = new ArrayList<>();

    private NoModelActuator noModelActuator;

    private boolean flag;


    public NoModelExcelListener(int maxCount, NoModelActuator noModelActuator, boolean flag) {
        this.noModelActuator = noModelActuator;
        this.flag = flag;
        if (flag){
            if (maxCount <= 0){
                this.maxCount = DEFAULT_BATCH_COUNT;
            }else {
                this.maxCount = maxCount;
            }
        }
    }

    @Override
    public void invoke(Map<Integer, String> data, AnalysisContext context) {
        rightList.add(data);
        if (flag && rightList.size() >= maxCount){
            try {
                noModelActuator.execute(rightList,errorList);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
            rightList.clear();
            errorList.clear();
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        try {
            if (!rightList.isEmpty() || !errorList.isEmpty()){
                noModelActuator.execute(rightList,errorList);
            }
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
    }

    @Override
    public void onException(Exception exception, AnalysisContext context){
        if (exception instanceof ExcelDataConvertException){
            exception = (ExcelDataConvertException) exception;
            ImmutablePair<Integer, Integer> integerIntegerImmutablePair = new ImmutablePair<>(((ExcelDataConvertException) exception).getRowIndex(), ((ExcelDataConvertException) exception).getColumnIndex());
//            log.error("第{}行解析失败",((ExcelDataConvertException) exception).getRowIndex());
            errorList.add(integerIntegerImmutablePair);
        }
    }
}
