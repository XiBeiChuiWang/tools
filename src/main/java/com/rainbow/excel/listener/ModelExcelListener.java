package com.rainbow.excel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.exception.ExcelDataConvertException;


import com.rainbow.excel.handel.Actuator;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.tuple.ImmutablePair;

import java.util.ArrayList;

@Slf4j
public class ModelExcelListener<T> extends AnalysisEventListener<T>{

    private static final int DEFAULT_BATCH_COUNT = 3000;

    private int maxCount;

    private ArrayList<T> rightList = new ArrayList();

    private ArrayList<ImmutablePair<Integer,Integer>> errorList = new ArrayList<>();

    private Actuator<T> actuator;

    private boolean flag;


    @Override
    public void invoke(T t, AnalysisContext analysisContext) {
        System.out.println(t);
        rightList.add(t);
        if (flag && rightList.size() >= maxCount){
            try {
                actuator.execute(rightList,errorList);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
            rightList.clear();
            errorList.clear();
        }
    }

    public ModelExcelListener(int maxCount, Actuator<T> actuator, boolean flag) {
        this.actuator = actuator;
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
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        try {
            if (!rightList.isEmpty() || !errorList.isEmpty()){
                actuator.execute(rightList,errorList);
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