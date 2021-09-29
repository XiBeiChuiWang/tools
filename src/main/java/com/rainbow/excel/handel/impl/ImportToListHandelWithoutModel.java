package com.rainbow.excel.handel.impl;




import com.rainbow.excel.entity.ResultList;
import com.rainbow.excel.handel.NoModelActuator;
import org.apache.commons.lang3.tuple.ImmutablePair;

import java.util.List;

public class ImportToListHandelWithoutModel implements NoModelActuator {

    private ResultList resultList;


    public ImportToListHandelWithoutModel(ResultList resultList) {
        this.resultList = resultList;
    }

    @Override
    public void execute(List rightList, List<ImmutablePair<Integer,Integer>> errorList) throws InterruptedException {
        resultList.setRightList(rightList);
        resultList.setErrorList(errorList);
    }
}
