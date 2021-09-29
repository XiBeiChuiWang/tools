package com.rainbow.excel.handel.impl;




import com.rainbow.excel.entity.ResultList;
import com.rainbow.excel.handel.Actuator;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.commons.lang3.tuple.ImmutablePair;


import java.util.List;

@Getter
@Setter
@NoArgsConstructor
public class ImportToListHandel<T> implements Actuator<T> {

    private ResultList<T> resultList;


    public ImportToListHandel(ResultList<T> resultList) {
        this.resultList = resultList;
    }

    @Override
    public void execute(List<T> rightList, List<ImmutablePair<Integer,Integer>> errorList) throws InterruptedException {
        resultList.setRightList(rightList);
        resultList.setErrorList(errorList);
    }
}
