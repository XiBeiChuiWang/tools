package com.rainbow.excel.handel;


import org.apache.commons.lang3.tuple.ImmutablePair;

import java.util.List;

@FunctionalInterface
public interface NoModelActuator {

    public void execute(List rightList, List<ImmutablePair<Integer, Integer>> errorList) throws InterruptedException;
}
