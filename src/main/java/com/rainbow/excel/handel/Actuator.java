package com.rainbow.excel.handel;

import org.apache.commons.lang3.tuple.ImmutablePair;

import java.util.List;

/**
 * @Author: zx
 * @Date: 2021/8/5 23:03
 */
@FunctionalInterface
public interface Actuator<T> {
    public void execute(List<T> rightList, List<ImmutablePair<Integer, Integer>> errorList) throws InterruptedException;
}
