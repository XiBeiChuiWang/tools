package com.rainbow.excel.entity;

import lombok.Getter;
import lombok.Setter;
import org.apache.commons.lang3.tuple.ImmutablePair;

import java.util.ArrayList;
import java.util.List;

@Getter
@Setter
public class ResultList<T>{
    private List<T> rightList = new ArrayList<>();
    private List<ImmutablePair<Integer,Integer>> errorList = new ArrayList<ImmutablePair<Integer,Integer>>();
}
