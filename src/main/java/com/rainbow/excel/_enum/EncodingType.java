package com.rainbow.excel._enum;

import lombok.Getter;

@Getter

public enum EncodingType {
    UTF_8("UTF-8"),
    GBK("GBK");

    String name;

    EncodingType(String name){
        this.name = name;
    }
}
