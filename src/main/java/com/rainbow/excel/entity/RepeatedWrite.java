package com.rainbow.excel.entity;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import lombok.Getter;
import lombok.Setter;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.locks.ReentrantLock;

@Getter
@Setter
public class RepeatedWrite {

    private OutputStream outputStream;

    private ExcelWriter excelWriter;

    private ConcurrentHashMap<Integer, WriteSheet> sheets = new ConcurrentHashMap<>();

    private ReentrantLock lock = new ReentrantLock();

    public RepeatedWrite firstWrite(List data,int sheetNo,Class headClass,String sheetName){
        if (sheetNo < 0 || headClass == null || sheetName == null){
            throw new RuntimeException("传入参数错误");
        }

        WriteSheet build = EasyExcel.writerSheet(sheetNo, sheetName).head(headClass).build();
        ExcelWriter write = excelWriter.write(data, build);

        sheets.put(sheetNo,build);

        return this;
    }

    public synchronized RepeatedWrite repeatedWrite(List data,int sheetNo){
        if (sheetNo < 0)
            throw new RuntimeException("传入参数错误");
//        if (data.size() == 0)
//            return this;
        Object o = data.get(0);

        WriteSheet writeSheet = sheets.get(sheetNo);
        if (writeSheet == null){
            throw new RuntimeException("writeSheet不能为空");
        }

        try {
            lock.lock();
            excelWriter.write(data,writeSheet);
        } finally {
            lock.unlock();
        }

        return this;
    }


    public OutputStream finish(){
        try {
            outputStream.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }
        if (excelWriter != null){
            excelWriter.finish();
        }
        return outputStream;
    }

}
