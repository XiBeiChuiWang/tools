package com.rainbow.excel.util;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.write.metadata.WriteSheet;

import com.google.common.collect.Lists;

import com.rainbow.excel._enum.EncodingType;
import com.rainbow.excel.csv.CsvReader;
import com.rainbow.excel.csv.CsvWriter;
import com.rainbow.excel.entity.ReadExcelSheet;
import com.rainbow.excel.entity.RepeatedWrite;
import com.rainbow.excel.entity.ResultList;
import com.rainbow.excel.entity.WriteExcelSheet;
import com.rainbow.excel.handel.impl.ImportToListHandel;
import com.rainbow.excel.handel.impl.ImportToListHandelWithoutModel;
import com.rainbow.excel.listener.ModelExcelListener;
import com.rainbow.excel.listener.NoModelExcelListener;
import lombok.extern.slf4j.Slf4j;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;

@Slf4j
public class EasyExcelTool {

    /**
     * 返回list集合，读的文件只有一个sheet(为第一个sheet)
     *
     * @param inputStream 输入流
     * @param clazz       实体类
     */
    public static ResultList readExcelToList(InputStream inputStream, Class clazz) {
        return readExcelToList(inputStream, 0, clazz);
    }

    /**
     * 读取一个sheet
     *
     * @param inputStream 输入流
     * @param sheet       sheet
     * @param clazz       实体类
     * @return
     */
    public static ResultList readExcelToList(InputStream inputStream, int sheet, Class clazz) {
        if (inputStream == null || clazz == null) {
            throw new RuntimeException("传入参数错误");
        }
        ExcelReader excelReader = null;
        try {
            ResultList objectResultList = new ResultList();
            excelReader = EasyExcel.read(inputStream, clazz, new ModelExcelListener(0, new ImportToListHandel(objectResultList), false)).build();
            ReadSheet readSheet = EasyExcel.readSheet(sheet).build();
            excelReader.read(readSheet);
            return objectResultList;
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (excelReader != null) {
                excelReader.finish();
            }
        }
        return null;
    }

    /**
     * 读取多个sheet
     *
     * @param inputStream 输入流
     * @param clazz       实体类（按照sheet顺序）
     * @return
     */
    public static List<ResultList> readExcelToList(InputStream inputStream, Class... clazz) {
        if (inputStream == null || clazz == null) {
            throw new RuntimeException("传入参数错误");
        }
        ExcelReader excelReader = null;
        try {
            excelReader = EasyExcel.read(inputStream).build();
            ArrayList<ResultList> resultLists = new ArrayList<>();
            for (int i = 0; i < clazz.length; i++) {
                ResultList resultList = new ResultList();
                ReadSheet build = EasyExcel.readSheet(i).head(clazz[i])
                        .registerReadListener(new ModelExcelListener(0, new ImportToListHandel(resultList), false)).build();
                excelReader.read(build);
                resultLists.add(resultList);
            }
            return resultLists;
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (excelReader != null) {
                excelReader.finish();
            }
        }
        return null;
    }

    /**
     * 不需要实体类得到List
     *
     * @param inputStream
     * @return
     */
    public static ResultList readExcelToListWithoutModel(InputStream inputStream) {
        return readExcelToListWithoutModel(inputStream, 0);
    }

    public static ResultList readExcelToListWithoutModel(InputStream inputStream, int sheet) {
        if (inputStream == null) {
            throw new RuntimeException("传入参数错误");
        }
        ExcelReader excelReader = null;
        try {
            ResultList resultList = new ResultList();
            excelReader = EasyExcel.read(inputStream, new NoModelExcelListener(0, new ImportToListHandelWithoutModel(resultList), false)).build();
            ReadSheet readSheet = EasyExcel.readSheet(sheet).build();
            excelReader.read(readSheet);
            return resultList;
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (excelReader != null) {
                excelReader.finish();
            }
        }
        return null;
    }

    /**
     * 读的文件只有一个sheet(为第一个sheet)
     *
     * @param file           excel文件
     * @param readExcelSheet 详见ReadExcelSheet
     */
    public static void readExcel(File file, ReadExcelSheet readExcelSheet) {
        readExcel(file, 0, readExcelSheet);
    }

    /**
     * @param file
     * @param sheet          sheet下标（从0开始）
     * @param readExcelSheet
     */
    public static void readExcel(File file, int sheet, ReadExcelSheet readExcelSheet) {
        checkFile(file);
        if (readExcelSheet == null) {
            throw new RuntimeException("传入参数错误");
        }
        ExcelReader build = null;
        try {
            build = EasyExcel.read(file, readExcelSheet.getClazz(), new ModelExcelListener(readExcelSheet.getOnceCount(), readExcelSheet.getActuator(), readExcelSheet.getFlag())).build();
            ReadSheet readSheet = EasyExcel.readSheet(sheet).build();
            build.read(readSheet);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (build != null) {
                build.finish();
            }
        }
    }


    /**
     * 读的文件有多个sheet
     *
     * @param file
     * @param readExcelSheets list中的每个子元素对应一个sheet
     */
    public static void readExcel(File file, List<ReadExcelSheet> readExcelSheets) throws IOException {
        checkFile(file);
        if (readExcelSheets == null) {
            throw new RuntimeException("传入参数错误");
        }
        FileInputStream fileInputStream = new FileInputStream(file);
        try {
            readExcel(fileInputStream, readExcelSheets);
        } finally {
            if (fileInputStream != null) {
                fileInputStream.close();
            }
        }
    }

    /**
     * @param inputStream    输入流
     * @param readExcelSheet
     */
    public static void readExcel(InputStream inputStream, ReadExcelSheet readExcelSheet) {
        readExcel(inputStream, 0, readExcelSheet);
    }

    /**
     * @param inputStream
     * @param sheet
     * @param readExcelSheet
     */
    public static void readExcel(InputStream inputStream, int sheet, ReadExcelSheet readExcelSheet) {
        if (inputStream == null || readExcelSheet == null) {
            throw new RuntimeException("传入参数错误");
        }
        ExcelReader excelReader = null;
        try {
            excelReader = EasyExcel.read(inputStream, readExcelSheet.getClazz(), new ModelExcelListener(readExcelSheet.getOnceCount(), readExcelSheet.getActuator(), readExcelSheet.getFlag())).build();
            ReadSheet readSheet = EasyExcel.readSheet(sheet).build();
            excelReader.read(readSheet);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (excelReader != null) {
                excelReader.finish();
            }
        }
    }

    /**
     * @param inputStream
     * @param readExcelSheets
     */
    public static void readExcel(InputStream inputStream, List<ReadExcelSheet> readExcelSheets) throws IOException {
        if (inputStream == null || readExcelSheets == null)
            throw new RuntimeException("传入参数错误");
        ExcelReader build = null;
        try {
            build = EasyExcel.read(inputStream).build();
            for (int i = 0; i < readExcelSheets.size(); i++) {
                ReadExcelSheet readExcelSheet = readExcelSheets.get(i);
                ReadSheet readSheet = EasyExcel.readSheet(i).head(readExcelSheet.getClazz())
                        .registerReadListener(new ModelExcelListener(readExcelSheet.getOnceCount(), readExcelSheet.getActuator(), readExcelSheet.getFlag())).build();
                build.read(readSheet);
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (inputStream != null) {
                inputStream.close();
            }
        }
    }


    /**
     * 写出到excel文件（单个sheet）
     *
     * @param outputStream
     * @param writeExcelSheet 详见WriteExcelSheet
     * @return 输出流, 请根据业务转为 xlsx/csv/xls
     */
    public static OutputStream writeExcel(OutputStream outputStream, WriteExcelSheet writeExcelSheet) throws UnsupportedEncodingException {
        ArrayList<WriteExcelSheet> writeExcelSheets = new ArrayList<>();
        writeExcelSheets.add(writeExcelSheet);
        return writeExcel(outputStream, writeExcelSheets);
    }

    /**
     * 写出到excel文件（多个sheet）
     *
     * @param outputStream
     * @param list
     * @return
     */
    public static OutputStream writeExcel(OutputStream outputStream, List<WriteExcelSheet> list) {

        if (outputStream == null || list == null)
            throw new RuntimeException("传入参数错误");
        ExcelWriter write = null;
        try {
            write = EasyExcel.write(outputStream).build();
            for (WriteExcelSheet writeExcelSheet : list) {
                WriteSheet writeSheet = EasyExcel.writerSheet(writeExcelSheet.getSheetName()).head(writeExcelSheet.getClazz()).build();
                write.write(writeExcelSheet.getList(), writeSheet);
            }
        } finally {
            write.finish();
        }
        return outputStream;
    }

    /**
     * @param response        controller传入
     * @param writeExcelSheet
     * @param fileName        文件名
     * @throws IOException
     */
    public static void writeExcel(HttpServletResponse response, WriteExcelSheet writeExcelSheet, String fileName) throws IOException {

        OutputStream outputStream = null;
        OutputStream outputStream1 = null;
        try {
            List<String> strings = Lists.newArrayList("xlsx");
            if (!fileNameIsUse(fileName, strings)) {
                throw new RuntimeException("目前仅支持xlsx与csv两种格式");
            }
            outputStream = response.getOutputStream();

            responseHandle(response, fileName);

            ArrayList<WriteExcelSheet> writeExcelSheets = new ArrayList<>();
            writeExcelSheets.add(writeExcelSheet);

            outputStream1 = writeExcel(outputStream, writeExcelSheets);
            outputStream1.flush();
        } finally {
            if (outputStream != null) {
                outputStream.close();
            }
            if (outputStream1 != null) {
                outputStream1.close();
            }
        }
    }

    /**
     * @param response
     * @param writeExcelSheets
     * @param fileName
     * @throws IOException
     */
    public static void writeExcel(HttpServletResponse response, List<WriteExcelSheet> writeExcelSheets, String fileName) throws IOException {

        OutputStream outputStream = null;
        OutputStream outputStream1 = null;
        try {
            List<String> strings = Lists.newArrayList("xlsx", "csv");
            if (!fileNameIsUse(fileName, strings)) {
                throw new RuntimeException("目前仅支持xlsx与csv两种格式");
            }
            outputStream = response.getOutputStream();

            responseHandle(response, fileName);

            outputStream1 = writeExcel(outputStream, writeExcelSheets);
            outputStream1.flush();
        } finally {
            if (outputStream != null) {
                outputStream.close();
            }
            if (outputStream1 != null) {
                outputStream1.close();
            }
        }
    }


    /**
     * 提供模板（不借助实体类）（1个sheet）
     *
     * @param outputStream 输入流
     * @param sheetName    sheetName
     * @param head
     * @param examples
     * @return
     */
    public static OutputStream writeExcelHead(OutputStream outputStream, String sheetName, List<String> head, List<String>... examples) {
        if (outputStream == null || head == null)
            throw new RuntimeException("传入参数错误");

        ExcelWriter write = null;
        try {
            write = EasyExcel.write(outputStream).build();
            ArrayList<List<String>> lists = new ArrayList<>();
            for (String s : head) {
                ArrayList<String> strings = new ArrayList<>();
                strings.add(s);
                lists.add(strings);
            }
            ArrayList<List<Object>> data = new ArrayList<>();
            for (List list : examples) {
                data.add(list);
            }
            WriteSheet build = EasyExcel.writerSheet(sheetName).head(lists).build();
            write.write(data, build);
        } finally {
            write.finish();
        }
        return outputStream;
    }


    public static OutputStream writeExcelHead(OutputStream outputStream, List<String> sheetName, List<List<String>> head, List<List<List<String>>> examples) {
        if (outputStream == null || head == null || sheetName.size() != head.size())
            throw new RuntimeException("传入参数错误");

        ExcelWriter write = null;
        try {
            write = EasyExcel.write(outputStream).build();
            for (int i = 0; i < sheetName.size(); i++) {
                ArrayList<List<String>> lists = new ArrayList<>();
                for (String s : head.get(i)) {
                    ArrayList<String> strings = new ArrayList<>();
                    strings.add(s);
                    lists.add(strings);
                }
                ArrayList<List<Object>> data = new ArrayList<>();
                for (List list : examples.get(i)) {
                    data.add(list);
                }

                WriteSheet build = EasyExcel.writerSheet(i, sheetName.get(i)).head(lists).build();
                write.write(data, build);
            }

        } finally {
            write.finish();
        }
        return outputStream;
    }

    /**
     * 重复多次写入，使用方式见示例
     *
     * @param outputStream
     * @return
     */
    public static RepeatedWrite repeatedWrite(OutputStream outputStream) {
        if (outputStream == null) {
            throw new RuntimeException("传入参数错误");
        }

        ExcelWriter excelWriter = null;

        excelWriter = EasyExcel.write(outputStream).build();
        RepeatedWrite repeatedWrite = new RepeatedWrite();
        repeatedWrite.setOutputStream(outputStream);
        repeatedWrite.setExcelWriter(excelWriter);
        return repeatedWrite;
    }


    public static RepeatedWrite repeatedWrite(HttpServletResponse response, String fileName) throws IOException {
        responseHandle(response, fileName);
        ServletOutputStream outputStream = response.getOutputStream();
        ExcelWriter excelWriter = null;

        // 这里 指定文件
        excelWriter = EasyExcel.write(outputStream).build();
        RepeatedWrite repeatedWrite = new RepeatedWrite();
        repeatedWrite.setOutputStream(outputStream);
        repeatedWrite.setExcelWriter(excelWriter);
        return repeatedWrite;
    }


    public static void writeExcelHead(HttpServletResponse response, String sheetName, List<String> head, String fileName, List<String>... examples) throws IOException {
        OutputStream outputStream = null;
        OutputStream outputStream1 = null;
        try {
            List<String> strings = Lists.newArrayList("xlsx", "csv");
            if (!fileNameIsUse(fileName, strings)) {
                throw new RuntimeException("目前仅支持xlsx与csv两种格式");
            }
            outputStream = response.getOutputStream();

            responseHandle(response, fileName);

            outputStream1 = writeExcelHead(outputStream, sheetName, head, examples);
            outputStream1.flush();
        } finally {
            if (outputStream != null) {
                outputStream.close();
            }
            if (outputStream1 != null) {
                outputStream1.close();
            }
        }
    }


    public static void writeExcelHead(HttpServletResponse response, List<String> sheetName, List<List<String>> head, List<List<List<String>>> examples, String fileName) throws IOException {
        OutputStream outputStream = null;
        OutputStream outputStream1 = null;
        try {
            List<String> strings = Lists.newArrayList("xlsx", "csv");
            if (!fileNameIsUse(fileName, strings)) {
                throw new RuntimeException("目前仅支持xlsx与csv两种格式");
            }
            outputStream = response.getOutputStream();

            responseHandle(response, fileName);

            outputStream1 = writeExcelHead(outputStream, sheetName, head, examples);
            outputStream1.flush();
        } finally {
            if (outputStream != null) {
                outputStream.close();
            }
            if (outputStream1 != null) {
                outputStream1.close();
            }
        }
    }


    /**
     * 唯一的方式去解析csv（不建议使用，csv为纯文本文件，占用内存且格式容易出现歧义）
     *
     * @param inputStream
     * @param clazz
     * @return
     */
    public static ResultList readCsvToModelList(InputStream inputStream, Class clazz) {
        return readCsvToModelList(inputStream, clazz, EncodingType.GBK);
    }


    public static ResultList readCsvToModelList(InputStream inputStream, Class clazz, EncodingType encodingType) {
        InputStreamReader inputStreamReader = null;
        try {
            inputStreamReader = new InputStreamReader(inputStream, encodingType.getName());
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        return readCsvToModelList(inputStreamReader, clazz);
    }

    /**
     * 唯一的方式去解析csv（不建议使用，csv为纯文本文件，占用内存且格式容易出现歧义）
     *
     * @param
     * @param clazz
     * @return
     */
    public static ResultList readCsvToModelList(Reader reader, Class clazz) {
        CsvReader csvReader = new CsvReader(reader,clazz);
        return csvReader.readModelList();
    }

    /**
     * 返回不带表头行的List集合
     * @param reader
     * @return
     */
    public static List<String[]> readCsvToStringListWithoutHead(Reader reader,Class clazz){
        CsvReader csvReader = new CsvReader(reader,clazz);
        return csvReader.readCsvToStringListWithoutHead();
    }

    public static List<String[]> readCsvToStringListWithoutHead(InputStream inputStream,EncodingType encodingType,Class clazz){
        InputStreamReader reader = null;
        try {
            reader = new InputStreamReader(inputStream,encodingType.getName());
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        return readCsvToStringListWithoutHead(reader,clazz);
    }

    public static List<String[]> readCsvToStringListWithoutHead(InputStream inputStream,Class clazz){
        return readCsvToStringListWithoutHead(inputStream,EncodingType.GBK,clazz);
    }

    /**
     * 返回带表头行的list集合
     * @param reader
     * @return
     */
    public static List<String[]> readCsvToStringListWithHead(Reader reader){
        CsvReader csvReader = new CsvReader(reader);
        return csvReader.readCsvToStringListWithHead();
    }

    public static List<String[]> readCsvToStringListWithHead(InputStream inputStream,EncodingType encodingType){
        InputStreamReader reader = null;
        try {
            reader = new InputStreamReader(inputStream,encodingType.getName());
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        return readCsvToStringListWithHead(reader);
    }

    public static List<String[]> readCsvToStringListWithHead(InputStream inputStream){
        return readCsvToStringListWithHead(inputStream,EncodingType.GBK);
    }

    /**
     * 写csv (Writer)
     *
     * @param writer
     * @param data
     * @param headClass
     * @return
     */
    public static Writer writeCsv(Writer writer, List data, Class headClass) {
        CsvWriter csvWriter = new CsvWriter(writer, headClass);
        return csvWriter.writeCsv(data).build();
    }

    public static void writeCsv(HttpServletResponse response, String fileName, List data, Class headClass) {
        try {
            responseHandle(response, fileName);
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        Writer writer1 = null;
        try {
            PrintWriter writer = response.getWriter();
            writer1 = writeCsv(writer, data, headClass);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (writer1 != null) {
                try {
                    writer1.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

    }

    /**
     * 写csv,默认编码格式为GBK
     *
     * @param outputStream
     * @param data
     * @param headClass
     * @return
     */
    public static Writer writeCsv(OutputStream outputStream, List data, Class headClass) {
        return writeCsv(outputStream, data, headClass, EncodingType.GBK);
    }

    /**
     * 写csv，可任意指定编码格式
     *
     * @param outputStream
     * @param data
     * @param headClass
     * @param encodingType
     * @return
     */
    public static Writer writeCsv(OutputStream outputStream, List data, Class headClass, EncodingType encodingType) {
        OutputStreamWriter outputStreamWriter = null;
        try {
            outputStreamWriter = new OutputStreamWriter(outputStream, encodingType.getName());
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        return writeCsv(outputStreamWriter, data, headClass);
    }

    /**
     * 重复写
     *
     * @param writer
     * @param headClass
     * @return
     */
    public static CsvWriter repeatedWriterCsv(Writer writer, Class headClass) {
        return new CsvWriter(writer, headClass);
    }

    /**
     * 重复写，默认为GBK编码
     *
     * @param outputStream
     * @param headClass
     * @return
     */
    public static CsvWriter repeatedWriterCsv(OutputStream outputStream, Class headClass) {
        return repeatedWriterCsv(outputStream, headClass, EncodingType.GBK);
    }

    /**
     * 指定写，可选择字符编码
     *
     * @param outputStream
     * @param headClass
     * @param encodingType
     * @return
     */
    public static CsvWriter repeatedWriterCsv(OutputStream outputStream, Class headClass, EncodingType encodingType) {
        OutputStreamWriter outputStreamWriter = null;
        try {
            outputStreamWriter = new OutputStreamWriter(outputStream, encodingType.getName());
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        return repeatedWriterCsv(outputStreamWriter, headClass);
    }

    /**
     * 向响应头写
     *
     * @param response
     * @param fileName
     * @param headClass
     * @return
     */
    public static CsvWriter repeatedWriterCsv(HttpServletResponse response, String fileName, Class headClass) {
        try {
            responseHandle(response, fileName);
            PrintWriter writer = response.getWriter();
            return repeatedWriterCsv(writer, headClass);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 不借助实体类，一些比较简单的写操作，只能调用  writeListCsv  方法
     *
     * @param writer
     * @return
     */
    public static CsvWriter repeatedWriterCsvWithoutModel(Writer writer) {
        CsvWriter csvWriter = new CsvWriter(writer);
        return csvWriter;
    }

    public static CsvWriter repeatedWriterCsvWithoutModel(OutputStream outputStream, EncodingType encodingType) {
        try {
            OutputStreamWriter outputStreamWriter = new OutputStreamWriter(outputStream, encodingType.getName());
            return repeatedWriterCsvWithoutModel(outputStreamWriter);
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        return null;
    }

    public static CsvWriter repeatedWriterCsvWithoutModel(OutputStream outputStream) {
        return repeatedWriterCsvWithoutModel(outputStream, EncodingType.GBK);
    }

    public static CsvWriter repeatedWriterCsvWithoutModel(HttpServletResponse response, String fileName) {
        try {
            responseHandle(response, fileName);
            PrintWriter writer = response.getWriter();
            return repeatedWriterCsvWithoutModel(writer);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }


    private static void checkFile(File file) {
        if (file == null)
            throw new RuntimeException("文件为空");
        String[] split = file.getName().split("\\.");
        if (!(split.length == 2 && (split[1].equals("xlsx") || split[1].equals("xls") || split[1].equals("csv")))) {
            throw new RuntimeException("文件格式错误");
        }
    }

    private static boolean fileNameIsUse(String fileName, List<String> list) {
        String[] strings = fileName.split("\\.");
        if (strings.length < 2) {
            return false;
        }

        for (String s : list) {
            if (strings[1].equals(s)) {
                return true;
            }
        }
        return false;
    }

    private static void responseHandle(HttpServletResponse response, String fileName) throws UnsupportedEncodingException {
        fileName = URLEncoder.encode(fileName, "UTF-8");
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf8");
        response.setHeader("Content-Disposition", "attachment; filename=" + fileName);
        response.setHeader("Pragma", "public");
        response.setHeader("Cache-Control", "no-store");
        response.addHeader("Cache-Control", "max-age=0");
    }
}
