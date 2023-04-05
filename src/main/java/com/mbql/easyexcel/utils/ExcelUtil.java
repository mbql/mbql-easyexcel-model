package com.mbql.easyexcel.utils;

import cn.hutool.core.text.CharSequenceUtil;
import cn.hutool.core.text.StrPool;
import cn.hutool.core.util.ArrayUtil;
import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.mbql.easyexcel.anno.ExcelNotation;
import com.mbql.easyexcel.anno.ExcelRequired;
import com.mbql.easyexcel.anno.ExcelSelector;
import com.mbql.easyexcel.enums.EasyExcelTypeEnum;
import com.mbql.easyexcel.handler.ExcelMergeDataWriteHandler;
import com.mbql.easyexcel.handler.ExcelSelectorDataWriteHandler;
import com.mbql.easyexcel.listener.ExcelCellDataListener;
import com.mbql.easyexcel.selector.ExcelSelectorResolve;
import com.mbql.easyexcel.strategy.AutoColumnWidthStrategy;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.util.List;
import java.util.Map;

/**
 * Excel 工具类
 *
 * @author slp
 */
@Slf4j
public class ExcelUtil {

    private ExcelUtil() {
    }

    /**
     * 默认的sheet名称
     */
    public static final String DEFAULT_SHEET_NAME = "Sheet1";

    /**
     * 默认合并列下标
     */
    public static final int[] MERGE_COLUMN_INDEX = {};

    /**
     * 默认合并行开始下标
     */
    public static final int MERGE_ROW_INDEX = 0;

    /**
     * 默认是否需要合并单元格
     */
    public static final boolean IS_MERGE_ROW = false;

    /**
     * 写Excel数据
     *
     * @param response response
     * @param fileName 文件名称
     * @param data     数据
     * @param clazz    类class
     * @param <T>      T
     */
    public static <T> void writeExcel(HttpServletResponse response, String fileName, List<T> data, Class<?> clazz) {
        writeExcel(response, fileName, ExcelTypeEnum.XLSX, data, clazz, MERGE_COLUMN_INDEX, MERGE_ROW_INDEX, IS_MERGE_ROW);
    }

    /**
     * 写Excel数据
     *
     * @param response  response
     * @param fileName  文件名称
     * @param sheetName sheet 文件名
     * @param data      数据
     * @param clazz     类class
     * @param <T>       T
     */
    public static <T> void writeExcel(HttpServletResponse response, String fileName, String sheetName, List<T> data, Class<?> clazz) {
        writeExcel(response, fileName, sheetName, ExcelTypeEnum.XLSX, data, clazz, MERGE_COLUMN_INDEX, MERGE_ROW_INDEX, IS_MERGE_ROW);
    }

    /**
     * 写Excel数据
     *
     * @param response         response
     * @param fileName         文件名称
     * @param excelType        文件类型
     * @param data             数据
     * @param clazz            类class
     * @param mergeColumnIndex 合并单元格列下标
     * @param mergeRowIndex    合并开始行
     * @param isMergeRow       是否合并单元格
     * @param <T>              T
     */
    public static <T> void writeExcel(HttpServletResponse response, String fileName, ExcelTypeEnum excelType, List<T> data, Class<?> clazz, int[] mergeColumnIndex, int mergeRowIndex, boolean isMergeRow) {
        writeExcel(response, fileName, DEFAULT_SHEET_NAME, excelType, data, clazz, mergeColumnIndex, mergeRowIndex, isMergeRow);
    }

    /**
     * 写Excel数据
     *
     * @param response         response
     * @param fileName         文件名称
     * @param data             数据
     * @param clazz            类class
     * @param mergeColumnIndex 合并单元格列下标
     * @param mergeRowIndex    合并开始行
     * @param isMergeRow       是否合并单元格
     * @param <T>              T
     */
    public static <T> void writeExcel(HttpServletResponse response, String fileName, List<T> data, Class<?> clazz, int[] mergeColumnIndex, int mergeRowIndex, boolean isMergeRow) {
        writeExcel(response, fileName, ExcelTypeEnum.XLSX, data, clazz, mergeColumnIndex, mergeRowIndex, isMergeRow);
    }

    /**
     * 写Excel数据
     *
     * @param response         response
     * @param fileName         文件名称
     * @param sheetName        sheet名称
     * @param excelType        文件类型
     * @param data             数据
     * @param clazz            类class
     * @param mergeColumnIndex 合并单元格列下标
     * @param mergeRowIndex    合并开始行
     * @param isMergeRow       是否合并单元格
     * @param <T>              T
     */
    public static <T> void writeExcel(HttpServletResponse response, String fileName, String sheetName, ExcelTypeEnum excelType, List<T> data, Class<?> clazz, int[] mergeColumnIndex, int mergeRowIndex, boolean isMergeRow) {
        OutputStream outputStream = null;
        Map<Integer, Short> requiredMap = getRequiredMap(clazz);
        Map<Integer, String> notationMap = getNotationMap(clazz);
        Map<Integer, ExcelSelectorResolve> selectedMap = getSelectedMap(clazz);
        ExcelSelectorDataWriteHandler writeHandler = new ExcelSelectorDataWriteHandler(notationMap, requiredMap, selectedMap);
        ExcelMergeDataWriteHandler mergeDataWriteHandler = new ExcelMergeDataWriteHandler(mergeColumnIndex, mergeRowIndex, isMergeRow);
        try {
            outputStream = getOutputStream(response, fileName, excelType);
            EasyExcelFactory.write(outputStream, clazz).registerWriteHandler(writeHandler).registerWriteHandler(mergeDataWriteHandler).registerWriteHandler(new AutoColumnWidthStrategy()).registerWriteHandler(getStyleStrategy()).excelType(excelType).sheet(sheetName).doWrite(data);
        } catch (Exception e) {
            log.error(e.getMessage(), e);
        } finally {
            if (outputStream != null) {
                try {
                    outputStream.close();
                } catch (IOException e) {
                    log.error(e.getMessage(), e);
                }
            }
        }
    }

    /**
     * 写Excel数据并返回文件流
     *
     * @param sheetName sheet名称
     * @param data      数据
     * @param clazz     类class
     * @param <T>       T
     * @return os
     */
    public static <T> ByteArrayOutputStream writeExcelBaOs(String sheetName, List<T> data, Class<?> clazz) {
        return writeExcelBaOs(sheetName, ExcelTypeEnum.XLSX, data, clazz, MERGE_COLUMN_INDEX, MERGE_ROW_INDEX, IS_MERGE_ROW);
    }

    /**
     * 写Excel数据并返回文件流
     *
     * @param sheetName        sheet名称
     * @param excelType        文件类型
     * @param data             数据
     * @param clazz            类class
     * @param mergeColumnIndex 合并单元格列下标
     * @param mergeRowIndex    合并开始行
     * @param isMergeRow       是否合并单元格
     * @param <T>              T
     * @return os
     */
    public static <T> ByteArrayOutputStream writeExcelBaOs(String sheetName, ExcelTypeEnum excelType, List<T> data, Class<?> clazz, int[] mergeColumnIndex, int mergeRowIndex, boolean isMergeRow) {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        Map<Integer, Short> requiredMap = getRequiredMap(clazz);
        Map<Integer, String> notationMap = getNotationMap(clazz);
        Map<Integer, ExcelSelectorResolve> selectedMap = getSelectedMap(clazz);
        ExcelSelectorDataWriteHandler writeHandler = new ExcelSelectorDataWriteHandler(notationMap, requiredMap, selectedMap);
        ExcelMergeDataWriteHandler mergeDataWriteHandler = new ExcelMergeDataWriteHandler(mergeColumnIndex, mergeRowIndex, isMergeRow);
        try {
            EasyExcelFactory.write(outputStream, clazz).registerWriteHandler(writeHandler).registerWriteHandler(mergeDataWriteHandler).registerWriteHandler(new AutoColumnWidthStrategy()).registerWriteHandler(getStyleStrategy()).excelType(excelType).sheet(sheetName).doWrite(data);
        } catch (Exception e) {
            log.error(e.getMessage(), e);
        } finally {
            try {
                outputStream.close();
            } catch (IOException e) {
                log.error(e.getMessage(), e);
            }
        }
        return outputStream;
    }

    /**
     * 写Excel数据并返回流字节数组
     *
     * @param sheetName sheet名称
     * @param data      数据
     * @param clazz     类class
     * @param <T>       T
     * @return byte[]
     */
    public static <T> byte[] writeExcelBytes(String sheetName, List<T> data, Class<?> clazz) {
        return writeExcelBytes(sheetName, ExcelTypeEnum.XLSX, data, clazz, MERGE_COLUMN_INDEX, MERGE_ROW_INDEX, IS_MERGE_ROW);
    }

    /**
     * 写Excel数据并返回流字节数组
     *
     * @param sheetName        sheet名称
     * @param excelType        文件类型
     * @param data             数据
     * @param clazz            类class
     * @param mergeColumnIndex 合并单元格列下标
     * @param mergeRowIndex    合并开始行
     * @param isMergeRow       是否合并单元格
     * @param <T>              T
     * @return byte[]
     */
    public static <T> byte[] writeExcelBytes(String sheetName, ExcelTypeEnum excelType, List<T> data, Class<?> clazz, int[] mergeColumnIndex, int mergeRowIndex, boolean isMergeRow) {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        Map<Integer, Short> requiredMap = getRequiredMap(clazz);
        Map<Integer, String> notationMap = getNotationMap(clazz);
        Map<Integer, ExcelSelectorResolve> selectedMap = getSelectedMap(clazz);
        ExcelSelectorDataWriteHandler writeHandler = new ExcelSelectorDataWriteHandler(notationMap, requiredMap, selectedMap);
        ExcelMergeDataWriteHandler mergeDataWriteHandler = new ExcelMergeDataWriteHandler(mergeColumnIndex, mergeRowIndex, isMergeRow);
        try {
            EasyExcelFactory.write(outputStream, clazz).registerWriteHandler(writeHandler).registerWriteHandler(mergeDataWriteHandler).registerWriteHandler(new AutoColumnWidthStrategy()).registerWriteHandler(getStyleStrategy()).excelType(excelType).sheet(sheetName).doWrite(data);
        } catch (Exception e) {
            log.error(e.getMessage(), e);
        } finally {
            try {
                outputStream.close();
            } catch (IOException e) {
                log.error(e.getMessage(), e);
            }
        }
        return outputStream.toByteArray();
    }

    /**
     * 读取不包含头信息的Excel
     *
     * @param file  文件
     * @param clazz 类class
     * @param <T>   T
     * @return list
     * @throws IOException io
     */
    public static <T> List<T> readExcelNotContainHeader(MultipartFile file, Class<T> clazz) throws IOException {
        return readExcel(1, file, clazz);
    }

    /**
     * 读取包含头信息的Excel
     *
     * @param file  文件
     * @param clazz 类class
     * @param <T>   T
     * @return list
     * @throws IOException io
     */
    public static <T> List<T> readExcelContainHeader(MultipartFile file, Class<T> clazz) throws IOException {
        return readExcel(0, file, clazz);
    }

    /**
     * 读取Excel
     *
     * @param rowNum 行数
     * @param file   文件
     * @param clazz  类class
     * @param <T>    T
     * @return list
     * @throws IOException io
     */
    public static <T> List<T> readExcel(int rowNum, MultipartFile file, Class<T> clazz) throws IOException {
        String fileName = file.getOriginalFilename();
        InputStream inputStream = file.getInputStream();
        return readExcel(rowNum, fileName, inputStream, clazz);
    }

    /**
     * 读取不包含头信息的Excel
     *
     * @param fileName    文件名称
     * @param inputStream 流
     * @param clazz       类
     * @param <T>         T
     * @return list
     */
    public static <T> List<T> readExcelNotContainHeader(String fileName, InputStream inputStream, Class<T> clazz) {
        return readExcel(1, fileName, inputStream, clazz);
    }

    /**
     * 读取包含头信息的Excel
     *
     * @param fileName    文件名称
     * @param inputStream 流
     * @param clazz       类
     * @param <T>         T
     * @return list
     */
    public static <T> List<T> readExcelContainHeader(String fileName, InputStream inputStream, Class<T> clazz) {
        return readExcel(0, fileName, inputStream, clazz);
    }

    /**
     * 读取Excel
     *
     * @param rowNum      行数
     * @param fileName    文件名称
     * @param inputStream 流
     * @param clazz       类
     * @param <T>         T
     * @return list
     */
    public static <T> List<T> readExcel(int rowNum, String fileName, InputStream inputStream, Class<T> clazz) {
        ExcelCellDataListener<T> dataListener = new ExcelCellDataListener<>();
        try {
            ExcelReader excelReader = getExcelReader(rowNum, fileName, inputStream, clazz, dataListener);
            if (excelReader == null) {
                return Lists.newArrayList();
            }
            List<ReadSheet> sheetList = excelReader.excelExecutor().sheetList();
            for (ReadSheet sheet : sheetList) {
                excelReader.read(sheet);
            }
            excelReader.finish();
        } finally {
            try {
                inputStream.close();
            } catch (IOException e) {
                log.error(e.getMessage(), e);
            }
        }
        return dataListener.getData();
    }

    /**
     * 读取不包含头信息的Excel
     *
     * @param file     文件
     * @param clazz    类class
     * @param listener 监听器
     * @param <T>      T
     * @return listener
     * @throws IOException io
     */
    public static <T> ExcelCellDataListener<T> readExcelNotContainHeader(MultipartFile file, Class<T> clazz, ExcelCellDataListener<T> listener) throws IOException {
        return readExcel(1, file, clazz, listener);
    }

    /**
     * 读取包含头信息的Excel
     *
     * @param file     文件
     * @param clazz    类class
     * @param listener 监听器
     * @param <T>      T
     * @return listener
     * @throws IOException io
     */
    public static <T> ExcelCellDataListener<T> readExcelContainHeader(MultipartFile file, Class<T> clazz, ExcelCellDataListener<T> listener) throws IOException {
        return readExcel(0, file, clazz, listener);
    }

    /**
     * 读取Excel
     *
     * @param rowNum   行数
     * @param file     文件
     * @param clazz    类class
     * @param listener 监听器
     * @param <T>      T
     * @return listener
     * @throws IOException io
     */
    public static <T> ExcelCellDataListener<T> readExcel(int rowNum, MultipartFile file, Class<T> clazz, ExcelCellDataListener<T> listener) throws IOException {
        String fileName = file.getOriginalFilename();
        InputStream inputStream = file.getInputStream();
        return readExcel(rowNum, fileName, inputStream, clazz, listener);
    }

    /**
     * 读取Excel
     *
     * @param rowNum      行数
     * @param fileName    文件名称
     * @param inputStream 流
     * @param clazz       类
     * @param listener    监听器
     * @param <T>         T
     * @return listener
     */
    public static <T> ExcelCellDataListener<T> readExcel(int rowNum, String fileName, InputStream inputStream, Class<T> clazz, ExcelCellDataListener<T> listener) {
        try {
            ExcelReader excelReader = getExcelReader(rowNum, fileName, inputStream, clazz, listener);
            if (excelReader == null) {
                return null;
            }
            List<ReadSheet> sheetList = excelReader.excelExecutor().sheetList();
            for (ReadSheet sheet : sheetList) {
                excelReader.read(sheet);
            }
            excelReader.finish();
        } finally {
            try {
                inputStream.close();
            } catch (IOException e) {
                log.error(e.getMessage(), e);
            }
        }
        return listener;
    }

    /**
     * 获取OutputStream
     *
     * @param response response
     * @param fileName 文件名称
     * @param typeEnum 文件类型
     * @return java.io.OutputStream
     * @throws Exception ex
     */
    private static OutputStream getOutputStream(HttpServletResponse response, String fileName, ExcelTypeEnum typeEnum) throws Exception {
        fileName = URLEncoder.encode(fileName, "UTF-8");
        response.setStatus(200);
        response.setCharacterEncoding("UTF-8");
        if (ExcelTypeEnum.CSV.equals(typeEnum)) {
            response.setContentType("application/csv");
        } else {
            response.setContentType("application/vnd.ms-excel");
        }
        response.setHeader("Content-Disposition", "attachment;filename=" + fileName + typeEnum.getValue());
        return response.getOutputStream();
    }

    /**
     * 获取ExcelReader
     *
     * @param rowNum      行数
     * @param fileName    文件名称
     * @param inputStream 流
     * @param clazz       类class
     * @param listener    监听
     * @return com.alibaba.excel.ExcelReader
     */
    private static ExcelReader getExcelReader(int rowNum, String fileName, InputStream inputStream, Class<?> clazz, ReadListener listener) {
        if (CharSequenceUtil.isBlank(fileName)) {
            return null;
        }
        String fileExtName = getFileExtName(fileName);
        EasyExcelTypeEnum typeEnum = EasyExcelTypeEnum.parseType(fileExtName);
        if (typeEnum == null) {
            log.info("表格类型错误.");
        }
        return EasyExcelFactory.read(inputStream, clazz, listener).headRowNumber(rowNum).build();
    }

    /**
     * 获取文件后缀名称 .xxx
     *
     * @param fileName 文件名称
     * @return java.lang.String
     */
    private static String getFileExtName(String fileName) {
        if (CharSequenceUtil.isBlank(fileName)) {
            return null;
        }
        int lastIndex = fileName.lastIndexOf(StrPool.DOT);
        if (lastIndex != -1) {
            return fileName.substring(lastIndex);
        }
        return null;
    }

    /**
     * 获取样式
     *
     * @return com.alibaba.excel.write.style.HorizontalCellStyleStrategy
     */
    public static HorizontalCellStyleStrategy getStyleStrategy() {
        // 表头样式
        WriteCellStyle headStyle = new WriteCellStyle();
        // 设置表头居中对齐
        headStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        // 内容样式
        WriteCellStyle contentStyle = new WriteCellStyle();
        contentStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        contentStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        contentStyle.setBorderLeft(BorderStyle.THIN);
        contentStyle.setBorderTop(BorderStyle.THIN);
        contentStyle.setBorderRight(BorderStyle.THIN);
        contentStyle.setBorderBottom(BorderStyle.THIN);
        //设置自动换行
        contentStyle.setWrapped(true);
        // 字体策略
        WriteFont contentWriteFont = new WriteFont();
        // 字体大小
        contentWriteFont.setFontHeightInPoints((short) 12);
        contentStyle.setWriteFont(contentWriteFont);
        return new HorizontalCellStyleStrategy(headStyle, contentStyle);
    }

    /**
     * 获取下拉的map
     *
     * @param clazz 类class
     * @return map
     */
    public static Map<Integer, ExcelSelectorResolve> getSelectedMap(Class<?> clazz) {
        Map<Integer, ExcelSelectorResolve> selectedMap = Maps.newHashMap();
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            if (!field.isAnnotationPresent(ExcelSelector.class) || !field.isAnnotationPresent(ExcelProperty.class)) {
                continue;
            }
            ExcelSelector excelSelector = field.getAnnotation(ExcelSelector.class);
            ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
            ExcelSelectorResolve resolve = new ExcelSelectorResolve();
            String[] data = resolve.resolveExcelSelector(excelSelector);
            if (ArrayUtil.isNotEmpty(data)) {
                resolve.setSelectorData(data);
                selectedMap.put(excelProperty.index(), resolve);
            }
        }
        return selectedMap;
    }

    /**
     * 获取必填列Map
     *
     * @param clazz 类class
     * @return map
     */
    public static Map<Integer, Short> getRequiredMap(Class<?> clazz) {
        Map<Integer, Short> requiredMap = Maps.newHashMap();
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            if (!field.isAnnotationPresent(ExcelRequired.class) || !field.isAnnotationPresent(ExcelProperty.class)) {
                continue;
            }
            ExcelRequired excelRequired = field.getAnnotation(ExcelRequired.class);
            ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
            requiredMap.put(excelProperty.index(), excelRequired.frontColor().getIndex());
        }
        return requiredMap;
    }

    /**
     * 获取批注Map
     *
     * @param clazz 类class
     * @return map
     */
    public static Map<Integer, String> getNotationMap(Class<?> clazz) {
        Map<Integer, String> notationMap = Maps.newHashMap();
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            if (!field.isAnnotationPresent(ExcelNotation.class) || !field.isAnnotationPresent(ExcelRequired.class)) {
                continue;
            }
            ExcelNotation excelNotation = field.getAnnotation(ExcelNotation.class);
            ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
            notationMap.put(excelProperty.index(), excelNotation.value());
        }
        return notationMap;
    }

}
