package com.mbql.easyexcel.strategy;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.write.handler.WriteHandler;
import com.google.common.collect.Maps;
import com.mbql.easyexcel.anno.ExcelExport;
import com.mbql.easyexcel.anno.ExcelMergeCell;
import com.mbql.easyexcel.error.ExcelPlusException;
import com.mbql.easyexcel.handler.ExcelMergeDataWriteHandler;
import com.mbql.easyexcel.handler.ExcelRowMergeHandler;
import com.mbql.easyexcel.utils.ExcelUtil;

import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

/**
 * Excel 合并选择处理策略
 *
 * @author slp
 */
public class ExcelMergeSelectorStrategy {

    private ExcelMergeSelectorStrategy() {
    }

    public static <T> WriteHandler findMergeWriteHandler(Class<? extends WriteHandler> handlerClass, ExcelExport excelExport, List<T> returnValue) {
        if (handlerClass.isAssignableFrom(ExcelMergeDataWriteHandler.class)) {
            return new ExcelMergeDataWriteHandler(ExcelUtil.MERGE_COLUMN_INDEX, ExcelUtil.MERGE_ROW_INDEX, ExcelUtil.IS_MERGE_ROW);
        } else if (handlerClass.isAssignableFrom(ExcelRowMergeHandler.class)) {
            return getRowMergeHandler(excelExport, returnValue);
        } else {
            throw new ExcelPlusException("未找到对应的 WriteHandler 处理器");
        }
    }

    private static <T> ExcelRowMergeHandler getRowMergeHandler(ExcelExport excelExport, List<T> returnValue) {
        Class<?> tempClass = returnValue.get(0).getClass();
        List<Field> tempFieldList = new ArrayList<>();
        while (tempClass != null) {
            Collections.addAll(tempFieldList, tempClass.getDeclaredFields());
            // Get the parent class and give it to yourself
            tempClass = tempClass.getSuperclass();
        }
        Field mergeField = tempFieldList.stream().filter(field -> field.getAnnotation(ExcelMergeCell.class) != null).findFirst().orElseThrow(() -> new ExcelPlusException("Excel 合并未找到 ExcelMergeCell 注解"));
        Map<Object, List<T>> mergeMap = returnValue.stream().collect(Collectors.groupingBy(e -> {
            Object value = null;
            try {
                mergeField.setAccessible(true);
                value = mergeField.get(e);
            } catch (IllegalAccessException ex) {
                ex.printStackTrace();
            }
            return value;
        }));

        // 要合并的起始行 = head 的行数
        int index = excelExport.headNumber();

        // index == -1 没有指定起始行
        if (index == -1) {
            ExcelProperty excelProperty = Optional.of(mergeField.getAnnotation(ExcelProperty.class))
                    .orElseThrow(() -> new ExcelPlusException("Excel 合并列未找到 ExcelProperty 注解"));
            index = excelProperty.value().length;
        }

        // 要合并的行，key：起始行，value：结束行
        Map<Integer, Integer> mapMerge = Maps.newHashMap();

        // 分组后数据的顺序会变化，清空数据，计算合并时重新添加
        returnValue.clear();
        for (List<T> vos : mergeMap.values()) {
            returnValue.addAll(vos);
            mapMerge.put(index, index + vos.size() - 1);
            index += vos.size();
        }
        return new ExcelRowMergeHandler(mapMerge, excelExport.mergeColumn());
    }

}
