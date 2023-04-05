package com.mbql.easyexcel.listener;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.text.CharSequenceUtil;
import cn.hutool.core.util.ObjectUtil;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.exception.ExcelDataConvertException;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.read.metadata.holder.ReadRowHolder;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.mbql.easyexcel.error.ExcelFailRecord;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;

import javax.validation.ConstraintViolation;
import javax.validation.Validation;
import javax.validation.Validator;
import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Excel 单元格数据监听器
 *
 * @author slp
 */
@Data
@Slf4j
public class ExcelCellDataListener<T> implements ReadListener<T> {

    /**
     * 数据集合
     */
    private List<T> data = CollUtil.newArrayList();

    /**
     * 记录失败行
     */
    private Map<Integer, ExcelFailRecord> failMap = Maps.newHashMap();

    /**
     * 校验失败记录
     */
    private List<String> errorList = Lists.newArrayList();

    @Override
    public void invoke(T bean, AnalysisContext context) {
        // 数据校验
        dataValidator(bean, context);

        boolean emptyRow = true;
        List<Field> fieldList = getFieldList(bean.getClass());
        for (Field field : fieldList) {
            field.setAccessible(true);
            try {
                Object fieldValue = field.get(bean);
                if (fieldValue instanceof String) {
                    if (CharSequenceUtil.isNotBlank((String) fieldValue)) {
                        emptyRow = false;
                        continue;
                    }
                }
                if (ObjectUtil.isNotNull(fieldValue)) {
                    emptyRow = false;
                }
            } catch (IllegalAccessException e) {
                log.error(e.getMessage(), e);
            }
        }
        if (!emptyRow) {
            // 不处理空数据行
            data.add(bean);
            ReadRowHolder readRowHolder = context.readRowHolder();
            log.info("rowIndex: {}, rowType: {}", readRowHolder.getRowIndex(), readRowHolder.getRowType());
        }
    }

    private void dataValidator(T bean, AnalysisContext context) {
        Validator validator = Validation.buildDefaultValidatorFactory().getValidator();
        Set<ConstraintViolation<T>> validate = validator.validate(bean);
        if (!validate.isEmpty()) {
            StringBuilder sb = new StringBuilder("第" + context.readRowHolder().getRowIndex() + "行数据校验失败：");
            for (ConstraintViolation<T> violation : validate) {
                sb.append(violation.getMessage()).append("; ");
            }
            errorList.add(sb.toString());
        }
    }

    @Override
    public void onException(Exception exception, AnalysisContext context) {
        log.error(exception.getMessage(), exception);
        if (exception instanceof ExcelDataConvertException) {
            ExcelDataConvertException e = (ExcelDataConvertException) exception;
            ExcelFailRecord excelFailRecord = new ExcelFailRecord();
            excelFailRecord.setRow(e.getRowIndex());
            excelFailRecord.setColumn(e.getColumnIndex());
            excelFailRecord.setFailMessage(e.getCause().getMessage());
            failMap.put(e.getRowIndex(), excelFailRecord);
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        log.info("Excel Deal Finish.");
    }

    private List<Field> getFieldList(Class<?> clazz) {
        Field[] fields = clazz.getDeclaredFields();
        return CollUtil.newArrayList(fields);
    }

}
