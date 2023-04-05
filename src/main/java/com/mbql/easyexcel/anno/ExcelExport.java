package com.mbql.easyexcel.anno;

import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.handler.WriteHandler;
import com.mbql.easyexcel.handler.ExcelMergeDataWriteHandler;

import java.lang.annotation.*;

/**
 * Excel 导出标识
 *
 * @author slp
 */
@Documented
@Target(ElementType.METHOD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelExport {

    /**
     * 文件名称 不能为空
     *
     * @return name
     */
    String name();

    /**
     * excel文件后缀 默认.xlsx
     *
     * @return suffix
     */
    ExcelTypeEnum suffix() default ExcelTypeEnum.XLSX;

    /**
     * sheet名称
     *
     * @return sheet
     */
    String sheetName() default "";

    /**
     * 是否合并
     *
     * @return 是否合并列
     */
    boolean isMerge() default false;

    /**
     * 要合并的列
     *
     * @return 合并列数组
     */
    int[] mergeColumn() default {};

    /**
     * 表头行数：合并时需计算，不指定默认-1 则取 ExcelProperty 注解 value 的长度
     *
     * @return int
     */
    int headNumber() default -1;

    /**
     * 单元格合并处理器类, 默认 ExcelMergeDataWriteHandler 处理类
     *
     * @return WriteHandler
     */
    Class<? extends WriteHandler> handlerClass() default ExcelMergeDataWriteHandler.class;

}
