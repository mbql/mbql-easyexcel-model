package com.mbql.easyexcel.anno;

import java.lang.annotation.*;

/**
 * Excel 合并单元格标识
 *
 * @author slp
 */
@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelMergeCell {

    /**
     * 合并列下标
     *
     * @return int
     */
    int mergeColumnIndex();

    /**
     * 合并起始行下标, 默认从第 2 行开始, 通常第 1 行为表头
     *
     * @return int
     */
    int mergeRowIndex() default 1;

}
