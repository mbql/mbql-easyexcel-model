package com.mbql.easyexcel.anno;

import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel 作用于标记为必填项
 *
 * @author slp
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelRequired {

    /**
     * 字体颜色
     *
     * @return IndexedColors
     */
    IndexedColors frontColor() default IndexedColors.RED;
}
