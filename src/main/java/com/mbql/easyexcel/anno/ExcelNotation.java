package com.mbql.easyexcel.anno;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel 批注
 *
 * @author slp
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelNotation {

    /**
     * 文本内容
     *
     * @return string
     */
    String value() default "";
}
