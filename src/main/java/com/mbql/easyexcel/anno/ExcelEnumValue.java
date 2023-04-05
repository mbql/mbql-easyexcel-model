package com.mbql.easyexcel.anno;

import java.lang.annotation.*;

/**
 * Excel 枚举值
 *
 * @author slp
 */
@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelEnumValue {

    /**
     * 枚举类型
     *
     * @return class name
     */
    Class<? extends Enum<?>> value();

}
