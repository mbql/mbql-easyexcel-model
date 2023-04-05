package com.mbql.easyexcel.anno;

import com.mbql.easyexcel.inter.ExcelSelectorService;

import java.lang.annotation.*;

/**
 * Excel 标识下拉框选择注解
 *
 * @author slp
 */
@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelSelector {

    /**
     * 固定数据
     *
     * @return String[]
     */
    String[] fixedSelector() default {};

    /**
     * 字典key
     *
     * @return String
     */
    String dictKeyValue() default "";

    /**
     * 服务类
     *
     * @return ExcelSelectorService
     */
    Class<? extends ExcelSelectorService>[] serviceClass() default {};
}
