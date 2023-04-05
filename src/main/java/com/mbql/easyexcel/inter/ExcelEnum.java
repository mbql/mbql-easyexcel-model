package com.mbql.easyexcel.inter;

/**
 * Excel 枚举处理顶级接口
 *
 * @author slp
 */
public interface ExcelEnum<C, V> {

    /**
     * 根据 code 获取枚举状态值
     *
     * @param code code
     * @return value
     */
    V getByCode(C code);

    /**
     * 根据 value 获取枚举 code
     *
     * @param value value
     * @return code
     */
    C getByValue(V value);

}
