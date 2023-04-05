package com.mbql.easyexcel.inter;

/**
 * Excel 下拉选择顶层接口
 *
 * @author slp
 */
public interface ExcelSelectorService {

    /**
     * 获取下拉数据
     *
     * @return java.lang.String[]
     */
    String[] getSelectorData();

    /**
     * 根据字典key获取下拉数据
     *
     * @param dictKeyValue 字典key
     * @return java.lang.String[]
     */
    String[] getSelectorData(String dictKeyValue);
}
