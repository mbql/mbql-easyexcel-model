package com.mbql.easyexcel.selector;

import cn.hutool.core.text.CharSequenceUtil;
import cn.hutool.core.util.ArrayUtil;
import com.mbql.easyexcel.anno.ExcelSelector;
import com.mbql.easyexcel.inter.ExcelSelectorService;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;

/**
 * Excel 下拉选择处理数据
 *
 * @author slp
 */
@Data
@Slf4j
public class ExcelSelectorResolve {

    /**
     * 下拉选起始行
     */
    private int startRow = 0;

    /**
     * 下拉选结束行
     */
    private int endRow = 500;

    /**
     * 下拉数据集
     */
    private String[] selectorData;

    /**
     * 解决Excel注解的下拉选数据获取
     *
     * @param excelSelector Excel下拉选择
     * @return java.lang.String[]
     */
    public String[] resolveExcelSelector(ExcelSelector excelSelector) {
        if (excelSelector == null) {
            return new String[]{};
        }

        String[] fixedSelector = excelSelector.fixedSelector();
        if (ArrayUtil.isNotEmpty(fixedSelector)) {
            return fixedSelector;
        }
        Class<? extends ExcelSelectorService>[] serviceClass = excelSelector.serviceClass();
        if (ArrayUtil.isNotEmpty(serviceClass)) {
            try {
                ExcelSelectorService excelSelectorService = serviceClass[0].newInstance();
                if (CharSequenceUtil.isBlank(excelSelector.dictKeyValue())) {
                    selectorData = excelSelectorService.getSelectorData();
                } else {
                    selectorData = excelSelectorService.getSelectorData(excelSelector.dictKeyValue());
                }
            } catch (InstantiationException | IllegalAccessException e) {
                log.error(e.getMessage(), e);
            }
        }
        return selectorData;
    }
}
