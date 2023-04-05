package com.mbql.easyexcel.enums;

import lombok.Getter;

import java.util.Arrays;

/**
 * Excel 类型格式枚举
 *
 * @author slp
 */
public enum EasyExcelTypeEnum {

    /**
     * csv
     */
    CSV(0, ".csv"),

    /**
     * xls
     */
    XLS(1, ".xls"),

    /**
     * xlsx
     */
    XLSX(2, ".xlsx");

    @Getter
    private final Integer type;

    @Getter
    private final String name;


    EasyExcelTypeEnum(Integer type, String name) {
        this.type = type;
        this.name = name;
    }

    public static EasyExcelTypeEnum parseType(String fileExtName) {
        return Arrays.stream(values()).filter(f -> fileExtName.equals(f.name)).findFirst().orElse(null);
    }

}
