package com.mbql.easyexcel.error;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;

/**
 * Excel 失败记录对象
 *
 * @author slp
 */
@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class ExcelFailRecord implements Serializable {

    /**
     * 单元格行
     */
    private Integer row;

    /**
     * 单元格列
     */
    private Integer column;

    /**
     * 失败消息
     */
    private String failMessage;

}
