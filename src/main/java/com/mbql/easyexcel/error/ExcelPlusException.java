package com.mbql.easyexcel.error;

import com.alibaba.excel.exception.ExcelRuntimeException;

/**
 * Excel 业务异常处理类
 * @author slp
 */
public class ExcelPlusException extends ExcelRuntimeException {

    public ExcelPlusException() {
    }

    public ExcelPlusException(String message) {
        super(message);
    }

    public ExcelPlusException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelPlusException(Throwable cause) {
        super(cause);
    }
}
