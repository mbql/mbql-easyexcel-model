package com.mbql.easyexcel.res;

import lombok.Data;
import org.apache.poi.ss.formula.functions.T;

import java.io.Serializable;

/**
 * 响应体
 *
 * @author slp
 */
@Data
public class R implements Serializable {

    /**
     * 响应状态码
     */
    private Integer code;

    /**
     * 响应消息
     */
    private String message;

    /**
     * 响应数据
     */
    private T data;

    private R() {
    }

    public static R success(Integer code, T data) {
        return new R(code, null, data);
    }

    public static R success() {
        return R.success(null);
    }

    public static R success(T data) {
        return success(200, data);
    }

    public static R fail(Integer code, String message) {
        return new R(code, message, null);
    }

    public static R fail(String message) {
        return fail(500, message);
    }

    public R(Integer code, String message, T data) {
        this.code = code;
        this.message = message;
        this.data = data;
    }

}
