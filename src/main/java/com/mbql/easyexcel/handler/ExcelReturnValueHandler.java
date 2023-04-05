package com.mbql.easyexcel.handler;

import com.mbql.easyexcel.anno.ExcelExport;
import lombok.RequiredArgsConstructor;
import org.springframework.core.MethodParameter;
import org.springframework.util.Assert;
import org.springframework.web.context.request.NativeWebRequest;
import org.springframework.web.method.support.HandlerMethodReturnValueHandler;
import org.springframework.web.method.support.ModelAndViewContainer;

import javax.servlet.http.HttpServletResponse;
import java.util.List;

/**
 * Excel 返回值处理器
 *
 * @author slp
 */
@RequiredArgsConstructor
public class ExcelReturnValueHandler implements HandlerMethodReturnValueHandler {

    private final ExcelWriteHandler excelWriteHandler;

    @Override
    public boolean supportsReturnType(MethodParameter returnType) {
        // 判断方法是否标注 ExcelExport 注解
        return returnType.getMethodAnnotation(ExcelExport.class) != null;
    }

    @Override
    public void handleReturnValue(Object returnValue, MethodParameter returnType, ModelAndViewContainer mavContainer, NativeWebRequest webRequest) {
        HttpServletResponse response = webRequest.getNativeResponse(HttpServletResponse.class);
        Assert.notNull(response, "Excel 导出 HttpServletResponse 为空");
        ExcelExport excelExport = returnType.getMethodAnnotation(ExcelExport.class);
        Assert.notNull(excelExport, "Excel 导出 ExcelExport 为空");
        mavContainer.setRequestHandled(true);
        // 判断返回值是否是list类型
        if ((returnValue instanceof List)) {
            // List 不为空，并且其中元素不是 list 处理 Excel
            List<?> objList = (List<?>) returnValue;
            if (!objList.isEmpty() && !(objList.get(0) instanceof List)) {
                excelWriteHandler.write(objList, excelExport, response);
            }
        }
    }
}
