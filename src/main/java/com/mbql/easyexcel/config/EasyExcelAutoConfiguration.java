package com.mbql.easyexcel.config;

import com.mbql.easyexcel.handler.ExcelReturnValueHandler;
import com.mbql.easyexcel.handler.ExcelWriteHandler;
import com.mbql.easyexcel.utils.SpringContextUtil;
import lombok.RequiredArgsConstructor;
import org.springframework.boot.autoconfigure.ImportAutoConfiguration;
import org.springframework.boot.autoconfigure.condition.ConditionalOnMissingBean;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.method.support.HandlerMethodReturnValueHandler;
import org.springframework.web.servlet.mvc.method.annotation.RequestMappingHandlerAdapter;

import javax.annotation.PostConstruct;
import java.util.ArrayList;
import java.util.List;

/**
 * Excel 自动配置类
 *
 * @author slp
 */
@Configuration
@ImportAutoConfiguration(value = { ExcelReturnValueHandler.class, ExcelWriteHandler.class })
@RequiredArgsConstructor
public class EasyExcelAutoConfiguration {

    private final RequestMappingHandlerAdapter requestMappingHandlerAdapter;

    private final ExcelReturnValueHandler excelReturnValueHandler;

    @Bean
    @ConditionalOnMissingBean
    public SpringContextUtil springContextUtil() {
        return new SpringContextUtil();
    }

    /**
     * 追加处理器到 SpringMvc 中
     */
    @PostConstruct
    public void setReturnValueHandlers() {
        List<HandlerMethodReturnValueHandler> returnValueHandlers = requestMappingHandlerAdapter.getReturnValueHandlers();
        List<HandlerMethodReturnValueHandler> newHandlers = new ArrayList<>();
        newHandlers.add(excelReturnValueHandler);
        assert returnValueHandlers != null;
        newHandlers.addAll(returnValueHandlers);
        requestMappingHandlerAdapter.setReturnValueHandlers(newHandlers);
    }

}
