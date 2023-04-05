package com.mbql.easyexcel.handler;

import cn.hutool.core.text.CharSequenceUtil;
import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.handler.WriteHandler;
import com.mbql.easyexcel.anno.ExcelExport;
import com.mbql.easyexcel.error.ExcelPlusException;
import com.mbql.easyexcel.strategy.AutoColumnWidthStrategy;
import com.mbql.easyexcel.strategy.ExcelMergeSelectorStrategy;
import com.mbql.easyexcel.utils.ExcelUtil;
import lombok.SneakyThrows;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.MediaTypeFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.List;

/**
 * Excel 写入处理器
 *
 * @author slp
 */
public class ExcelWriteHandler {

    @SneakyThrows(IOException.class)
    public <T> void write(List<T> returnValue, ExcelExport excelExport, HttpServletResponse response) {
        // 文件名
        String fileName = String.format("%s%s", URLEncoder.encode(excelExport.name(), "UTF-8"), excelExport.suffix().getValue())
                .replace("\\+", "%20");
        // 根据实际的文件类型找到对应的 contentType
        String contentType = MediaTypeFactory.getMediaType(fileName)
                .map(MediaType::toString)
                .orElse("application/vnd.ms-excel");
        response.setContentType(contentType);
        response.setCharacterEncoding(StandardCharsets.UTF_8.name());
        response.setHeader(HttpHeaders.CONTENT_DISPOSITION, "attachment;filename*=utf-8''" + fileName);

        Class<?> clazz = returnValue.get(0).getClass();
        ExcelWriterSheetBuilder excelWriterSheetBuilder = EasyExcelFactory.write(response.getOutputStream(), clazz)
                .excelType(excelExport.suffix())
                .registerWriteHandler(new ExcelSelectorDataWriteHandler(ExcelUtil.getNotationMap(clazz),
                        ExcelUtil.getRequiredMap(clazz), ExcelUtil.getSelectedMap(clazz)))
                .registerWriteHandler(new AutoColumnWidthStrategy())
                .registerWriteHandler(ExcelUtil.getStyleStrategy())
                .sheet(CharSequenceUtil.isNotBlank(excelExport.sheetName()) ? excelExport.sheetName() : ExcelUtil.DEFAULT_SHEET_NAME);
        if (excelExport.isMerge()) {
            WriteHandler handler = ExcelMergeSelectorStrategy.findMergeWriteHandler(excelExport.handlerClass(), excelExport, returnValue);
            fieldValid(excelExport, handler);
            excelWriterSheetBuilder.registerWriteHandler(handler);
        }
        excelWriterSheetBuilder.doWrite(returnValue);
    }

    private void fieldValid(ExcelExport excelExport, WriteHandler handler) {
        if (handler instanceof ExcelRowMergeHandler) {
            if (excelExport.mergeColumn().length <= 0) {
                throw new ExcelPlusException("Excel合并时，合并的列 mergeColumn 属性不能为空");
            }
        }
    }

}
