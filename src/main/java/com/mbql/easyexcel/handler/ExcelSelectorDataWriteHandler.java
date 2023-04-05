package com.mbql.easyexcel.handler;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.metadata.data.DataFormatData;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.util.StyleUtil;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.mbql.easyexcel.selector.ExcelSelectorResolve;
import lombok.Data;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.util.Map;

/**
 * Excel 下拉选择写入处理器
 *
 * @author slp
 */
@Data
public class ExcelSelectorDataWriteHandler implements SheetWriteHandler, CellWriteHandler {

    /**
     * 批注
     */
    private final Map<Integer, String> notationMap;

    /**
     * 表头列字体颜色
     */
    private final Map<Integer, Short> headColumnMap;

    /**
     * 下拉选数据
     */
    private final Map<Integer, ExcelSelectorResolve> selectedMap;

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        Sheet sheet = writeSheetHolder.getSheet();
        DataValidationHelper helper = sheet.getDataValidationHelper();
        if (CollUtil.isEmpty(selectedMap)) {
            return;
        }
        selectedMap.forEach((k, v) -> {
            // 下拉 首行 末行 首列 末列
            CellRangeAddressList list = new CellRangeAddressList(v.getStartRow(), v.getEndRow(), k, k);
            // 下拉值
            DataValidationConstraint constraint = helper.createExplicitListConstraint(v.getSelectorData());
            DataValidation validation = helper.createValidation(constraint, list);
            validation.setErrorStyle(DataValidation.ErrorStyle.STOP);
            validation.setShowErrorBox(true);
            validation.setSuppressDropDownArrow(true);
            validation.createErrorBox("提示", "请输入下拉选项中的内容");
            sheet.addValidationData(validation);
        });
    }

    @Override
    public void afterCellDispose(CellWriteHandlerContext context) {
        WriteCellData<?> cellData = context.getFirstCellData();
        WriteCellStyle writeCellStyle = cellData.getOrCreateStyle();

        // 单元格设置为文本格式
        DataFormatData dataFormatData = new DataFormatData();
        dataFormatData.setIndex((short) 49);
        writeCellStyle.setDataFormatData(dataFormatData);

        if (Boolean.TRUE.equals(context.getHead())) {
            Cell cell = context.getCell();
            WriteSheetHolder writeSheetHolder = context.getWriteSheetHolder();
            Sheet sheet = writeSheetHolder.getSheet();
            Workbook workbook = writeSheetHolder.getSheet().getWorkbook();
            Drawing<?> drawing = sheet.createDrawingPatriarch();
            // 设置标题字体样式
            WriteFont headWriteFont = new WriteFont();
            // 加粗
            headWriteFont.setBold(true);
            if (CollUtil.isNotEmpty(headColumnMap) && headColumnMap.containsKey(cell.getColumnIndex())) {
                // 设置字体颜色
                headWriteFont.setColor(headColumnMap.get(cell.getColumnIndex()));
            }
            writeCellStyle.setWriteFont(headWriteFont);
            CellStyle cellStyle = StyleUtil.buildCellStyle(workbook, null, writeCellStyle);
            cell.setCellStyle(cellStyle);

            if (CollUtil.isNotEmpty(notationMap) && notationMap.containsKey(cell.getColumnIndex())) {
                // 批注内容
                String notationContext = notationMap.get(cell.getColumnIndex());
                // 创建绘图对象
                Comment comment = drawing.createCellComment(new XSSFClientAnchor(0, 0, 0, 0, (short) cell.getColumnIndex(), 0, (short) 5, 5));
                comment.setString(new XSSFRichTextString(notationContext));
                cell.setCellComment(comment);
            }
        }
    }

}
