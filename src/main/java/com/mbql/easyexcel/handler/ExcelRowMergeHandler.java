package com.mbql.easyexcel.handler;

import com.alibaba.excel.write.handler.RowWriteHandler;
import com.alibaba.excel.write.handler.context.RowWriteHandlerContext;
import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Map;

/**
 * Excel 合并策略处理类
 *
 * @author slp
 */
@Data
@AllArgsConstructor
public class ExcelRowMergeHandler implements RowWriteHandler {

    /**
     * 合并的起始行：key：开始，value；结束
     */
    private final Map<Integer, Integer> map;

    /**
     * 要合并的列
     */
    private int[] cols;

    @Override
    public void afterRowDispose(RowWriteHandlerContext context) {
        // 如果是 head 或者 当前行不是合并的起始行，跳过
        if (Boolean.TRUE.equals(context.getHead()) || !map.containsKey(context.getRowIndex())) {
            return;
        }
        Integer endRow = map.get(context.getRowIndex());
        if (!context.getRowIndex().equals(endRow)) {
            // 编列合并的列，合并行
            for (int col : cols) {
                // CellRangeAddress(起始行, 结束行, 起始列, 结束列)
                context.getWriteSheetHolder().getSheet().addMergedRegionUnsafe(new CellRangeAddress(context.getRowIndex(), endRow, col, col));
            }
        }
    }

}
