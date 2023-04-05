package com.mbql.easyexcel.handler;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.mbql.easyexcel.anno.ExcelMergeCell;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Field;
import java.util.List;

/**
 * Excel 合并单元格写入处理器
 *
 * @author slp
 */
@Slf4j
@Data
public class ExcelMergeDataWriteHandler implements CellWriteHandler {

    /**
     * 合并列下标数组
     */
    private final int[] mergeColumnIndex;

    /**
     * 合并行下标
     */
    private final int mergeRowIndex;

    /**
     * 是否需要合并单元格
     */
    private final boolean isMergeRow;

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<WriteCellData<?>> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        // 当前行下标
        int curRowIndex = cell.getRowIndex();

        // 当前列下标
        int curColIndex = cell.getColumnIndex();

        Field[] fields = writeSheetHolder.getClazz().getDeclaredFields();

        // 基于注解实现是否需要进行单元格合并
        for (Field field : fields) {
            if (field.isAnnotationPresent(ExcelMergeCell.class)) {
                ExcelMergeCell mergeCell = field.getAnnotation(ExcelMergeCell.class);
                if (curRowIndex > mergeCell.mergeRowIndex()) {
                    if (curColIndex == mergeCell.mergeColumnIndex()) {
                        mergeWithPrevRow(writeSheetHolder, cell, curRowIndex, curColIndex);
                    }
                }
            }
        }

        // 判断是否需要进行单元格合并
        if (isMergeRow && curRowIndex > mergeRowIndex) {
            for (int columnIndex : mergeColumnIndex) {
                if (curColIndex == columnIndex) {
                    mergeWithPrevRow(writeSheetHolder, cell, curRowIndex, curColIndex);
                    break;
                }
            }
        }
    }

    private void mergeWithPrevRow(WriteSheetHolder writeSheetHolder, Cell cell, int curRowIndex, int curColIndex) {
        // 获取当前行的当前列和上一行的当前列数据，通过上一行数据是否相同进行合并
        Object curData = cell.getCellType() == CellType.STRING ? cell.getStringCellValue() : cell.getNumericCellValue();
        Cell preCell = cell.getSheet().getRow(curRowIndex - 1).getCell(curColIndex);
        Object preData = preCell.getCellType() == CellType.STRING ? preCell.getStringCellValue() : preCell.getNumericCellValue();

        // 比较当前行的第一列的单元格与上一行是否相同，相同合并当前单元格与上一行
        if (curData.equals(preData)) {
            Sheet sheet = writeSheetHolder.getSheet();
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            boolean isMerged = false;
            for (int i = 0; i < mergedRegions.size() && !isMerged; i++) {
                CellRangeAddress cellRangeAddress = mergedRegions.get(i);
                // 若上一个单元格已经被合并，则先移出原有的合并单元，再重新添加合并单元
                if (cellRangeAddress.isInRange(curRowIndex - 1, curColIndex)) {
                    sheet.removeMergedRegion(i);
                    cellRangeAddress.setLastRow(curRowIndex);
                    sheet.addMergedRegion(cellRangeAddress);
                    isMerged = true;
                }
            }
            // 若上一个单元格未被合并，则新增合并单元
            if (!isMerged) {
                CellRangeAddress cellRangeAddress = new CellRangeAddress(curRowIndex - 1, curRowIndex, curColIndex, curColIndex);
                sheet.addMergedRegion(cellRangeAddress);
            }
        }
    }

}
