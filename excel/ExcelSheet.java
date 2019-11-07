package com.lcfc.budget.excel;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.math.BigDecimal;
import java.util.List;
import java.util.Map;

/**
 * excel sheet
 *
 * @author lin.jiale
 * @since 2019-11-1 23:04
 */
public class ExcelSheet {
    ExcelDocument     documentRef;
    XSSFSheet         sheet;
    String            name;
    List<DataColCell> dataColCells = Lists.newArrayList();

    Map<ExcelCol, XSSFCellStyle> cellStyleCache = Maps.newConcurrentMap();

    public ExcelSheet(String name, ExcelDocument document) {
        this.documentRef = document;
        this.name = name;
    }

    /**
     * sheet写入数据
     *
     * @param dataList
     * @param cols
     * @param <T>
     */
    public <T> void write(List<T> dataList, List<ExcelCol> cols) {
        sheet = documentRef.getWorkbook().createSheet(name);
        //提前创建头需要使用的行数量，
        int headerTotalRows = cols.stream().map(t -> t.takeRows()).max((a, b) -> a > b ? 1 : -1).get();
        for (int i = 0; i < headerTotalRows; i++) {
            sheet.createRow(i);
        }

        int colIndex = 0;
        //渲染每个列，下个列的开始列序号为上个列总的列数量+1
        for (ExcelCol col : cols) {
            prepareDrawHeaderCol(col, 0, colIndex);
            renderHeaderCol(col, 0, colIndex);
            colIndex += col.takeCols();
        }

        //渲染数据行
        int dataRowIndex = headerTotalRows;
        Row dataRow;
        Cell dataCell;
        Object dataVal;
        if (dataColCells.size() > 0) {
            for (T t : dataList) {
                dataRow = sheet.createRow(dataRowIndex);
                for (DataColCell fieldCell : dataColCells) {
                    dataCell = dataRow.createCell(fieldCell.colIndex);
                    dataCell.setCellStyle(documentRef.dataCellStyle);
                    //如果数据单元格有设置数据，则直接使用绑定的数据，否则使用当前遍历的数据
                    dataVal = fieldCell.dataCol.value.apply(fieldCell.dataCol.data == null ? t : fieldCell.dataCol.data);
                    if (dataVal != null) {
                        if (dataVal.getClass().equals(BigDecimal.class)
                                || dataVal.getClass().equals(Integer.class)
                                || dataVal.getClass().equals(Double.class)
                                || dataVal.getClass().equals(Float.class)) {
                            dataCell.setCellType(Cell.CELL_TYPE_NUMERIC);
                            if (dataVal.getClass().equals(BigDecimal.class)) {
                                dataCell.setCellValue(((BigDecimal) dataVal).doubleValue());
                            } else {
                                dataCell.setCellValue(Double.valueOf(dataVal.toString()));
                            }
                        } else {
                            dataCell.setCellType(Cell.CELL_TYPE_STRING);
                            dataCell.setCellValue(dataVal.toString());
                        }
                    }
                }
                dataRowIndex++;
            }
        }
    }

    /**
     * 预先创建表格头的行/列。并设置默认样式
     *
     * @param col
     * @param startRowIndex
     * @param startColIndex
     */
    private void prepareDrawHeaderCol(ExcelCol col, int startRowIndex, int startColIndex) {
        //列填充的总行数
        int colDrawRows = col.takeRows();
        //列填充的总列数
        int colDrawCols = col.takeCols();
        int colRenderRows = 0;
        int colRenderCols = 0;
        XSSFCell cell;
        do {
            colRenderCols = 0;
            do {
                cell = sheet.getRow(startRowIndex + colRenderRows).createCell(startColIndex + colRenderCols);
                cell.setCellStyle(documentRef.getHeaderCellStyle());
                colRenderCols++;
            } while (colRenderCols < colDrawCols);
            colRenderRows++;
        } while (colRenderRows < colDrawRows);
    }

    /**
     * 渲染表格头
     *
     * @param col           列
     * @param startRowIndex 起始行号ß
     * @param startColIndex 起始列号
     */
    private void renderHeaderCol(ExcelCol col, int startRowIndex, int startColIndex) {
        XSSFCell cell = sheet.getRow(startRowIndex).getCell(startColIndex);
        cell.setCellValue(new XSSFRichTextString(col.title));

        //合并单元格
        int treeColEndRowIndex = startRowIndex + col.rows - 1;
        int treeColEndColIndex = col.takeCols() + startColIndex - 1;
        CellRangeAddress rangeAddress = new CellRangeAddress(startRowIndex, treeColEndRowIndex, startColIndex, treeColEndColIndex);
        sheet.addMergedRegion(rangeAddress);

        //应用自定义样式
        cell.setCellStyle(getHeaderCustomerStyle(col));

        //检索出dataCol，用于后续列表数据行渲染
        if (col instanceof DataCol) {
            dataColCells.add(new DataColCell(startColIndex, ((DataCol) col)));
            if (col.width > 0) {
                sheet.setColumnWidth(cell.getColumnIndex(), col.width * 256);
            } else {
                sheet.autoSizeColumn(cell.getColumnIndex());
            }
        }

        //树形表头-遍历子节点
        if (col instanceof TreeCol) {
            List<ExcelCol> subCols = ((TreeCol) col).getSubCols();
            if (subCols.size() > 0) {
                startRowIndex += col.rows;
                for (ExcelCol subCol : subCols) {
                    renderHeaderCol(subCol, startRowIndex, startColIndex);
                    startColIndex += subCol.takeCols();
                }
            }
        }
    }

    /**
     * 获取列样式，如果当前列未设置样式，会自动使用父级的样式
     * 减少样式创建， 防止导出的excel样式丢失
     *
     * @param col
     * @return
     */
    private XSSFCellStyle getHeaderCustomerStyle(ExcelCol col) {
        ExcelColStyle style = col.style;
        if (style == null) {
            ExcelCol parent = col._parentRef == null ? null : col._parentRef.get();
            if (parent != null) {
                return getHeaderCustomerStyle(parent);
            }
        }
        if (cellStyleCache.containsKey(col)) {
            return cellStyleCache.get(col);
        } else {
            if (style == null) {
                return documentRef.getHeaderCellStyle();
            }
            XSSFCellStyle cellStyle = documentRef.workbook.createCellStyle();
            cellStyle.cloneStyleFrom(documentRef.getHeaderCellStyle());
            cellStyle.getFont().setColor(style.getFontColor());
            cellStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
            cellStyle.setFillForegroundColor(style.getBackgroundColor());
            cellStyleCache.put(col, cellStyle);
            return cellStyle;
        }
    }

    /**
     * 数据列单元信息
     *
     * @param <T>
     */
    class DataColCell<T> {
        int        colIndex;
        DataCol<T> dataCol;

        public DataColCell(int colIndex, DataCol<T> dataCol) {
            this.colIndex = colIndex;
            this.dataCol = dataCol;
        }
    }
}
