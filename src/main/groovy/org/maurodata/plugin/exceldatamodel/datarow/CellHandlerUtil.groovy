package org.maurodata.plugin.exceldatamodel.datarow

import groovy.transform.CompileStatic
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress

@CompileStatic
class CellHandlerUtil {

    private static final DataFormatter dataFormatter = new DataFormatter()

    static String getCellValue(final Cell cell) {
        cell ? dataFormatter.formatCellValue(cell).replaceAll(/’/, '\'').replaceAll(/—/, '-').trim() : ''
    }

    static String getCellValue(final Row row, final Integer column) {
        row && column != null ? getCellValue(row.getCell(column)) : ''
    }

    static CellRangeAddress getMergeRegion(final Cell cell) {
        getMergeRegion(cell?.sheet, cell.rowIndex, cell.columnIndex)
    }

    static CellRangeAddress getMergeRegion(final Sheet sheet, final int rowIndex, final int columnIndex) {
        sheet?.mergedRegions?.find {it.isInRange(rowIndex, columnIndex)}
    }

    static Cell findCell(final Row row, final String content, final int firstColumn, final int lastColumn) {
        row.find {Cell cell -> cell.columnIndex >= firstColumn && cell.columnIndex <= lastColumn && getCellValue(cell) == content} as Cell
    }

    static int getCellValueAsInt(Cell cell, int defaultValue) {
        String cellValue = getCellValue(cell)
        cellValue ? cellValue.toInteger() : defaultValue
    }
}