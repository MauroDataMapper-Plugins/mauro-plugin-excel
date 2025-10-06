package org.maurodata.plugin.exceldatamodel.datarow

import org.maurodata.domain.datamodel.EnumerationValue

import groovy.transform.CompileStatic
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFRow

@CompileStatic
class EnumerationDataRow extends DataRow {

    static final int KEY_COLUMN_INDEX = 8
    static final int VALUE_COLUMN_INDEX = 9

    String key
    String value

    EnumerationDataRow() {
    }

    EnumerationDataRow(EnumerationValue enumerationValue) {
        key = enumerationValue.key
        value = enumerationValue.value
    }

    @Override
    void initialiseRow(Row row) {
        setRow(row)
        key = CellHandlerUtil.getCellValue(row, KEY_COLUMN_INDEX)
        value = CellHandlerUtil.getCellValue(row, VALUE_COLUMN_INDEX)
    }

    @Override
    XSSFRow buildRow(XSSFRow row) {
        addCellToRow row, KEY_COLUMN_INDEX, key
        addCellToRow row, VALUE_COLUMN_INDEX, value
        row
    }

    @Override
    int getFirstMetadataColumn() {
        null
    }
}
