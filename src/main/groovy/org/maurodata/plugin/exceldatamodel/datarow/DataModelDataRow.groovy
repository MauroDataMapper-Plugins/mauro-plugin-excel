package org.maurodata.plugin.exceldatamodel.datarow

import org.maurodata.domain.datamodel.DataModel

import groovy.transform.CompileStatic
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFRow

@CompileStatic
class DataModelDataRow extends StandardDataRow {

    DataModel dataModel
    String sheetKey
    String name
    String description
    String author
    String organisation
    String type

    DataModelDataRow() {
    }

    DataModelDataRow(DataModel dataModel) {
        this.dataModel = dataModel
        String[] words = dataModel.label.split(' ')
        if (words.size() == 1) sheetKey = dataModel.label.toUpperCase()
        else {
            StringBuffer sheetKeyString = new StringBuffer()
            words.each {String word ->
                if (word.length()) sheetKeyString.append(word[0].toUpperCase())
            }
            sheetKey = sheetKeyString.toString()
        }
        name = dataModel.label
        description = dataModel.description
        author = dataModel.author
        organisation = dataModel.organisation
        type = dataModel.dataModelType
    }

    @Override
    void initialiseRow(Row row) {
        setRow(row)
        sheetKey = CellHandlerUtil.getCellValue(row, 0)
        name = CellHandlerUtil.getCellValue(row, 1)
        description = CellHandlerUtil.getCellValue(row, 2)
        author = CellHandlerUtil.getCellValue(row, 3)
        organisation = CellHandlerUtil.getCellValue(row, 4)
        type = CellHandlerUtil.getCellValue(row, 5)
        extractMetadataFromColumnIndex()
    }

    @Override
    XSSFRow buildRow(XSSFRow row) {
        addCellToRow row, 0, sheetKey
        addCellToRow row, 1, name
        addCellToRow row, 2, description, true
        addCellToRow row, 3, author
        addCellToRow row, 4, organisation
        addCellToRow row, 5, type
        addMetadataToRow row
        row
    }

    @Override
    int getFirstMetadataColumn() {
        6
    }
}
