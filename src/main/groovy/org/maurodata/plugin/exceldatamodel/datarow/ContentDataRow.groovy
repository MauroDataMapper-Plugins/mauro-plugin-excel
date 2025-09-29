package org.maurodata.plugin.exceldatamodel.datarow

import org.maurodata.domain.datamodel.DataClass
import org.maurodata.domain.datamodel.DataElement
import org.maurodata.domain.datamodel.DataType
import org.maurodata.domain.datamodel.EnumerationValue
import org.maurodata.domain.model.AdministeredItem
import org.maurodata.domain.model.Item

import groovy.transform.CompileStatic
import org.apache.poi.sl.usermodel.Sheet
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.CellUtil
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet

@CompileStatic
class ContentDataRow extends StandardDataRow<EnumerationDataRow> {

    private static final String DATACLASS_PATH_SPLIT_REGEX = ~/\|/

    String dataClassPath
    String dataElementName
    String description
    String dataTypeName
    String dataTypeDescription
    String referenceToDataClassPath

    int minMultiplicity
    int maxMultiplicity

    ContentDataRow() {
        super(EnumerationDataRow)
    }

    ContentDataRow(AdministeredItem catalogueItem) {
        description = catalogueItem.description

        if (catalogueItem instanceof DataClass) {
            dataClassPath = buildPath(catalogueItem)
            minMultiplicity = catalogueItem.minMultiplicity ?: 0
            maxMultiplicity = catalogueItem.maxMultiplicity ?: 0
        } else if (catalogueItem instanceof DataElement) {
            dataClassPath = buildPath(catalogueItem.dataClass)
            dataElementName = catalogueItem.label
            minMultiplicity = catalogueItem.minMultiplicity ?: 0
            maxMultiplicity = catalogueItem.maxMultiplicity ?: 0

            DataType dataType = catalogueItem.dataType
            dataTypeName = dataType.label
            dataTypeDescription = dataType.description

            if (dataType.isReferenceType()) {
                referenceToDataClassPath = buildPath(dataType.referenceClass)
            } else if (dataType.isEnumerationType()) {
                dataType.enumerationValues.sort {it.key}.each {EnumerationValue enumerationValue ->
                    mergedContentRows << new EnumerationDataRow(enumerationValue)
                }
            }
            // HERE: There is also "Model Type"
        }
    }

    @Override
    void initialiseRow(Row row) {
        setRow(row)
        dataClassPath = CellHandlerUtil.getCellValue(row, 0)
        dataElementName = CellHandlerUtil.getCellValue(row, 1)
        description = CellHandlerUtil.getCellValue(row, 2)
        minMultiplicity = CellHandlerUtil.getCellValueAsInt(row.getCell(3), 1)
        maxMultiplicity = CellHandlerUtil.getCellValueAsInt(row.getCell(4), 1)
        dataTypeName = CellHandlerUtil.getCellValue(row, 5)
        dataTypeDescription = CellHandlerUtil.getCellValue(row, 6)
        referenceToDataClassPath = CellHandlerUtil.getCellValue(row, 7)
        extractMetadataFromColumnIndex()
    }

    @Override
    XSSFRow buildRow(XSSFRow row) {
        addCellToRow row, 0, dataClassPath
        addCellToRow row, 1, dataElementName
        addCellToRow row, 2, description, true
        addCellToRow row, 3, minMultiplicity
        addCellToRow row, 4, maxMultiplicity
        addCellToRow row, 5, dataTypeName
        addCellToRow row, 6, dataTypeDescription, true
        addCellToRow row, 7, referenceToDataClassPath
        addMetadataToRow row

        if (!mergedContentRows) return row

        XSSFRow lastRow = row
        XSSFSheet sheet = row.sheet
        mergedContentRows.size().times {int rowNum ->
            lastRow = CellUtil.getRow(row.rowNum + rowNum, sheet) as XSSFRow
            mergedContentRows[rowNum].buildRow(lastRow)
        }
        if (row.rowNum != lastRow.rowNum) {
            sheet.getRow(METADATA_KEY_ROW_INDEX).lastCellNum.times {int columnIndex ->
                if (columnIndex == 8 || columnIndex == 9) return // Enum columns
                sheet.addMergedRegion(new CellRangeAddress(row.rowNum, lastRow.rowNum, columnIndex, columnIndex))
            }
        }

        row
    }

    @Override
    int getFirstMetadataColumn() {
        10
    }

    List<String> getDataClassPathList() {
        dataClassPath?.split(DATACLASS_PATH_SPLIT_REGEX)?.toList()*.trim()
    }

    List<String> getReferenceToDataClassPathList() {
        referenceToDataClassPath?.split(DATACLASS_PATH_SPLIT_REGEX)?.toList()*.trim()
    }

    private String buildPath(DataClass dataClass) {
        if (!dataClass.parentDataClass) return dataClass.label
        "${buildPath(dataClass.parentDataClass)}|${dataClass.label}"
    }
}
