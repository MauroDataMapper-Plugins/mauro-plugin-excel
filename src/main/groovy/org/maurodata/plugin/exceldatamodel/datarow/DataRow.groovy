package org.maurodata.plugin.exceldatamodel.datarow

import groovy.transform.CompileStatic
import groovy.util.logging.Slf4j
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellPropertyType
import org.apache.poi.ss.usermodel.RichTextString
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.CellUtil
import org.apache.poi.xssf.usermodel.XSSFRow

import java.time.LocalDate
import java.time.LocalDateTime

@Slf4j
@CompileStatic
abstract class DataRow {

    static final int METADATA_NAMESPACE_ROW_INDEX = 0
    static final int METADATA_KEY_ROW_INDEX = 1

    Row row

    private List<MetadataColumn> metadata = []

    List<MetadataColumn> getMetadata() {
        return metadata
    }

    static Map<String, Set<String>> getMetadataNamespaceAndKeys(List<DataRow> dataRows) {
        List<MetadataColumn> flattenedMetaDataColumns = dataRows.collectMany {DataRow dataRow -> dataRow.metadata}
        Map<String, List<MetadataColumn>> namespaceToMetadataColumns = flattenedMetaDataColumns.groupBy {MetadataColumn metadataColumn -> metadataColumn.namespace}
        Map<String, Set<String>> metadataNamespaceAndKeys = namespaceToMetadataColumns
            .collectEntries {[it.key, it.value.key as Set]}
            .findAll {it.value} as Map<String, Set<String>>
    }

    abstract void initialiseRow(Row row)

    abstract XSSFRow buildRow(XSSFRow row)

    abstract int getFirstMetadataColumn()

    void addCellToRow(Row row, int cellNum, Object content, boolean wrapText = false) {
        Cell cell = row.createCell(cellNum)
        if (content == null) {
            cell.setBlank()
        } else if (content instanceof Boolean) {
            cell.setCellValue((Boolean) content)
        } else if (content instanceof Calendar) {
            cell.setCellValue((Calendar) content)
        } else if (content instanceof Date) {
            cell.setCellValue((Date) content)
        } else if (content instanceof Double) {
            cell.setCellValue((Double) content)
        } else if (content instanceof LocalDate) {
            cell.setCellValue((LocalDate) content)
        } else if (content instanceof LocalDateTime) {
            cell.setCellValue((LocalDateTime) content)
        } else if (content instanceof RichTextString) {
            cell.setCellValue((RichTextString) content)
        } else if (content instanceof String) {
            cell.setCellValue((String) content)
        } else {
            cell.setCellValue(content.toString())
        }

        if (wrapText) {
            CellUtil.setCellStyleProperty(cell, CellPropertyType.WRAP_TEXT, wrapText)
            cell.cellStyle.setShrinkToFit(!wrapText)
        }
    }

    void addMetadataToRow(Row row) {
        if (!metadata) return
        Row namespaceRow = row.sheet.getRow(METADATA_NAMESPACE_ROW_INDEX)
        Row keyRow = row.sheet.getRow(METADATA_KEY_ROW_INDEX)
        metadata.groupBy {it.namespace}.each {String namespace, List<MetadataColumn> metadataColumns ->
            Cell namespaceHeader = namespaceRow.find {Cell cell -> CellHandlerUtil.getCellValue(cell) == namespace} as Cell
            CellRangeAddress mergeRegion = CellHandlerUtil.getMergeRegion(namespaceHeader)
            metadataColumns.each {MetadataColumn metadata ->
                log.debug('Adding {}:{}', metadata.namespace, metadata.key)
                Cell keyHeader = mergeRegion
                    ? CellHandlerUtil.findCell(keyRow, metadata.key, mergeRegion.firstColumn, mergeRegion.lastColumn)
                    : CellHandlerUtil.findCell(keyRow, metadata.key, namespaceHeader.columnIndex, namespaceHeader.columnIndex)
                addCellToRow(row, keyHeader.columnIndex, metadata.value, true)
            }
        }
    }

    void addToMetadata(String namespace, String key, String value) {
        metadata << new MetadataColumn(namespace: namespace, key: key, value: value)
    }

    void extractMetadataFromColumnIndex() {
        Row namespaceRow = row.sheet.getRow(METADATA_NAMESPACE_ROW_INDEX)
        Row keyRow = row.sheet.getRow(METADATA_KEY_ROW_INDEX)
        List<Cell> cells = row.findAll {Cell cell ->
            cell.columnIndex >= firstMetadataColumn &&
            cell.columnIndex <= Math.max(namespaceRow.lastCellNum, keyRow.lastCellNum) &&
            CellHandlerUtil.getCellValue(cell)
        } as List<Cell>
        cells.each {Cell cell ->
            CellRangeAddress mergeRegion = CellHandlerUtil.getMergeRegion(row.sheet, namespaceRow.rowNum, cell.columnIndex)
            String namespaceHeader = CellHandlerUtil.getCellValue(namespaceRow, cell.columnIndex)
            String keyHeader = CellHandlerUtil.getCellValue(keyRow, cell.columnIndex)
            String namespace = keyHeader ? namespaceHeader ?: CellHandlerUtil.getCellValue(namespaceRow, mergeRegion?.firstColumn) : null
            String key = keyHeader ?: namespaceHeader
            metadata << new MetadataColumn(namespace: namespace, key: key, value: CellHandlerUtil.getCellValue(cell))
        }
    }


}
