package org.maurodata.plugin.exceldatamodel.workbook

import org.maurodata.domain.datamodel.DataClass
import org.maurodata.domain.datamodel.DataElement
import org.maurodata.domain.datamodel.DataModel
import org.maurodata.domain.datamodel.DataModelType
import org.maurodata.domain.datamodel.DataType
import org.maurodata.domain.datamodel.EnumerationValue
import org.maurodata.domain.facet.Metadata
import org.maurodata.domain.model.AdministeredItem
import org.maurodata.plugin.exceldatamodel.datarow.CellHandlerUtil
import org.maurodata.plugin.exceldatamodel.datarow.ContentDataRow
import org.maurodata.plugin.exceldatamodel.datarow.DataModelDataRow
import org.maurodata.plugin.exceldatamodel.datarow.DataRow
import org.maurodata.plugin.exceldatamodel.datarow.EnumerationDataRow
import org.maurodata.plugin.exceldatamodel.datarow.MetadataColumn
import org.maurodata.plugin.exceldatamodel.datarow.StandardDataRow

import groovy.transform.CompileStatic
import groovy.util.logging.Slf4j
import org.apache.poi.EncryptedDocumentException
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellPropertyType
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.CellUtil
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder

import java.time.Duration

@Slf4j
@CompileStatic
class Utils {

    static final String DATAMODEL_SHEETNAME = 'DataModels'
    static final String CONTENT_TEMPLATE_SHEETNAME = 'KEY_1'

    static final int DATAMODELS_NUM_HEADER_ROWS = 2
    static final int CONTENT_NUM_HEADER_ROWS = 2
    static final int DATAMODELS_ID_COLUMN_INDEX = 1
    static final int CONTENT_ID_COLUMN_INDEX = 0

    static final Map<String, String> DATAMODEL_SHEET_COLUMNS = [name        : 'Name',
                                                                description : 'Description',
                                                                author      : 'Author',
                                                                organisation: 'Organisation',
                                                                sheetKey    : 'Sheet Key',
                                                                type        : 'Type']
    static final Map<String, String> ENUM_SHEET_COLUMNS = [dataModelName  : 'DataModel Name',
                                                           enumerationName: 'Enumeration Name',
                                                           description    : 'Description',
                                                           key            : 'Key',
                                                           value          : 'Value']
    static final Map<String, String> MODEL_SHEET_COLUMNS = [dataClassPath    : 'DataClass Path',
                                                            dataElementName  : 'DataElement Name',
                                                            description      : 'Description',
                                                            minMultiplicity  : 'Minimum[ \r\n]+Multiplicity',
                                                            maxMultiplicity  : 'Maximum[ \r\n]+Multiplicity',
                                                            dataTypeName     : 'DataType Name',
                                                            dataTypeReference: 'DataType Reference']

    static DataFormatter dataFormatter = new DataFormatter()

    static XSSFWorkbook loadTemplateWorkbook(final String resourceName) {
        try (final InputStream is = Utils.class.classLoader.getResourceAsStream(resourceName)) {
            if (is == null) {throw new RuntimeException("EFS01: Could not find ${resourceName} on classpath")}
            WorkbookFactory.create(is) as XSSFWorkbook
        } catch (final EncryptedDocumentException ignored) {
            throw new RuntimeException("EFS02: Excel file ${resourceName} could not be read as it is encrypted")
        } catch (final IOException ex) {
            throw new RuntimeException("EFS04: Excel file ${resourceName} could not be read", ex)
        }
    }

    static void loadDataModelsIntoWorkbook(final Collection<DataModel> dataModels, final XSSFWorkbook workbook) {

        final WorkbookCellAndBorderStyles workbookCellAndBorderStyles = new WorkbookCellAndBorderStyles(workbook)

        final List<DataModelDataRow> dataRows = dataModels.collect {final DataModel dataModel ->
            final DataModelDataRow dmdr = new DataModelDataRow(dataModel)
            dataModel.metadata.each {final Metadata md ->
                dmdr.addToMetadata(md.namespace, md.key, md.value)
            }
            dmdr
        }

        final XSSFSheet dataModelSheet = workbook.getSheet(DATAMODEL_SHEETNAME)
        addMetadataHeadersToSheet(dataModelSheet, dataRows)

        dataRows.each {final DataModelDataRow dataRow ->
            loadDataModelDataRowToSheet(dataModelSheet, dataRow, workbookCellAndBorderStyles)
            loadContentDataRowsToSheet(dataRow, workbook, workbookCellAndBorderStyles)
        }
        removeTemplateSheet(workbook)
    }

    static <K extends StandardDataRow> void addMetadataHeadersToSheet(final XSSFSheet sheet, final List<K> dataRows, final int cellStyleColumn = 0) {
        final Map<String, Set<String>> metadataMapping = DataRow.getMetadataNamespaceAndKeys(dataRows as List<DataRow>)
        final Row firstRow = sheet.getRow(0)
        final Row secondRow = sheet.getRow(1)
        int namespaceColumnIndex = firstRow.lastCellNum

        metadataMapping.sort().each {final String namespace, final Set<String> keys ->
            final Cell namespaceHeader = buildHeaderCell(firstRow, namespaceColumnIndex, cellStyleColumn, namespace)
            if (keys.size() == 1) {
                buildHeaderCell(secondRow, namespaceHeader.columnIndex, cellStyleColumn, keys.first())
                namespaceColumnIndex++
            } else {
                sheet.addMergedRegion(new CellRangeAddress(
                    DataRow.METADATA_NAMESPACE_ROW_INDEX, DataRow.METADATA_NAMESPACE_ROW_INDEX,
                    namespaceHeader.columnIndex, namespaceHeader.columnIndex + keys.size() - 1))
                keys.eachWithIndex {final String key, final int i ->
                    buildHeaderCell(secondRow, namespaceHeader.columnIndex + i, cellStyleColumn, key)
                }
                namespaceColumnIndex += keys.size()
            }
        }
    }

    static Cell buildHeaderCell(final Row headerRow, final int columnIndex, final int cellStyleColumn, final String cellValue, final boolean createMergeRegion = false) {
        final Cell headerCell = headerRow.createCell(columnIndex).tap {
            setCellValue cellValue
        }

        final Cell copyCell = headerRow.getCell(cellStyleColumn)
        final CellRangeAddress mergedRegion = CellHandlerUtil.getMergeRegion(copyCell)

        if (!createMergeRegion || !mergedRegion) {
            headerCell.setCellStyle(copyCell.cellStyle)
            return headerCell
        }

        final CellRangeAddress newRegion = mergedRegion.copy().tap {
            firstColumn = headerCell.columnIndex
            lastColumn = headerCell.columnIndex
        }
        (newRegion.firstRow..newRegion.lastRow).each {final int rowIndex ->
            final Row newRegionRow = headerRow.sheet.getRow(rowIndex)
            final Cell newRegionHeaderCell = CellUtil.getCell(newRegionRow, headerCell.columnIndex)
            newRegionHeaderCell.setCellStyle(copyCell.cellStyle)
            if (rowIndex == newRegion.lastRow) {
                CellUtil.setCellStyleProperty(newRegionHeaderCell, CellPropertyType.BORDER_BOTTOM,
                                              newRegionRow.getCell(cellStyleColumn).cellStyle.borderBottom)
            }
        }
        headerRow.sheet.addMergedRegion(newRegion)
        headerCell
    }

    static void loadDataModelDataRowToSheet(final XSSFSheet sheet, final DataModelDataRow dataRow, final WorkbookCellAndBorderStyles workbookCellAndBorderStyles) {
        log.debug('Adding row to DataModel [{}] sheet', sheet.sheetName)
        dataRow.buildRow(sheet.createRow(sheet.lastRowNum + 1))
        configureDataModelSheetStyle(sheet, dataRow, workbookCellAndBorderStyles)
    }

    static void configureDataModelSheetStyle(final XSSFSheet sheet, final DataModelDataRow dataRow, final WorkbookCellAndBorderStyles workbookCellAndBorderStyles) {
        log.debug('Configuring DataModel [{}] sheet style', sheet.sheetName)
        configureSheetStyle(sheet, dataRow.firstMetadataColumn, DATAMODELS_NUM_HEADER_ROWS, workbookCellAndBorderStyles)
    }

    static void configureContentSheetStyle(final XSSFSheet sheet, final List<ContentDataRow> dataRows, final WorkbookCellAndBorderStyles workbookCellAndBorderStyles) {
        log.debug('Configuring DataModel [{}] content sheet style', sheet.sheetName)
        configureSheetStyle(sheet, dataRows.first().firstMetadataColumn, CONTENT_NUM_HEADER_ROWS, workbookCellAndBorderStyles)
    }

    static XSSFSheet configureSheetStyle(final XSSFSheet sheet, final int metadataColumnIndex, final int numberOfHeaderRows,
                                         final WorkbookCellAndBorderStyles workbookCellAndBorderStyles) {
        boolean isRowColoured = false
        final int lastColumnIndex = sheet.max {it.lastCellNum}.lastCellNum

        for (int rowIndex = numberOfHeaderRows; rowIndex <= sheet.lastRowNum; rowIndex++) {
            XSSFRow row = sheet.getRow(rowIndex)
            if (!row) continue

            setRowStyle(row, isRowColoured, metadataColumnIndex, 0, lastColumnIndex, workbookCellAndBorderStyles)

            CellRangeAddress enumRowsRegion = CellHandlerUtil.getMergeRegion(row.getCell(0))
            if (enumRowsRegion) {
                boolean isEnumRowColoured = isRowColoured
                (enumRowsRegion.firstRow..enumRowsRegion.lastRow).each {final int enumRowIndex ->
                    final Closure setEnumRowStyle = {final int start, final int finish ->
                        setRowStyle(sheet.getRow(enumRowIndex), isEnumRowColoured, metadataColumnIndex, start, finish, workbookCellAndBorderStyles)
                    }
                    setEnumRowStyle(0, EnumerationDataRow.KEY_COLUMN_INDEX)
                    setEnumRowStyle(EnumerationDataRow.KEY_COLUMN_INDEX, EnumerationDataRow.VALUE_COLUMN_INDEX + 1)
                    setEnumRowStyle(EnumerationDataRow.VALUE_COLUMN_INDEX + 1, lastColumnIndex)
                    isEnumRowColoured = !isEnumRowColoured
                }
                rowIndex = enumRowsRegion.lastRow
            } else {
                rowIndex = row.rowNum
            }

            isRowColoured = !isRowColoured
        }

        final List<Integer> columnSizes = sheet.sheetName == DATAMODEL_SHEETNAME ? [3, 4, 5] : [5, 7, 9]
        autoSizeColumns(sheet, [0, 1] + columnSizes)
        autoSizeHeaderColumnsAfter(sheet, metadataColumnIndex)
        sheet
    }

    static void setRowStyle(final XSSFRow row, final boolean colourRow, final int metadataColumnIndex, final int start, final int finish,
                            final WorkbookCellAndBorderStyles workbookCellAndBorderStyles) {
        final XSSFCellStyle cellStyle = colourRow ? workbookCellAndBorderStyles.colouredCellStyle : workbookCellAndBorderStyles.defaultCellStyle
        final XSSFCellStyle borderStyle = colourRow ? workbookCellAndBorderStyles.metadataColouredBorderStyle : workbookCellAndBorderStyles.metadataBorderStyle
        final CellStyle workbookDefaultStyle = row.sheet.workbook.getCellStyleAt(0)
        (start..<finish).each {final int cellIndex ->
            Cell cell = row.getCell(cellIndex)
            final CellStyle styleToUse = cellIndex == metadataColumnIndex - 1 ? borderStyle : cellStyle
            if (!cell) {
                cell = CellUtil.getCell(row, cellIndex)
                cell.setCellStyle(styleToUse)
            } else if (!cell.cellStyle || cell.cellStyle == workbookDefaultStyle) {
                cell.setCellStyle(styleToUse)
            } else if (cell.cellStyle.wrapText) {
                copyCellStyleInto(cell, styleToUse)
            }
        }
    }

    static void copyCellStyleInto(final Cell cell, final XSSFCellStyle style) {
        if (style.getFillPattern() == FillPatternType.SOLID_FOREGROUND) {
            CellUtil.setCellStyleProperty(cell, CellPropertyType.FILL_PATTERN, FillPatternType.SOLID_FOREGROUND)
            (cell.cellStyle as XSSFCellStyle).setFillForegroundColor(style.fillForegroundColorColor)
        }
        final XSSFColor borderColour = style.leftBorderXSSFColor
        (cell.cellStyle as XSSFCellStyle).tap {
            setBorderTop BorderStyle.THIN
            setBorderLeft BorderStyle.THIN
            setBorderBottom BorderStyle.THIN
            setBorderRight BorderStyle.THIN
            setBorderColor XSSFCellBorder.BorderSide.TOP, borderColour
            setBorderColor XSSFCellBorder.BorderSide.LEFT, borderColour
            setBorderColor XSSFCellBorder.BorderSide.BOTTOM, borderColour
            setBorderColor XSSFCellBorder.BorderSide.RIGHT, borderColour
        }
    }

    static void loadContentDataRowsToSheet(final DataModelDataRow dataRow, final XSSFWorkbook workbook, final WorkbookCellAndBorderStyles workbookCellAndBorderStyles) {
        log.debug('Adding content rows to DataModel [{}] sheet', dataRow.sheetKey)
        final List<ContentDataRow> contentDataRows = createContentDataRows(dataRow.dataModel.id, dataRow.dataModel.childDataClasses.sort {it.label})
        if (!contentDataRows) return
        final XSSFSheet contentSheet = loadContentSheet(dataRow, workbook)
        loadDataRowsIntoSheet(contentSheet, contentDataRows)
        configureContentSheetStyle(contentSheet, contentDataRows, workbookCellAndBorderStyles)
    }

    static <K extends AdministeredItem> List<ContentDataRow> createContentDataRows(final UUID dataModelId, final List<K> catalogueItems,
                                                                                   final List<ContentDataRow> dataRows = []) {
        log.debug('Creating content rows')
        catalogueItems.each {final AdministeredItem catalogueItem ->
            log.trace('Creating content {} : {}', catalogueItem.domainType, catalogueItem.label)
            final ContentDataRow cdr = new ContentDataRow(catalogueItem)
            // HERE: be wary of DTOs that hold Metadata
            catalogueItem.metadata.each {final Metadata md ->
                cdr.addToMetadata(md.namespace, md.key, md.value)
            }
            dataRows << cdr
            if (catalogueItem instanceof DataClass) {
                final DataClass dataClass = (DataClass) catalogueItem
                createContentDataRows(dataModelId, dataClass.dataElements, dataRows)
                createContentDataRows(dataModelId, dataClass.dataClasses, dataRows)
            }
        }
        dataRows
    }

    static XSSFSheet loadContentSheet(final DataModelDataRow dataRow, final XSSFWorkbook workbook) {
        log.debug('Loading DataModel [{}] content sheet', dataRow.name)
        final XSSFSheet contentSheet = createContentSheetFromTemplate(dataRow.sheetKey, workbook)
        dataRow.sheetKey = contentSheet.sheetName
        return contentSheet
    }

    static XSSFSheet createContentSheetFromTemplate(final String name, final XSSFWorkbook workbook) {

        final int sheetIndex = workbook.getSheetIndex(CONTENT_TEMPLATE_SHEETNAME)
        if (sheetIndex == -1) {throw new RuntimeException("Missing content sheet template")}

        String chosenName = name
        int existingWithName = workbook.getSheetIndex(chosenName)
        do {
            if (existingWithName >= 0) {
                chosenName = chosenName + '.1'
            }
            existingWithName = workbook.getSheetIndex(chosenName)
        } while (existingWithName >= 0)

        return workbook.cloneSheet(sheetIndex, chosenName)
    }

    static void loadDataRowsIntoSheet(final XSSFSheet sheet, final List<ContentDataRow> dataRows) {
        log.debug('Loading DataRows into sheet')
        addMetadataHeadersToSheet(sheet, dataRows, 3)
        dataRows.each {final ContentDataRow dataRow ->
            log.debug('Adding to row {}', sheet.lastRowNum + 1)
            dataRow.buildRow(sheet.createRow(sheet.lastRowNum + 1))
        }
    }

    static void autoSizeColumns(final XSSFSheet sheet, final List<Integer> columns) {
        columns.each {final Integer it -> sheet.autoSizeColumn(it, true)}
    }

    static void autoSizeHeaderColumnsAfter(final XSSFSheet sheet, final int columnIndex) {
        sheet.getRow(0)
            .findAll {final Cell cell -> cell.columnIndex >= columnIndex && CellHandlerUtil.getCellValue(cell)}
            .each {sheet.autoSizeColumn((it as Cell).columnIndex, true)}
    }

    static void removeTemplateSheet(final XSSFWorkbook workbook) {
        log.debug('Removing the template sheet')
        workbook.removeSheetAt(workbook.getSheetIndex(CONTENT_TEMPLATE_SHEETNAME))
        workbook.setActiveSheet(0)
    }


    static void simpleLoadDataModelsIntoWorkbook(final Collection<DataModel> dataModels, final XSSFWorkbook workbook) {

        final XSSFSheet dataModelsSheet = workbook.createSheet('DataModels')
        final XSSFSheet enumerationsSheet = workbook.createSheet('Enumerations')

        final List<Map<String, String>> dataModelsSheetArray = []
        final List<Map<String, String>> enumerationsSheetArray = []

        dataModels.each {final DataModel dataModel ->
            final String sheetKey = createSheetKey(dataModel.label)
            // Create the sheet for the DM data
            final XSSFSheet dataModelSheet = workbook.createSheet(sheetKey)

            // Add row to the DM sheet
            dataModelsSheetArray.add(buildDataModelRow(dataModel, sheetKey))

            // Add the enumerations to the enumeration sheet
            enumerationsSheetArray.addAll(buildEnumerationDataRows(dataModel.label, dataModel.enumerationValues))

            final List<Map<String, String>> dataModelSheetArray = []

            dataModel.childDataClasses.each {final DataClass dataClass ->
                dataModelSheetArray.addAll(buildDataClassRows(dataClass, null))
            }
            writeArrayToSheet(workbook, dataModelSheet, dataModelSheetArray)
        }

        writeArrayToSheet(workbook, dataModelsSheet, dataModelsSheetArray)
        writeArrayToSheet(workbook, enumerationsSheet, enumerationsSheetArray)

    }

    static String createSheetKey(final String dataModelName) {

        final String[] words = dataModelName.split(' ')
        if (words.size() == 1) {
            return dataModelName.toUpperCase()
        } else {
            final StringBuilder sheetKey = new StringBuilder(dataModelName.length())

            words.each {final String word ->
                if (word.length() > 0) {
                    final String newWord = word.replaceAll('^[^A-Za-z]*', '')

                    sheetKey.append(newWord[0].toUpperCase() + newWord.replaceAll('[^0-9]', ''))
                }
            }
            return sheetKey.toString()
        }
    }

    static Map<String, String> buildDataModelRow(final DataModel dataModel, final String sheetKey) {
        final Map<String, String> dataRow = [
            Name        : dataModel.label,
            Description : dataModel.description,
            Author      : dataModel.author,
            Organisation: dataModel.organisation,
            'Sheet Key' : sheetKey,
            Type        : dataModel.dataModelType,
        ]

        dataModel.metadata.each {final Metadata metadata ->
            final String key = "${metadata.namespace}:${metadata.key}"
            dataRow[key] = metadata.value
        }
        dataRow
    }

    static List<Map<String, String>> buildEnumerationDataRows(final String dataModelLabel, final Collection<EnumerationValue> enumerationValues) {

        enumerationValues.collect {final EnumerationValue enumValue ->
            buildEnumerationValueRow(dataModelLabel, enumValue)
        } as List<Map<String, String>>
    }

    static Map<String, String> buildEnumerationValueRow(final String dataModelLabel, final EnumerationValue enumerationValue) {
        final Map<String, String> dataRow = [
            'DataModel Name'  : dataModelLabel,
            'Enumeration Name': enumerationValue.label,
            Description       : enumerationValue.description,
            Key               : enumerationValue.key,
            Value             : enumerationValue.value,
        ]

        enumerationValue.metadata.each {final Metadata metadata ->
            final String key = "${metadata.namespace}:${metadata.key}"
            dataRow[key] = metadata.value
        }

        dataRow
    }

    static List<Map<String, String>> buildDataClassRows(final DataClass dataClass, final String path) {

        final List<Map<String, String>> dataClassSheetArray = []

        final String dataClassPath = path ? "${path} | ${dataClass.label}" : dataClass.label

        final Map<String, String> dataRow = [
            'DataClass Path'          : dataClassPath,
            'DataElement Name'        : '',
            Description               : dataClass.description,
            'Minimum \r\nMultiplicity': dataClass.minMultiplicity || dataClass.minMultiplicity == 0 ? dataClass.minMultiplicity.toString() : '',
            'Maximum \r\nMultiplicity': dataClass.maxMultiplicity || dataClass.maxMultiplicity == 0 ? dataClass.maxMultiplicity.toString() : '',
            'DataType Name'           : '',
            'DataType Reference'      : ''
        ]

        dataClass.metadata.each {final Metadata metadata ->
            final String key = "${metadata.namespace}:${metadata.key}"
            dataRow[key] = metadata.value
        }

        dataClassSheetArray.add(dataRow)

        dataClass.dataElements?.each {final DataElement dataElement ->
            final Map dataElementDataRow = [
                'DataClass Path'          : dataClassPath,
                'DataElement Name'        : dataElement.label,
                Description               : dataElement.description,
                'Minimum \r\nMultiplicity': dataElement.minMultiplicity || dataElement.minMultiplicity == 0 ? dataElement.minMultiplicity.toString() : '',
                'Maximum \r\nMultiplicity': dataElement.maxMultiplicity || dataElement.maxMultiplicity == 0 ? dataElement.maxMultiplicity.toString() : '',
                'DataType Name'           : dataElement.dataType.label,
            ]
            if (dataElement.dataType.isReferenceType()) {
                final List<String> classPath = getClassPath(dataElement.dataType.referenceClass, [])
                if (!classPath.isEmpty()) {
                    dataElementDataRow['DataType Reference'] = classPath.join(' | ')
                }
            }

            dataElement.metadata.each {final Metadata metadata ->
                final String key = "${metadata.namespace}:${metadata.key}"
                dataElementDataRow[key] = metadata.value
            }
            dataClassSheetArray.add(dataElementDataRow)
        }
        dataClass.dataClasses?.each {final DataClass childDataClass ->
            dataClassSheetArray.addAll(buildDataClassRows(childDataClass, dataClassPath))
        }
        dataClassSheetArray
    }

    static List<String> getClassPath(final DataClass dataClass, final List<String> path) {
        // We don't want paths with nulls in
        if (dataClass.label == null) {return []}
        path.add(0, dataClass.label)
        if (dataClass.parentDataClass) {
            return getClassPath(dataClass.parentDataClass, path)
        } else {
            return path
        }
    }

    static void writeArrayToSheet(final XSSFWorkbook workbook, final XSSFSheet sheet, final List<Map<String, String>> array) {
        final long start = System.currentTimeMillis()
        final Set<String> headers = [] as Set<String>
        if (array.size() > 0) {
            headers.addAll(array.get(0).keySet())
        }
        array.eachWithIndex {final Map<String, String> entry, final int i ->
            if (i != 0) {
                headers.addAll(entry.keySet())
            }
        }
        final XSSFRow headerRow = sheet.createRow(0)
        headers.eachWithIndex {final String header, final Integer idx ->
            final Cell headerCell = headerRow.createCell(idx)
            headerCell.setCellValue(header)
        }
        setHeaderRow(headerRow, workbook)

        final long substart = System.currentTimeMillis()

        array.eachWithIndex {final Map<String, String> map, final Integer rowIdx ->

            final Row valueRow = sheet.createRow(rowIdx + 1)
            headers.eachWithIndex {final String header, final Integer cellIdx ->
                final Cell valueCell = valueRow.createCell(cellIdx)
                valueCell.setCellValue(map[header] ?: '')
            }
        }
        log.debug('Writing rows took {}', timeTaken(substart))
        // Disabled this as it adds 1/3 of the time to run
        // substart = System.currentTimeMillis()

        autoSizeColumns(sheet, 0..(headers.size() + 1))
        // log.debug('Resizing columns took {}',  Utils.timeTaken(substart))
        log.debug('Writing array of {} rows to sheet took {}', array.size(), timeTaken(start))
    }

    static void setHeaderRow(final XSSFRow row, final XSSFWorkbook workbook) {

        XSSFCellStyle style = row.getRowStyle()
        if (!style) {
            style = workbook.createCellStyle()
        }

        final Font font = workbook.createFont()
        font.setBold(true)
        style.setFont(font)
        style.setFillForegroundColor((short) 123456)
        style.setWrapText(true)
        row.setRowStyle(style)

        final int n = row.getLastCellNum()
        for (int i = 0; i < n; i++) {
            row.getCell(i).setCellStyle(style)
        }

        row.setHeight((short) -1)
    }

    static String timeTaken(final long start) {
        getTimeString(System.currentTimeMillis() - start)
    }

    static String getTimeString(final long duration) {
        durationToString(Duration.ofMillis(duration))
    }

    static String durationToString(final Duration duration) {
        final StringBuilder sb = new StringBuilder(64)

        final int hrs = duration.toHoursPart()
        final int mins = duration.toMinutesPart()
        final int secs = duration.toSecondsPart()
        final int ms = duration.toMillisPart()

        if (hrs > 0) {
            sb.append(duration.toHoursPart()).append(' hr')
            if (hrs > 1) sb.append('s')
            sb.append(' ')
        }

        if (mins > 0) {
            sb.append(mins).append(' min')
            if (mins > 1) sb.append('s')
            sb.append(' ')
        }

        if (secs > 0) {
            sb.append(secs).append(' sec')
            if (secs > 1) sb.append('s')
            sb.append(' ')
        }

        sb.append(ms).append(' ms').toString()

        sb.toString()
    }


    static List<Map<String, String>> getSheetValues(final Map<String, String> intendedColumnMappings, final XSSFSheet sheet) {
        final int rows = sheet.getPhysicalNumberOfRows()
        final List<Map<String, String>> returnValues = new ArrayList<>(rows)
        // Nothing to see here, move along
        if (rows < 2) {return returnValues}
        final Map<String, Integer> expectedSheetColumns = new HashMap<>(intendedColumnMappings.size())
        final Map<String, Integer> otherSheetColumns = [:]
        final XSSFRow firstRow = sheet.getRow(0)

        for (int col = 0; col < 16384; col++) {
            final Cell cell = firstRow.getCell(col)
            if (!cell) {break}

            final String headerText = getCellValueAsString(cell)
            boolean found = false
            Iterator<Map.Entry<String, String>> intendedColumnMappingsIterator = intendedColumnMappings.entrySet().iterator()

            while (intendedColumnMappingsIterator.hasNext()) {
                final Map.Entry entry = intendedColumnMappingsIterator.next()
                final String columnName = entry.key
                final String regex = entry.value
                if (headerText.toLowerCase().trim() ==~ regex.toLowerCase().trim()) {
                    expectedSheetColumns[columnName] = col
                    found = true
                    break
                }
            }

            if (!found) {
                otherSheetColumns[headerText] = col
            }
        }

        if (expectedSheetColumns.size() != intendedColumnMappings.size()) {
            throw new RuntimeException("EIS03: Missing header: ${sheet.getSheetName()} sheet should include the following headers: ${intendedColumnMappings.values()}")
        }

        final Iterator<Row> rowIterator = sheet.rowIterator()
        // burn the header row
        rowIterator.next()

        while (rowIterator.hasNext()) {
            final XSSFRow row = (XSSFRow) rowIterator.next()
            final Map<String, String> rowValues = new HashMap<>(intendedColumnMappings.size())
            intendedColumnMappings.each {final String columnName, final String regex ->
                final String value = getCellValueAsString(row.getCell(expectedSheetColumns[columnName]))
                rowValues[columnName] = value
            }
            otherSheetColumns.keySet().each {final String columnName ->
                final String value = getCellValueAsString(row.getCell(otherSheetColumns[columnName]))
                rowValues[columnName] = value
            }
            returnValues << rowValues
        }
        return returnValues
    }

    static String getCellValueAsString(final Cell cell) {
        cell ? dataFormatter.formatCellValue(cell).replaceAll(/’/, '\'').replaceAll(/—/, '-').trim() : ''
    }

    static Map<String, Map<String, DataType>> calculateEnumerationDataTypes(final List<Map<String, String>> sheetValues, final String myNamespace) {

        final Map<String, Map<String, DataType>> returnValues = [:]
        sheetValues.each {final Map<String, String> columnValues ->
            final String dataModelName = columnValues.dataModelName
            final String label = columnValues.enumerationName
            final String description = columnValues.description
            final String key = columnValues.key
            final String value = columnValues.value

            Map<String, DataType> modelEnumTypes = returnValues[dataModelName]
            if (!modelEnumTypes) {
                modelEnumTypes = [:]
                returnValues[dataModelName] = modelEnumTypes
            }
            DataType enumerationDataType = modelEnumTypes[label]
            if (!enumerationDataType) {
                enumerationDataType = new DataType(label: label, description: description)
                modelEnumTypes[label] = enumerationDataType
            }
            final EnumerationValue enumerationValue = new EnumerationValue(key: key, value: value)
            addMetadataFromExtraColumns(enumerationValue, ENUM_SHEET_COLUMNS, columnValues, myNamespace)
            enumerationDataType.enumerationValues << enumerationValue
        }
        return returnValues
    }

    static void addMetadataFromExtraColumns(final AdministeredItem entity, final Map<String, String> expectedColumnMappings, final Map<String, String> columnValues,
                                            final String myNamespace) {
        final Set<String> expectedKeys = expectedColumnMappings.keySet()
        columnValues.keySet().each {final String columnName ->
            if (!expectedKeys.contains(columnName)) {
                String key
                String namespace
                if (columnName.contains(':')) {
                    final String[] components = columnName.split(':')
                    namespace = components[0]
                    key = components[1]
                } else {
                    namespace = myNamespace
                    key = columnName
                }
                if (columnValues[columnName]) {
                    entity.metadata << new Metadata(namespace: namespace, key: key, value: columnValues[columnName])
                }
            }
        }
    }

    static DataModelType dataModelType(final String type) {
        final String dataModelTypeMatch = type.toLowerCase().replaceAll('[ _]', '')
        final DataModelType dataModelType
        switch (dataModelTypeMatch) {
            case 'dataasset':
                dataModelType = DataModelType.DATA_ASSET
                break
            case 'datastandard':
                dataModelType = DataModelType.DATA_STANDARD
                break
            default:
                throw new RuntimeException("SEIS03: Invalid Data Model Type '${type}' ")
        }
        return dataModelType
    }

    static DataModel dataModelFromRow(final Map<String, String> columnValues) {

        final DataModelType dataModelType = dataModelType(columnValues.type)

        new DataModel(label: columnValues.name,
                      description: columnValues.description,
                      organisation: columnValues.organisation,
                      author: columnValues.author,
                      dataModelType: dataModelType)
    }

    static void addClassesAndElements(final DataModel dataModel, final XSSFSheet dataModelSheet, final Map<String, DataType> enumerationDataTypes, final String myNamespace) {
        final List<Map<String, String>> sheetValues = getSheetValues(MODEL_SHEET_COLUMNS, dataModelSheet)
        final List<Map<String, String>> referenceTypeDataElementRows = sheetValues.findAll {final Map<String, String> it -> it.dataTypeReference}
        final Map<String, DataType> modelDataTypes = [:]
        DataClass dataClass = null
        String previousDataClassPath = null
        (sheetValues - referenceTypeDataElementRows).each {final Map<String, String> row ->
            // Dont try to import any rows without a dataClasspath as we dont know where to place them
            if (row.dataClassPath) {
                final String dataClassPath = row.dataClassPath
                final String name = row.dataElementName
                AdministeredItem createdElement

                if (!previousDataClassPath?.equals(dataClassPath) && name || !name) {
                    // We're dealing with a data class
                    previousDataClassPath = dataClassPath
                    final String description = row.description
                    final String minMult = row.minMultiplicity
                    final String maxMult = row.maxMultiplicity

                    dataClass = getOrCreateClassFromPath(dataModel, dataClassPath, name ? '' : description,
                                                         name ? 1 : minMult ? Integer.parseInt(minMult) : 1,
                                                         name ? 1 : maxMult ? maxMult == '*' ? -1 : Integer.parseInt(maxMult) : 1)
                    createdElement = name ? addDataElement(dataModel, dataClass, modelDataTypes, enumerationDataTypes, row) : dataClass
                } else {
                    createdElement = addDataElement(dataModel, dataClass, modelDataTypes, enumerationDataTypes, row)
                }
                addMetadataFromExtraColumns(createdElement, MODEL_SHEET_COLUMNS, row, myNamespace)
            }
        }
        referenceTypeDataElementRows.each {final Map<String, String> row ->
            // Dont try an import any rows without a dataClasspath as we dont know where to place them
            if (row.dataClassPath) {
                final String dataClassPath = row.dataClassPath
                final String description = row.description
                final String minMult = row.minMultiplicity
                final String maxMult = row.maxMultiplicity
                dataClass = getOrCreateClassFromPath(dataModel, dataClassPath, description,
                                                     minMult ? Integer.parseInt(minMult) : 1,
                                                     maxMult ? maxMult == '*' ? -1 : Integer.parseInt(maxMult) : 1)
                final AdministeredItem createdElement = addDataElement(dataModel, dataClass, modelDataTypes, enumerationDataTypes, row)
                addMetadataFromExtraColumns(createdElement, MODEL_SHEET_COLUMNS, row, myNamespace)
            }
        }
        modelDataTypes.values().each {final DataType dataType ->
            dataModel.dataTypes << dataType
        }
        enumerationDataTypes.values().each {final DataType dataType ->
            dataModel.dataTypes << dataType
        }
    }

    static DataClass getOrCreateClassFromPath(final DataModel dataModel, final String path, final String description = '', final Integer minMultiplicity = null,
                                              final Integer maxMultiplicity = null) {
        declareDataClassByPath(dataModel, getDataClassPathLabels(path), description, minMultiplicity, maxMultiplicity)
    }

    static DataElement addDataElement(final DataModel dataModel, final DataClass parentDataClass, final Map<String, DataType> modelDataTypes,
                                      final Map<String, DataType> enumerationDataTypes, final Map<String, String> row) {
        final String name = row.dataElementName
        final String description = row.description
        final String minMult = row.minMultiplicity
        final String maxMult = row.maxMultiplicity
        final String typeName = row.dataTypeName
        final String typeReference = row.dataTypeReference
        // We're dealing with a data element
        final DataType elementDataType
        if (enumerationDataTypes[typeName]) {
            elementDataType = enumerationDataTypes[typeName]
        } else if (modelDataTypes[typeName]) {
            elementDataType = modelDataTypes[typeName]
        } else {
            if (typeReference) {
                final DataClass referenceDataClass = findDataClassByPath(dataModel, getDataClassPathLabels(typeReference))

                elementDataType = declareReferenceDataTypeForDataModel(dataModel, referenceDataClass)
            } else {
                elementDataType = declarePrimitiveDataTypeForDataModel(dataModel, typeName)
            }
            modelDataTypes[typeName] = elementDataType
        }

        declareDataElement(parentDataClass, name, description, elementDataType, minMult ? Integer.parseInt(minMult, 10) : 1,
                           maxMult ? maxMult == '*' ? -1 : Integer.parseInt(maxMult, 10) : 1)
    }

    static List<String> getDataClassPathLabels(final String dataClassPath) {
        dataClassPath.split('\\|').toList()*.trim()
    }

    /*
    Ensures that the path of DataModel -> DataClass (-> DataClass)* exists and that the DataClass at the end of the chain is updated with
    these parameters
    the 'declare' in the name reflects that this acts declaratively
     */

    static DataClass declareDataClassByPath(final DataModel dataModel, final List<String> pathLabels, final String description, final Integer minMultiplicity = 1,
                                            final Integer maxMultiplicity = 1) {

        DataClass parentDataClass = null
        final int n = pathLabels.size()

        for (int p = 0; p < n; p++) {

            final String localDataClassLabel = pathLabels.get(p)

            // Try to find localDataClassLabel in either the parentDataClass or the dataModel

            DataClass localDataClass
            if (parentDataClass == null) {
                localDataClass = dataModel.dataClasses.find {final DataClass candidateDataClass -> candidateDataClass.label == localDataClassLabel}
            } else {
                localDataClass = parentDataClass.dataClasses.find {final DataClass candidateDataClass -> candidateDataClass.label == localDataClassLabel}
            }

            if (localDataClass == null) {
                localDataClass = new DataClass(dataModel: dataModel, label: localDataClassLabel)

                if (parentDataClass == null) {
                    dataModel.dataClasses << localDataClass
                } else {
                    parentDataClass.dataClasses << localDataClass
                }
            }

            parentDataClass = localDataClass

            // Last one in the chain
            if (p == n - 1) {
                localDataClass.description = description
                localDataClass.minMultiplicity = minMultiplicity
                localDataClass.maxMultiplicity = maxMultiplicity
                return localDataClass
            }
        }
        return null
    }

    static DataClass findDataClassByPath(final DataModel dataModel, final List<String> pathLabels) {

        DataClass parentDataClass = null

        DataClass localDataClass = null
        final int n = pathLabels.size()
        for (int p = 0; p < n; p++) {
            final String localDataClassLabel = pathLabels.get(p)

            if (parentDataClass == null) {
                localDataClass = dataModel.dataClasses.find {final DataClass candidateDataClass -> candidateDataClass.label == localDataClassLabel}
            } else {
                localDataClass = parentDataClass.dataClasses.find {final DataClass candidateDataClass -> candidateDataClass.label == localDataClassLabel}
            }

            if (localDataClass == null) {
                throw new RuntimeException("DCS01: Cannot find DataClass for path [${pathLabels.join('|')}] as DataClass [${localDataClassLabel}] does not exist")
            }

            parentDataClass = localDataClass
        }

        return localDataClass
    }

    static DataType declareReferenceDataTypeForDataModel(final DataModel dataModel, final DataClass referencedDataClass, final String description = null) {

        DataType referenceDataType =
            dataModel.dataTypes.find {final DataType candidateDataType -> candidateDataType.referenceClass != null && candidateDataType.referenceClass.is(referencedDataClass)}

        if (referenceDataType == null) {

            referenceDataType = new DataType()
            referenceDataType.domainType(DataType.DataTypeKind.REFERENCE_TYPE)
            referenceDataType.referenceClass = referencedDataClass
            if (description != null && !description.trim().isEmpty()) {
                referenceDataType.description = description
            }
            dataModel.dataTypes << referenceDataType
        }

        return referenceDataType
    }

    static DataType declarePrimitiveDataTypeForDataModel(final DataModel dataModel, final String label, final String description = null) {

        DataType primitiveDataType =
            dataModel
                .dataTypes
                .find {final DataType candidateDataType -> candidateDataType.dataTypeKind == DataType.DataTypeKind.PRIMITIVE_TYPE && candidateDataType.label == label}

        if (primitiveDataType == null) {

            primitiveDataType = new DataType(label: label)
            primitiveDataType.domainType(DataType.DataTypeKind.PRIMITIVE_TYPE)
            if (description != null && !description.trim().isEmpty()) {
                primitiveDataType.description = description
            }

            dataModel.dataTypes << primitiveDataType
        }

        return primitiveDataType
    }

    static DataType declareEnumerationDataTypeForDataModel(final DataModel dataModel, final String label, final String description = null) {

        DataType enumerationDataType =
            dataModel
                .dataTypes
                .find {final DataType candidateDataType -> candidateDataType.dataTypeKind == DataType.DataTypeKind.ENUMERATION_TYPE && candidateDataType.label == label}

        if (enumerationDataType == null) {

            enumerationDataType = new DataType(label: label)
            enumerationDataType.domainType(DataType.DataTypeKind.ENUMERATION_TYPE)
            if (description != null && !description.trim().isEmpty()) {
                enumerationDataType.description = description
            }

            dataModel.dataTypes << enumerationDataType
        }

        return enumerationDataType
    }

    static DataElement declareDataElement(final DataClass asChildOfDataClass, final String label, final String description, final DataType dataType,
                                          final Integer minMultiplicity = 0, final Integer maxMultiplicity = 1) {

        if (label == null) {throw new RuntimeException("Attempt to declare a DataElement without a label")}
        if (dataType == null) {throw new RuntimeException("Attempt to declare a DataElement without a DataType")}

        DataElement dataElement = asChildOfDataClass.dataElements.find {final DataElement candidateDataElement -> candidateDataElement.label == label}

        if (dataElement == null) {
            dataElement = new DataElement(label: label, description: description, minMultiplicity: minMultiplicity, maxMultiplicity: maxMultiplicity,
                                          dataModel: asChildOfDataClass.dataModel, dataClass: asChildOfDataClass)
            if (dataType.label == null && dataType.label.trim().isEmpty()) {dataType.label = "${label}-dataType"}
            dataElement.dataType = dataType
            asChildOfDataClass.dataElements << dataElement
        } else {
            if (dataElement.label == null && label != null) {dataElement.label = label}
            if (dataElement.description == null && description != null) {dataElement.description = description}
            if (dataElement.dataType == null && dataType != null) {dataElement.dataType = dataType}
            if (dataElement.minMultiplicity == null && minMultiplicity != null) {dataElement.minMultiplicity = minMultiplicity}
            if (dataElement.maxMultiplicity == null && maxMultiplicity != null) {dataElement.maxMultiplicity = maxMultiplicity}
        }

        return dataElement
    }

    static void loadDataModelsFromWorkbook(final XSSFWorkbook workbook, final List<DataModel> loadedDataModels, final String namespace) {

        final List<DataModelDataRow> dataModelDataRows = loadDataModelDataRows(workbook)

        dataModelDataRows.forEach {final DataModelDataRow dataRow ->

            if (workbook.getSheet(dataRow.sheetKey) == null) {
                log.warn('DataModel [{}] with key [{}] will not be imported as it has no corresponding sheet', dataRow.name, dataRow.sheetKey)
                return []
            }

            List<ContentDataRow> contentDataRows = loadContentDataRows(workbook, dataRow.sheetKey)
            if (!contentDataRows) return []
            loadedDataModels << loadDataModel(dataRow, contentDataRows, namespace)
        }

    }

    static List<DataModelDataRow> loadDataModelDataRows(final XSSFWorkbook workbook) {
        loadDataRows(workbook, DataModelDataRow, DATAMODEL_SHEETNAME, CONTENT_NUM_HEADER_ROWS, CONTENT_ID_COLUMN_INDEX)
    }

    static List<ContentDataRow> loadContentDataRows(final XSSFWorkbook workbook, final String sheetName) {
        loadDataRows(workbook, ContentDataRow, sheetName, CONTENT_NUM_HEADER_ROWS, CONTENT_ID_COLUMN_INDEX)
    }

    static <K extends StandardDataRow> List<K> loadDataRows(final XSSFWorkbook workbook, final Class<K> dataRowClass, final String sheetName, final int numberOfHeaderRows,
                                                            final int idCellIndex) {

        final XSSFSheet sheet = workbook.getSheet(sheetName)
        if (sheet == null) {throw new RuntimeException("EFS05: No sheet named [${sheetName}]")}

        final List<CellRangeAddress> mergedRegions = sheet.mergedRegions.findAll {
            it.firstColumn == idCellIndex && it.lastRow != it.firstRow
        }.sort {final CellRangeAddress it -> it.firstRow}
        log.trace('{} merged regions found, reduced down to {} useable regions', sheet.mergedRegions.size(), mergedRegions.size())

        final Iterator<Row> rowIterator = sheet.rowIterator()
        numberOfHeaderRows.times {final Integer it -> rowIterator.next()}

        final List<K> dataRows = []

        while (rowIterator.hasNext()) {
            XSSFRow row = rowIterator.next() as XSSFRow
            String cellValue = CellHandlerUtil.getCellValue(row.getCell(idCellIndex))
            if (!cellValue) continue

            log.trace('Examining row {} at {}', cellValue, row.rowNum)
            K dataRow = dataRowClass.getDeclaredConstructor().newInstance().tap {initialiseRow row}
            dataRows << dataRow

            Integer lastRowIndexOfMerge = mergedRegions.find {final CellRangeAddress it -> it.firstRow == row.rowNum}?.lastRow
            if (!lastRowIndexOfMerge) continue

            log.trace('Merged region for {} from {} to {}(inclusive)', cellValue, row.rowNum, lastRowIndexOfMerge)
            dataRow.addToMergedContentRows(row)
            while (rowIterator.hasNext() && row.rowNum < lastRowIndexOfMerge) {
                row = rowIterator.next() as XSSFRow
                dataRow.addToMergedContentRows(row)
            }
            log.trace('Data row has {} extra content', dataRow.mergedContentRows.size())
        }

        log.debug('Loaded {} data rows from sheet [{}]', dataRows.size(), sheet.sheetName)
        dataRows
    }

    static DataModel loadDataModel(final DataModelDataRow dataModelDataRow, final List<ContentDataRow> contentDataRows, final String namespace) {

        log.info('Loading DataModel [{}]', dataModelDataRow.name)

        final DataModel dataModel = createDataModel(dataModelDataRow, namespace)

        log.info('Loading DataClass rows for DataModel [{}]', dataModel.label)
        loadDataClassRows(contentDataRows).each {findOrCreateDataClass(dataModel, it, namespace)}

        log.info('Loading DataElement rows for DataModel [{}]', dataModel.label)
        loadDataElementRows(contentDataRows).each {findOrCreateDataElement(dataModel, it, namespace)}

        dataModel
    }

    static DataModel createDataModel(DataModelDataRow dataRow, final String namespace) {
        log.debug('Creating DataModel [{}]', dataRow.name)
        final DataModel dataModel = new DataModel(label: dataRow.name,
                                                  description: dataRow.description,
                                                  organisation: dataRow.organisation,
                                                  author: dataRow.author,
                                                  dataModelType: dataModelType(dataRow.type).label,
                                                  )
        addMetadataToCatalogueItem(dataModel, dataRow.metadata, namespace)
        dataModel
    }

    static Collection<List<ContentDataRow>> loadDataClassRows(final List<ContentDataRow> dataRows) {
        dataRows.groupBy {final ContentDataRow it -> it.dataClassPath}.sort().values()
    }

    static List<ContentDataRow> loadDataElementRows(final List<ContentDataRow> dataRows) {
        dataRows.findAll {final ContentDataRow it -> it.dataElementName}
    }

    static DataClass findOrCreateDataClass(final DataModel dataModel, final List<ContentDataRow> dataRows, final String namespace) {

        final ContentDataRow dataRow = dataRows.find {final ContentDataRow it -> !it.dataElementName} ?: dataRows.first()
        log.debug('Creating DataClass [{}]', dataRow.dataClassPath)
        declareDataClassByPath(
            dataModel, dataRow.dataClassPathList,
            dataRow.dataElementName ? null : dataRow.description,
            dataRow.dataElementName ? 1 : dataRow.minMultiplicity,
            dataRow.dataElementName ? 1 : dataRow.maxMultiplicity).tap {final DataClass dataClass ->
            if (!dataRow.dataElementName) addMetadataToCatalogueItem(dataClass, dataRow.metadata, namespace)
        }
    }

    static DataElement findOrCreateDataElement(final DataModel dataModel, final ContentDataRow dataRow, final String namespace) {
        DataClass dataClass = findDataClassByPath(dataModel, dataRow.dataClassPathList)
        DataType dataType = findOrCreateDataType(dataModel, dataRow)
        log.debug('Adding DataElement [{}] to DataClass [{}] with DataType [{}]', dataRow.dataElementName, dataClass.label, dataType.label)
        declareDataElement(
            dataClass, dataRow.dataElementName, dataRow.description, dataType, dataRow.minMultiplicity, dataRow.maxMultiplicity).tap {
            addMetadataToCatalogueItem(it, dataRow.metadata, namespace)
        }
    }

    static void addMetadataToCatalogueItem(final AdministeredItem catalogueItem, final List<MetadataColumn> metadata, final String namespace) {
        metadata.each {MetadataColumn metadataColumn ->
            log.debug('Adding Metadata [{}] to CatalogueItem [{}]', metadataColumn.key, catalogueItem.label)

            catalogueItem.metadata << new Metadata(namespace: metadataColumn.namespace ?: namespace, key: metadataColumn.key, value: metadataColumn.value)
        }
    }

    static DataType findOrCreateDataType(DataModel dataModel, ContentDataRow dataRow) {
        if (dataRow.mergedContentRows) return findOrCreateEnumerationType(dataModel, dataRow)
        if (dataRow.referenceToDataClassPath) return findOrCreateReferenceType(dataModel, dataRow)
        findOrCreatePrimitiveType(dataModel, dataRow)
    }

    static DataType findOrCreateEnumerationType(DataModel dataModel, ContentDataRow dataRow) {
        log.debug('DataElement [{}] has an EnumerationType', dataRow.dataElementName)
        final DataType enumerationDataType = declareEnumerationDataTypeForDataModel(dataModel, dataRow.dataTypeName, dataRow.dataTypeDescription)

        if (enumerationDataType.enumerationValues.isEmpty()) {
            dataRow
                .mergedContentRows
                .each {final EnumerationDataRow enumerationDataRow -> enumerationDataType.enumerationValues << new EnumerationValue(key: enumerationDataRow.key, value:
                    enumerationDataRow.value)}
        }

        return enumerationDataType
    }

    static DataType findOrCreateReferenceType(DataModel dataModel, ContentDataRow dataRow) {
        log.debug('DataElement [{}] has a ReferenceType', dataRow.dataElementName)
        DataClass referenceClass = findDataClassByPath(dataModel, dataRow.referenceToDataClassPathList)
        declareReferenceDataTypeForDataModel(
            dataModel, referenceClass, dataRow.dataTypeDescription)
    }

    static DataType findOrCreatePrimitiveType(DataModel dataModel, ContentDataRow dataRow) {
        log.debug('DataElement [{}] has a PrimitiveType', dataRow.dataElementName)
        declarePrimitiveDataTypeForDataModel(dataModel, dataRow.dataTypeName, dataRow.dataTypeDescription)
    }

}
