/**
 * Example Mauro Importer Plugin
 * -----------------------------
 *
 * This file provides an example Importer plugin for Mauro.
 * The importer implements 'DataModelImporterPlugin' - other interfaces for other model types are
 * available - e.g. 'TerminologyImporterPlugin', etc
 *
 * The generic parameter determines the parameters provided.  'FileImportParameters' provides the
 * default set of import options, plus a file to provide the model details.  You may choose to extend
 * this with additional import options, and use this as the generic parameter instead.  A method
 * `importParametersClass` returns this class.
 *
 */
package org.maurodata.plugin.exceldatamodel

import org.maurodata.domain.datamodel.DataModel
import org.maurodata.domain.datamodel.DataType
import org.maurodata.domain.datamodel.EnumerationValue
import org.maurodata.domain.model.Item
import org.maurodata.domain.model.ItemReferencer
import org.maurodata.plugin.exceldatamodel.workbook.Utils
import org.maurodata.plugin.importer.DataModelImporterPlugin
import org.maurodata.plugin.importer.FileImportParameters
import org.maurodata.plugin.importer.FileParameter

import groovy.transform.CompileStatic
import groovy.util.logging.Slf4j
import jakarta.inject.Singleton
import org.apache.poi.EncryptedDocumentException
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

@Slf4j
@Singleton
@CompileStatic
class SimpleExcelDataModelImporter implements DataModelImporterPlugin<FileImportParameters> {

    String version = '1.0.0'

    String displayName = 'Simple Excel (XLSX) Importer'

    public static final String CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

    @Override
    Boolean handlesContentType(final String contentType) {
        return contentType.equalsIgnoreCase(CONTENT_TYPE)
    }

    @Override
    Class<FileImportParameters> importParametersClass() {
        return FileImportParameters
    }


    @Override
    List<DataModel> importDomain(FileImportParameters params) {

        if (params.importFile.fileContents.size() == 0) throw new RuntimeException('EIS02: Cannot import empty file')
        log.info('Importing {}', params.importFile.fileName)

        final FileParameter fileParameter = params.importFile

        try (final XSSFWorkbook workbook = WorkbookFactory.create(fileParameter.getInputStream()) as XSSFWorkbook) {

            final XSSFSheet dataModelsSheet = workbook.getSheet('DataModels')
            final XSSFSheet enumerationsSheet = workbook.getSheet('Enumerations')

            if (!dataModelsSheet) {
                throw new RuntimeException('EFS02: The Excel file does not include a sheet called [DataModels]')
            }

            Map<String, Map<String, DataType>> enumerationDataTypes
            if (enumerationsSheet) {
                enumerationDataTypes = Utils.calculateEnumerationDataTypes(Utils.getSheetValues(Utils.ENUM_SHEET_COLUMNS, enumerationsSheet), this.getNamespace())
            } else {
                enumerationDataTypes = [:]
            }

            List<Map<String, String>> sheetValues = Utils.getSheetValues(Utils.DATAMODEL_SHEET_COLUMNS, dataModelsSheet)
            final List<DataModel> dataModels = new ArrayList<>(sheetValues.size())

            sheetValues.forEach  {final Map<String, String> row ->

                final DataModel dataModel = Utils.dataModelFromRow(row)
                Utils.addMetadataFromExtraColumns(dataModel, Utils.DATAMODEL_SHEET_COLUMNS, row, this.getNamespace())

                final String sheetKey = row.sheetKey
                final XSSFSheet modelSheet = workbook.getSheet(sheetKey)
                if (!dataModelsSheet) {
                    throw new RuntimeException("EFS06: The Excel file does not include a sheet called ${sheetKey} referenced for " +
                                               "model'${dataModel.label}'")
                }

                Utils.addClassesAndElements(dataModel, modelSheet, enumerationDataTypes[dataModel.label] ?: [:] as Map<String, DataType>, this.getNamespace())

                dataModels << dataModel
            }

            return dataModels

        } catch (final EncryptedDocumentException ignored) {
            throw new RuntimeException("EFS02: Excel file ${params.importFile.fileName} could not be read as it is encrypted")
        } catch (final IOException ex) {
            throw new RuntimeException("EFS04: Excel file ${params.importFile.fileName} could not be read", ex)
        }
    }
}
