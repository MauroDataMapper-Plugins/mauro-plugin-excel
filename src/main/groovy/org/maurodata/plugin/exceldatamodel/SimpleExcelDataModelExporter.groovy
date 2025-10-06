/**
 * Example Mauro Exporter Plugin
 * -----------------------------
 *
 * This file provides an example Exporter plugin for Mauro.
 * The exporter implements 'DataModelExporterPlugin' - other interfaces for other model types are
 * available - e.g. 'TerminologyExporterPlugin', etc
 *
 * This is intended to take one or more models and return a file - in this case represented by a
 * byte array of contents, a file name, and a file extension.  The 'content type' is provided for the API to
 * provide back to the client.
 */
package org.maurodata.plugin.exceldatamodel

import org.maurodata.domain.datamodel.DataModel
import org.maurodata.plugin.exceldatamodel.workbook.Utils
import org.maurodata.plugin.exporter.DataModelExporterPlugin

import groovy.transform.CompileStatic
import groovy.util.logging.Slf4j
import jakarta.inject.Singleton
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

@Slf4j
@Singleton
@CompileStatic
class SimpleExcelDataModelExporter implements DataModelExporterPlugin {

    final String version = '1.0.0'

    void setVersion(String version) {
        throw new RuntimeException("Can't set the version")
    }

    final String displayName = 'Simple Excel (XLSX) Exporter'

    void setDisplayName(String displayName) {
        throw new RuntimeException("Can't set the display name")
    }

    public static final String CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

    static final String DATAMODELS_TEMPLATE_FILENAME = 'Template_DataModel_Import_File.xlsx'

    @Override
    Boolean getCanExportMultipleDomains() {
        true
    }

    @Override
    String getFileExtension() {
        return 'xlsx'
    }

    @Override
    String getContentType() {
        return CONTENT_TYPE
    }

    /**
     * TODO: Implement this method to provide the file name for a downloadable file.
     * This is typically related to the name of the dataModel - e.g. model.label
     * @return
     */
    @Override
    String getFileName(DataModel model) {
        return null
    }

    @Override
    byte[] exportModel(DataModel model) {
        return exportModels([model])
    }

    @Override
    byte[] exportModels(Collection<DataModel> models) {
        log.info('Exporting DataModels to Excel')

        try (final XSSFWorkbook workbook = new XSSFWorkbook()) {

            Utils.simpleLoadDataModelsIntoWorkbook(models, workbook)

            try (final ByteArrayOutputStream exportStream = new ByteArrayOutputStream(models.size() * 100 * 1024)) {
                workbook.write(exportStream)
                log.info('DataModels exported')
                return exportStream.toByteArray()
            }
        }
    }
}
