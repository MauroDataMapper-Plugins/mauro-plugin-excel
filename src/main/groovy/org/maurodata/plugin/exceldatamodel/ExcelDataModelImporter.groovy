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

import org.maurodata.plugin.exceldatamodel.workbook.Utils
import org.maurodata.plugin.importer.FileParameter

import groovy.transform.CompileStatic
import groovy.util.logging.Slf4j
import jakarta.inject.Singleton
import org.maurodata.domain.datamodel.DataModel
import org.maurodata.plugin.importer.DataModelImporterPlugin
import org.maurodata.plugin.importer.FileImportParameters

import org.apache.poi.EncryptedDocumentException
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook

@Slf4j
@Singleton
@CompileStatic
class ExcelDataModelImporter implements DataModelImporterPlugin<FileImportParameters> {

    String version = '1.0.0'

    String displayName = 'Excel (XLSX) Importer'

    public static final String CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

    @Override
    Boolean handlesContentType(final String contentType) {
        return contentType.equalsIgnoreCase(CONTENT_TYPE)
    }

    @Override
    Class<FileImportParameters> importParametersClass() {
        return FileImportParameters
    }

    /**
     * TODO: This is the method that does all the work - implement this method.
     * Takes as input, a `params` object with all the parameters.
     * Returns a list of DataModel objects, which will be persisted by the framework.
     */
    @Override
    List<DataModel> importDomain(final FileImportParameters params) {

        if (params.importFile.fileContents.size() == 0) throw new RuntimeException('EIS02: Cannot import empty file')
        log.info('Importing {}', params.importFile.fileName)

        final FileParameter fileParameter = params.importFile

        final List<DataModel> loadedDataModels = []

        try (final XSSFWorkbook workbook = WorkbookFactory.create(fileParameter.getInputStream()) as XSSFWorkbook) {

            Utils.loadDataModelsFromWorkbook(workbook, loadedDataModels, this.getNamespace())

        } catch (final EncryptedDocumentException ignored) {
            throw new RuntimeException("EFS02: Excel file ${params.importFile.fileName} could not be read as it is encrypted")
        } catch (final IOException ex) {
            throw new RuntimeException("EFS04: Excel file ${params.importFile.fileName} could not be read", ex)
        }

        return loadedDataModels
    }


}
