package org.maurodata.plugin.exceldatamodel.datarow

import groovy.transform.CompileStatic
import org.apache.poi.ss.usermodel.Row

@CompileStatic
abstract class StandardDataRow<K extends EnumerationDataRow> extends DataRow {

    Class<K> enumerationDataRowClass
    List<K> mergedContentRows = []

    StandardDataRow() {
    }

    StandardDataRow(Class<K> enumerationDataRowClass) {
        this.enumerationDataRowClass = enumerationDataRowClass
    }

    void addToMergedContentRows(Row row) {
        if (!enumerationDataRowClass) return
        K enumerationRow = enumerationDataRowClass.getDeclaredConstructor().newInstance()
        enumerationRow.initialiseRow(row)
        mergedContentRows << enumerationRow
    }
}