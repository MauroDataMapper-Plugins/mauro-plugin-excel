package org.maurodata.plugin.exceldatamodel.datarow

import groovy.transform.CompileStatic

@CompileStatic
class MetadataColumn {

    String namespace
    String key
    String value

    @Override
    boolean equals(Object metadataColumn) {
        if (is(metadataColumn)) return true
        if (getClass() != metadataColumn.class) return false
        namespace == (metadataColumn as MetadataColumn).namespace && key == (metadataColumn as MetadataColumn).key
    }

    @Override
    int hashCode() {
        (namespace?.hashCode() ?: 0) * 31 + (key?.hashCode() ?: 0)
    }

    @Override
    String toString() {
        "${namespace}/${key}"
    }
}
