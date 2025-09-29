package org.maurodata.plugin.exceldatamodel.workbook

import groovy.transform.CompileStatic
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder

import java.awt.Color

@CompileStatic
class WorkbookCellAndBorderStyles {

    private static final Color BORDER_COLOUR = Color.decode('#FFFFFF')
    private static final Color CELL_COLOUR = Color.decode('#7BABF5')
    private static final double BORDER_COLOUR_TINT = -0.35d
    private static final double CELL_COLOUR_TINT = 0.6d

    XSSFCellStyle defaultCellStyle
    XSSFCellStyle colouredCellStyle
    XSSFCellStyle metadataBorderStyle
    XSSFCellStyle metadataColouredBorderStyle

    WorkbookCellAndBorderStyles(final XSSFWorkbook workbook) {
        this.defaultCellStyle = workbook.createCellStyle().tap {
            XSSFColor borderColour = new XSSFColor(BORDER_COLOUR, null).tap {
                setTint BORDER_COLOUR_TINT
            }
            setBorderTop BorderStyle.THIN
            setBorderLeft BorderStyle.THIN
            setBorderBottom BorderStyle.THIN
            setBorderRight BorderStyle.THIN
            setBorderColor XSSFCellBorder.BorderSide.TOP, borderColour
            setBorderColor XSSFCellBorder.BorderSide.LEFT, borderColour
            setBorderColor XSSFCellBorder.BorderSide.BOTTOM, borderColour
            setBorderColor XSSFCellBorder.BorderSide.RIGHT, borderColour
        }
        this.colouredCellStyle = workbook.createCellStyle().tap {
            cloneStyleFrom defaultCellStyle
            setFillForegroundColor new XSSFColor(CELL_COLOUR, null).tap {
                setTint CELL_COLOUR_TINT
            }
            setFillPattern FillPatternType.SOLID_FOREGROUND
        }
        this.metadataBorderStyle = workbook.createCellStyle().tap {
            cloneStyleFrom defaultCellStyle
            setBorderRight BorderStyle.DOUBLE
        } as XSSFCellStyle
        this.metadataColouredBorderStyle = workbook.createCellStyle().tap {
            cloneStyleFrom colouredCellStyle
            setBorderRight BorderStyle.DOUBLE
        } as XSSFCellStyle
    }
}
