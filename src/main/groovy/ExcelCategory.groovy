import org.apache.poi.hssf.util.CellRangeAddress as CRA
import org.apache.poi.ss.util.CellRangeAddress as CSRA
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.util.*
import org.apache.poi.ss.usermodel.*
import com.xlson.groovycsv.CsvParser

/**
 * A small utility to make producing Excel workbooks really easy 
 * with Groovy.
 */
class ExcelCategory {

    static CellStyle boldStyle

    static CellStyle redStyle 
    
    static CellStyle centered 

    static hlink_style
	
	static void reset() {
		boldStyle = null
		redStyle = null
		centered = null
		hlink_style = null
	}
    
    static Map styleOperators = [
         red : { Font font, CellStyle style -> 
            font.setColor(HSSFColor.RED.index);
        },
        gray : { Font font, CellStyle style -> 
            font.setColor(HSSFColor.GREY_50_PERCENT.index);
        },
        blue : { Font font, CellStyle style -> 
            font.setColor(HSSFColor.BLUE.index);
        },
        pink : { Font font, CellStyle style -> 
            font.setColor(HSSFColor.PINK.index);
        },
        green : { Font font, CellStyle style -> 
            font.setColor(HSSFColor.GREEN.index);
        },
        bottomBorder : { Font font, CellStyle style ->
            style.setBorderBottom((short)1)
        },
        bold : { Font font, style ->
            font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        },
        center : { Font font, style ->
            style.setAlignment(CellStyle.ALIGN_CENTER)
        },
        bgOrange : { Font font, CellStyle style ->
            style.setFillForegroundColor(HSSFColor.LIGHT_ORANGE.index)
            //style.setFillPattern(CellStyle.SOLID_FOREGROUND);
            style.setFillPattern(CellStyle.THIN_FORWARD_DIAG);
        }
    ]

    /**
     * Parse a CSV file with auto column names
     */
    static parseCSV(Object obj, String fileName, String separator='\t') {
        new CsvParser().parse(new File(fileName).text, separator: separator)
    }

    /**
     * Parse the given file with given column names and return an iterable object for scanning the lines
     */
    static parseCSV(Object obj, String fileName, List cols, String separator='\t') {
        new CsvParser().parse(new File(fileName).text, readFirstLine:true, separator: separator, columnNames: cols)
    }

    static Workbook workbook(Object obj) {
        Workbook wb = new XSSFWorkbook();
    }

    static Workbook newWorkbook(Object obj) {
        Workbook wb = new XSSFWorkbook();
    }

    static HashMap<Sheet,Boolean> firstRow = [:]

    static Row row(Sheet sheet) {
        if(!firstRow[sheet]) {
            Row r = sheet.createRow(0)
            firstRow[sheet] = Boolean.TRUE;
            return r
        }
        else {
            return sheet.createRow(sheet.lastRowNum+1)
        }
    }

    static void save(Workbook wb, String fileName) {
        FileOutputStream fileOut = new FileOutputStream(fileName);
        fileOut.withStream { wb.write(it) }
    }

    static Cell link(Cell cell, String url) {

        if(!hlink_style) {
            hlink_style = cell.row.sheet.workbook.createCellStyle();
            Font hlink_font = cell.row.sheet.workbook.createFont();
            hlink_font.setUnderline(Font.U_SINGLE);
            hlink_font.setColor(IndexedColors.BLUE.getIndex());
            hlink_style.setFont(hlink_font);
        }

        Hyperlink link = cell.sheet.workbook.creationHelper.createHyperlink(Hyperlink.LINK_URL);
        link.address = url
        cell.hyperlink = link
        cell.cellStyle = hlink_style
        return cell 
    }    
        
    static Cell cell(Cell cell, Object value) {
        cell.row.addCell(value)
    }

    static Cell bold(Cell cell) {
        createCachedBoldStyle(cell.sheet.workbook)
        cell.cellStyle = boldStyle
        return cell
    }

    def static createCachedBoldStyle(Workbook wb) {
        if(!boldStyle) {
            createBoldStyle(wb);
        }
        return boldStyle
    }

    static CellStyle createBoldStyle(Workbook wb) {
        boldStyle = wb.createCellStyle();
        Font font = wb.createFont();
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        boldStyle.setFont(font)
        return boldStyle
    }

    static Cell red(Cell cell) {
        def wb = cell.sheet.workbook
        createCachedRedStyle(wb)
        cell.cellStyle = redStyle
        return cell
    }
    
    static Cell center(Cell cell){
        if(!centered) {
            centered = cell.sheet.workbook.createCellStyle()
            centered.setAlignment(CellStyle.ALIGN_CENTER)
        }
        cell.cellStyle = centered
    }

    def static createCachedRedStyle(Workbook wb) {
        if(!redStyle) {
            createRedStyle(wb)
        }
        return redStyle
    }
    
    static CellStyle createRedStyle(Workbook wb) {
        redStyle = wb.createCellStyle();
        Font fontRed = wb.createFont();
        fontRed.setColor(HSSFColor.RED.index);
        redStyle.setFont(fontRed)
        return redStyle
    }
    
    static Cell addCell(Row row, Object value) {
      Cell c = row.createCell(row.lastCellNum>=0?row.lastCellNum : 0)
      try {
          c.setCellValue(Double.parseDouble(value.toString()))
      }
      catch(NumberFormatException e) {
          c.setCellValue(value.toString())
      }
      return c
    }

    static Row add(Row row, Object... values) {
        values.each { if(it instanceof List) it.each { row.addCell(it) } else row.addCell(it) }
        return row
    }
    
    static List<Cell> addCells(Row row, Object... values) {
        def cellsAdded = []
        values.each { if(it instanceof List) it.each { cellsAdded.add(row.addCell(it)) } else cellsAdded.add(row.addCell(it)) }
        return cellsAdded
    }

    static Sheet autoFilter(Sheet sheet, String cellRange) {
        // If no starting row provided, default to row 1
        // ie: we accept the form "A:G" or "A2:G"
        if(!(cellRange =~ /[0-9]/)) {
            cellRange = cellRange.split(":")[0] + "1:" + cellRange.split(":")[1]
        }
        sheet.setAutoFilter(CellRangeAddress.valueOf(cellRange+(sheet.lastRowNum+1)))   
        return sheet
    }
    
    static Sheet autoSize(Sheet sheet) {
        def lastColNum = sheet.getRow(1).getLastCellNum()
        for(int i=0; i<lastColNum; ++i) {
            sheet.autoSizeColumn((short)i)
        }
        return sheet
    }
}
