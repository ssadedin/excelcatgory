import org.apache.poi.ss.usermodel.*

class ExcelBuilder {
    
    Workbook workbook = null
    
    Sheet sheet = null
    
    Row row = null
    
    CellStyle activeStyle = null
    
    Map styleCache = [ : ]
    
    List activeStyles = []
    
    public ExcelBuilder() {
    }
    
    Workbook build(Closure c){
        this.workbook = ExcelCategory.workbook(new Object())
        this.workbook.metaClass.save = { fileName ->
            ExcelCategory.save(this.workbook, fileName)
        }
        c.delegate = this
		ExcelCategory.reset()
        use(ExcelCategory) {
          c()
        }
        return this.workbook
    }
    
    Sheet sheet(String name, Closure c = null) {
        sheet =  workbook.createSheet(name)
        if(c != null) {
          c.delegate = this
          c()
          return sheet
        }
    }
    
    Row row(Closure c) {
        if(sheet == null) 
            sheet = workbook.createSheet("Sheet1")
        row = sheet.row()
        c.delegate=this
        c()
        return row
    }
    
    Cell cell(Object contents) {
        Cell c = row.addCell(contents)
        if(this.activeStyle)
            c.cellStyle = this.activeStyle
        return c
    }
    
    List<Cell> cells(Object... objs) {
        def cellsAdded = row.addCells(objs)
        if(this.activeStyle) {
            cellsAdded.each { it.cellStyle = activeStyle }
        }
        return cellsAdded
    }
    
    void center(Closure c) {
        applyStyle("center",c)
    }
    
    void bold(Closure c) {
        applyStyle("bold",c)
    }
    
    void applyStyle(String styleDesc, Closure c = null) {
        this.activeStyles.add(styleDesc)
        int styleIndex = activeStyles.size()-1
        String key = this.activeStyles.sort().join(",")
        if(!this.styleCache[key]) {
            // Create a font and cell style
            CellStyle newStyle = workbook.createCellStyle();
            Font font = workbook.createFont()
            this.activeStyles.sort().each { 
                ExcelCategory.styleOperators[it](font,newStyle)
            }
            newStyle.setFont(font)
            styleCache[key] = newStyle
        }
        
        // Apply new style
        this.activeStyle = styleCache[key]
        if(c != null) {
          c.delegate = this
          c()
          this.activeStyles.remove(styleIndex)
          
          // Set back to old style
          this.activeStyle = this.styleCache[this.activeStyles.sort().join(",")]
        }
    }
    
    void withStyle(Closure c) {
        int styleSize = activeStyles.size()
        try {
            c.delegate = this
            c()
        }
        finally {
            while(activeStyles.size() > styleSize)
                popStyle()
        }
    }
    
    void popStyle() {
      this.activeStyles.remove(activeStyles.size()-1)
          
      // Set back to old style
      this.activeStyle = this.styleCache[this.activeStyles.sort().join(",")] 
    }
    
    void red(Closure c) {
        applyStyle("red",c)
    }
    
    void bottomBorder(Closure c) {
        applyStyle("bottomBorder",c)
    }
    
    void plain() {
        this.activeStyle = null
        this.activeStyles = []
    }
    
    void save(String fileName) {
        ExcelCategory.save(this.workbook, fileName)
    }
}