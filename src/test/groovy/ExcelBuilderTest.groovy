import static org.junit.Assert.*;

import org.junit.Test;


class ExcelBuilderTest {

    @Test
    public void test() {
        new ExcelBuilder().build {
            sheet("Test") {
                row {
                    bold {
                      cell("hello")
                      cell("world")
                    }
                }
                row {
                    red { bold {
                        cell(46.2)
                        cell(72)
                      }
                    }
                }
                row {
                    red {
                        cell(46.2)
                        cell(72).bold()
                    }
                }
                row {
                    bottomBorder {
                        cell(46.2)
                        
                        applyStyle("bgOrange") {
                            cell(72)
                        }
                    }
                }            }
                row { cells("plain","plain") }
        }.save("test.xlsx")
		
        new ExcelBuilder().build {
            sheet("Test") {
                row {
                    bold {
                      cell("hello")
                      cell("world").bold()
                    }
                }
            }
        }.save("test2.xlsx")
		
    }
    
//    @Test
    public void testImperativeStyles() {
        new ExcelBuilder().build {
            sheet("Test") {
                row {
                    bold()
                    cell("hello")
                    cell("world")
                    plain()
                }
                row {
                    cells("mars","is","over","there")
                }
                row {
                    red()
                    cell(46.2)
                    cell(72)
                    plain()
                }
            }.autoSize()
        }.save("imperative_styles.xlsx")
    } 
}
