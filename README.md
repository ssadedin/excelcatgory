excelcatgory
============

A small utility to ease making Excel files with groovy. It is just a thin wrapper on Apache POI. There are a lot of Excel builders out there, and this is really just one more that satisfies my particular desire for simplicity and concision.

Build / Install
============

    git clone https://github.com/ssadedin/excelcatgory.git
    cd excelcategory
    ./gradlew jar

If you want to make it available to all your groovy scripts:

   cp build/libs/excelcatgory.jar ~/.groovy/lib

Then you can test it out with a simple example such as:

    groovy -cp  build/libs/excelcatgory.jar  -e 'new ExcelBuilder().build { row { cells("hello","world") } }.save("test.xlsx")'

A more complex example:

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
