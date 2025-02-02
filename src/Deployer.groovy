// @Grapes(
//    [  @Grab(group='org.apache.poi', module='poi', version='5.2.2'),
//       @Grab(group='org.apache.poi', module='poi-ooxml', version='5.2.2'),
//       @Grab(group = 'org.apache.commons', module = 'commons-lang3', version = '3.6')
// //     ,
// //       @Grab(group='org.apache.logging.log4j', module='log4j-core', version='2.17.2')
//    ])
import org.apache.commons.lang3.StringUtils
import org.apache.poi.ss.usermodel.*
import org.apache.poi.hssf.usermodel.*
import org.apache.poi.xssf.usermodel.*
import org.apache.poi.ss.util.*
import org.apache.poi.ss.usermodel.*
import java.io.*

class Deployer {
    int tries = 0
    Script script
   
    def parse(path) {
       script.echo("File path is " + path )
        // Reading file from local directory
        FileInputStream file = new FileInputStream(
                new File(path));

        // Create Workbook instance holding reference to
        // .xlsx file
        HSSFWorkbook workbook2 = new HSSFWorkbook(file);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        // Get first/desired sheet from the workbook
        XSSFSheet sheet = workbook.getSheetAt(0);

        // Iterate through each rows one by one
        Iterator<Row> rowIterator = sheet.iterator();
        Row row = rowIterator.next()
        def headers = getRowData(row)

        def rows = []
        while (rowIterator.hasNext()) {
            row = rowIterator.next()
            rows << getRowData(row)
        }
        file.close()
        [headers, rows]
    }

    def getRowData(Row row) {
        def data = []
        for (Cell cell : row) {
            getValue(row, cell, data)
        }
        data
    }

    def getRowReference(Row row, Cell cell) {
        def rowIndex = row.getRowNum()
        def colIndex = cell.getColumnIndex()
        CellReference ref = new CellReference(rowIndex, colIndex)
        ref.getRichStringCellValue().getString()
    }

    def getValue(Row row, Cell cell, List data) {
        def rowIndex = row.getRowNum()
        def colIndex = cell.getColumnIndex()
        def value = ""
        switch (cell.getCellType()) {
            case CellType.STRING:
                value = cell.getRichStringCellValue().getString();
                break;
            case CellType.NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    value = cell.getDateCellValue();
                } else {
                    value = cell.getNumericCellValue();
                }
                break;
            case CellType.BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case CellType.FORMULA:
                value = cell.getCellFormula();
                break;
            default:
                value = ""
        }
        data[colIndex] = value
        data
    }

    def toXml(header, row) {
        def obj = "<object>\n"
        row.eachWithIndex { datum, i ->
            def headerName = header[i]
            obj += "\t<$headerName>$datum</$headerName>\n"
        }
        obj += "</object>"
        script.echo(obj)
    }
   
       def runFile(String path) {
//        File file = new File(".");
//        for(String fileNames : file.list()) script.echo(fileNames);
// //         script.echo(script.pwd)
        def filename = path + '\\temp.xlsx'
//         GroovyExcelParser parser = new GroovyExcelParser()
        
        def (headers, rows) = parse(filename)
        script.echo("Headers")
        script.echo("------------------")
        headers.each { header ->
            script.echo( header )
        }
        script.echo( "\n" + "Rows" + "\n" + "------------------")
        rows.each { row ->
            toXml(headers, row)
            script.echo(row[1])
        }
    }

    def run(String path) {
       script.echo("Workspace directory paith " + path)
        while (tries < 2) {
            Thread.sleep(1000)
            tries++
            script.echo("tries is numeric: " + StringUtils.isAlphanumeric("" + tries))
        }
       runFile(path)
    }
}
