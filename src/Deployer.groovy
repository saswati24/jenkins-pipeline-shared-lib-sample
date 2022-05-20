@Grapes(
   [  @Grab(group='org.apache.poi', module='poi', version='5.2.2'),
      @Grab(group='org.apache.poi', module='poi-ooxml', version='5.2.2'),
      @Grab(group = 'org.apache.commons', module = 'commons-lang3', version = '3.6'),
      @Grab(group='org.apache.logging.log4j', module='log4j-core', version='2.17.2')]
)
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
   
   def createFile(){
              // Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        // Creating a blank Excel sheet
        XSSFSheet sheet
                = workbook.createSheet("student Details");

        // Creating an empty TreeMap of string and Object][]
        // type
        Map<String, Object[]> data
                = new TreeMap<>();

        // Writing data to Object[]
        // using put() method
        data.put("1",
                new Object[] { "ID", "NAME", "LASTNAME" });
        data.put("2",
                new Object[] { 1, "Pankaj", "Kumar" });
        data.put("3",
                new Object[] { 2, "Prakashni", "Yadav" });
        data.put("4", new Object[] { 3, "Ayan", "Mondal" });
        data.put("5", new Object[] { 4, "Virat", "kohli" });

        // Iterating over data and writing it to sheet
        Set<String> keyset = data.keySet();

        int rownum = 0;

        for (String key : keyset) {

            // Creating a new row in the sheet
            Row row = sheet.createRow(rownum++);

            Object[] objArr = data.get(key);

            int cellnum = 0;

            for (Object obj : objArr) {

                // This line creates a cell in the next
                // column of that row
                Cell cell = row.createCell(cellnum++);

                if (obj instanceof String)
                    cell.setCellValue((String)obj);

                else if (obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }

        // Try block to check for exceptions
        try {

            // Writing the workbook
            FileOutputStream out = new FileOutputStream(
                    new File("temp.xlsx"));
            workbook.write(out);

            // Closing file output connections
            out.close();

            // Console message for successful execution of
            // program
            script.echo(
                    "temp.xlsx written successfully on disk.");
        }

        // Catch block to handle exceptions
        catch (Exception e) {

            // Display exceptions along with line number
            // using printStackTrace() method
            e.printStackTrace();
        }
   }
   
    def parse(path) {
        // Reading file from local directory
        FileInputStream file = new FileInputStream(
                new File(path));

        // Create Workbook instance holding reference to
        // .xlsx file
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
   
       def runFile() {
//        File file = new File(".");
//        for(String fileNames : file.list()) script.echo(fileNames);
//         script.echo(script.pwd)
//         def filename = 'C:\\Users\\jaidi\\OneDrive\\Documents\\GitHub\\simple-java-maven-app\\temp.xlsx'
        def filename = 'temp.xlsx'
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

    def run() {
        while (tries < 2) {
            Thread.sleep(1000)
            tries++
            script.echo("tries is numeric: " + StringUtils.isAlphanumeric("" + tries))
        }
       createFile()
       runFile()
    }
}
