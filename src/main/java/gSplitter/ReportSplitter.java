package gSplitter;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReportSplitter {

    private final String fileName;
    private final int maxRows;
    private int curFileIndex = 1;

    public ReportSplitter(String fileName, final int maxRows) {

        ZipSecureFile.setMinInflateRatio(0);

        this.fileName = fileName;
        this.maxRows = maxRows;

        try {
            /* Read in the original Excel file. */
            OPCPackage pkg = OPCPackage.open(new File(fileName));
            XSSFWorkbook workbook = new XSSFWorkbook(pkg);
            XSSFSheet sheet = workbook.getSheetAt(0);

            /* Only split if there are more rows than the desired amount. */
            if (sheet.getPhysicalNumberOfRows() >= maxRows) {
                splitWorkbook(workbook);
            }
            pkg.close();
        }
        catch (EncryptedDocumentException | IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    private void splitWorkbook(XSSFWorkbook workbook) {

        SXSSFWorkbook wb = new SXSSFWorkbook();
        Sheet sh = wb.createSheet();

        Row newRow;
        Cell newCell;

        int rowCount = 0;
        int colCount = 0;
        Row firstRow = null;

        XSSFSheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {
            newRow = sh.createRow(rowCount++);

            /* Time to create a new workbook? */
            if (rowCount == maxRows) {
                writeWorkBook(wb);
                wb = new SXSSFWorkbook();
                sh = wb.createSheet();
                rowCount = 0;
                newRow = sh.createRow(rowCount++);
                if (firstRow != null)
                    for (Cell cell : firstRow) {
                        newCell = newRow.createCell(colCount++);
                        newCell = setValue(newCell, cell);

                        CellStyle newStyle = wb.createCellStyle();
                        newStyle.cloneStyleFrom(cell.getCellStyle());
                        newCell.setCellStyle(newStyle);
                    }
                colCount = 0;
                newRow = sh.createRow(rowCount++);
            }

            for (Cell cell : row) {
                newCell = newRow.createCell(colCount++);
                newCell = setValue(newCell, cell);

                CellStyle newStyle = wb.createCellStyle();
                newStyle.cloneStyleFrom(cell.getCellStyle());
                newCell.setCellStyle(newStyle);
            }
            if (firstRow == null)
                firstRow = row;
            colCount = 0;
        }

        /* Only add the last workbook if it has content */
        if (wb.getSheetAt(0).getPhysicalNumberOfRows() > 0) {
            writeWorkBook(wb);
        }
    }

    /*
     * Grabbing cell contents can be tricky. We first need to determine what
     * type of cell it is.
     */
    private Cell setValue(Cell newCell, Cell cell) {
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                newCell.setCellValue(cell.getRichStringCellValue().getString());
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    newCell.setCellValue(cell.getDateCellValue());
                } else {
                    newCell.setCellValue(cell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                newCell.setCellValue(cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                newCell.setCellFormula(cell.getCellFormula());
                break;
            default:
                System.out.println("Could not determine cell type");
        }
        return newCell;
    }

    /* Write all the workbooks to disk. */
    private void writeWorkBook(SXSSFWorkbook wbs) {
        FileOutputStream out;
        try {
            String newFileName = fileName.substring(0, fileName.length() - 5);
            String fileName = newFileName + "_" + (curFileIndex++)  + ".xlsx";
            System.out.println("Writing to :'" + fileName + '\'');
            out = new FileOutputStream(new File(fileName));
            wbs.write(out);
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args){
        /* This will create a new workbook every 1000 rows. */
        String fileName = findByAttribute(args, "/f:");
        if (fileName == null)
            return;
        String val = findByAttribute(args, "/s:");
        int rowNum = 999;
        if (val != null)
            rowNum = Integer.parseInt(val.replaceAll("\\D", ""));
        System.out.println("Filename: '" + fileName + "', batch size:" + rowNum  );
        new ReportSplitter(fileName, rowNum);
    }

    private static String findByAttribute(String[] args, String s) {
        String value = Arrays.stream(args).filter(l -> l.startsWith(s)).findFirst().orElse(null);
        return value == null ? null : value.replaceFirst(s, "");
    }

}

