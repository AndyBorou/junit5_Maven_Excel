package writeReadExcelFile;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;


import java.io.*;
import java.util.Iterator;

public class ApacheExcelExample {

    private static final String FILE_NAME = "src/test/resources/MyFirstExcel.xlsx";

    @Test
    public void writeInformationInExcelFile() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("employeesInformation");
        Object[][] datatypes = {
                {"Name", "Position", "Work Experience"},
                {"Eric", "QA", 3},
                {"Kile", "Java Developer", 4},
                {"Stan", "JS Developer", 8}
        };

        int rowNum = 0;
        System.out.println("Creating excel file");

        for (Object[] employeeInformation : datatypes) {
            Row row = sheet.createRow(rowNum++);
            int colNum = 0;
            for (Object employee : employeeInformation) {
                Cell cell = row.createCell(colNum++);
                if (employee instanceof String) {
                    cell.setCellValue((String) employee);
                } else if (employee instanceof Integer) {
                    cell.setCellValue((Integer) employee);
                }
            }
        }

        try {
            FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Information was written");
    }

    @Test
    public void readInformationFromExcelFile() {
        try {
            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();

            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                while (cellIterator.hasNext()) {

                    Cell currentCell = cellIterator.next();
                    if (currentCell.getCellTypeEnum() == CellType.STRING) {
                        System.out.print(currentCell.getStringCellValue() + "--");
                    } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                        System.out.print(currentCell.getNumericCellValue() + "--");
                    }
                }
                System.out.println();
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
