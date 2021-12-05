import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * @author Leonid Cheremshantsev
 */
public class XlsxWorker {

    public List<String> readFirstColumn(InputStream inputStream) throws IOException {
        List<String> strings = new ArrayList<>();

        try (Workbook workbook = new XSSFWorkbook(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                Cell cell = row.getCell(0);
                switch (cell.getCellType()) {
                    case STRING:
                        strings.add(cell.getStringCellValue());
                        break;
                    case NUMERIC:
                        strings.add(String.valueOf(cell.getNumericCellValue()));
                        break;
                    case BOOLEAN:
                        strings.add(String.valueOf(cell.getBooleanCellValue()));
                        break;
                }
            }
        }

        return strings;
    }


    public void write(String path) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet();

            for (int i = 0; i < 100; i++) {
                Row row = sheet.createRow(i);
                row.createCell(0).setCellValue(i);
                row.createCell(1).setCellValue(i + i);
                row.createCell(2).setCellValue(i * i);
            }

            workbook.write(new FileOutputStream(path));
        }
    }
}
