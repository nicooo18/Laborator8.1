import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import java.io.FileOutputStream;
import java.io.IOException;


public class ExcelWriter {
    public void genereazaFisierExcelCuFormuleSiCulori() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Note");

        Object[][] studenti = {
                {"Name", "Surname", "Grade 1", "Grade 2", "Grade 3", "Grade 4", "Max", "Average"},
                {"Amit", "Shukla", 9, 8, 7, 5},
                {"Lokesh", "Gupta", 8, 9, 6, 7},
                {"John", "Adwards", 8, 8, 7, 6},
                {"Brian", "Schultz", 7, 6, 8, 9}
        };

        // Stiluri
        XSSFFont headerFont = workbook.createFont();
        headerFont.setBold(true);

        XSSFCellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFCellStyle yellowStyle = workbook.createCellStyle();
        yellowStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        yellowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        for (int i = 0; i < studenti.length; i++) {
            Row row = sheet.createRow(i);
            Object[] rand = studenti[i];
            for (int j = 0; j < rand.length; j++) {
                Cell cell = row.createCell(j);
                if (rand[j] instanceof String)
                    cell.setCellValue((String) rand[j]);
                else if (rand[j] instanceof Integer)
                    cell.setCellValue((Integer) rand[j]);

                // Stil antet
                if (i == 0)
                    cell.setCellStyle(headerStyle);
            }

            // Calculeaza Max si Average doar pentru randurile cu date
            if (i > 0) {
                int rowIndex = i + 1; // Excel este 1-based

                Cell maxCell = row.createCell(6); // Coloana G (Max)
                maxCell.setCellFormula("MAX(C" + rowIndex + ":F" + rowIndex + ")");
                maxCell.setCellStyle(yellowStyle);

                Cell avgCell = row.createCell(7); // Coloana H (Average)
                avgCell.setCellFormula("AVERAGE(C" + rowIndex + ":F" + rowIndex + ")");
                avgCell.setCellStyle(yellowStyle);
            }
        }

        // Autosize
        for (int i = 0; i < 8; i++) {
            sheet.autoSizeColumn(i);
        }

        // Scrie fisierul
        try (FileOutputStream out = new FileOutputStream("output8.xlsx")) {
            workbook.write(out);
            workbook.close();
            System.out.println("Fisierul 'output8.xlsx' a fost generat cu succes.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
