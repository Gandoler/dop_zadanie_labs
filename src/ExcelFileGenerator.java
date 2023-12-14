import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Random;

public class ExcelFileGenerator {
    public static void main(String[] args) {
        Workbook workbook = new XSSFWorkbook(); // Создаем новую книгу Excel
        FileOutputStream fileOut = null;

        try {
            Sheet sheet = workbook.createSheet("Data");

            // Добавляем заголовки столбцов
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("X");
            headerRow.createCell(1).setCellValue("Y1");
            headerRow.createCell(2).setCellValue("Y2");

            Random random = new Random();
            int rows = random.nextInt(10) + 5; // Генерируем случайное количество строк данных от 5 до 14

            for (int i = 1; i <= rows; i++) {
                Row row = sheet.createRow(i);

                // Генерируем случайные значения для X, Y1 и Y2
                int xValue = i;
                int y1Value = random.nextInt(100);
                int y2Value = random.nextInt(100);

                row.createCell(0).setCellValue(xValue);
                row.createCell(1).setCellValue(y1Value);
                row.createCell(2).setCellValue(y2Value);
            }

            // Сохраняем данные в Excel-файл
            fileOut = new FileOutputStream("/Users/gl.krutoimail.ru/Desktop/bonus/data.xlsx");
            workbook.write(fileOut);
            System.out.println("Excel файл успешно создан!");
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (workbook != null) {
                    ((XSSFWorkbook) workbook).close(); // Закрываем книгу после использования
                }
                if (fileOut != null) {
                    fileOut.close(); // Закрываем поток FileOutputStream
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
