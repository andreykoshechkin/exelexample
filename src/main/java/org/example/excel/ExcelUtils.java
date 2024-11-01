package org.example.excel;

import lombok.experimental.UtilityClass;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.data.ExcelExportable;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

@UtilityClass
public class ExcelUtils {

    public static <T extends ExcelExportable> void writeDataToExcel(String filePath, String sheetName, T data, int[] columns) throws IOException {
        if (data == null) {
            throw new IllegalArgumentException("Данные не должны быть null");
        }
        writeDataToExcel(filePath, sheetName, List.of(data), columns);
    }

    private static <T extends ExcelExportable> void writeDataToExcel(String filePath, String sheetName, List<T> data, int[] columns) throws IOException {
        if (data == null) {
            throw new IllegalArgumentException("Данные не должны быть null");
        }

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(sheetName);

        int rowCount = 0;
        CellStyle headerStyle = setDefaultStyle(workbook);
        Row headerRow = sheet.createRow(rowCount++);
        headerRow.setHeight((short) (20 * 20));     //Настройка высоты header


        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue("Column " + columns[i]); // Или используйте реальные имена столбцов
            cell.setCellStyle(headerStyle);
        }

        // Write data
        setDataInRow(data, columns, sheet, rowCount); // Передаем стиль для данных

        for (int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Save file to the specified path
        try (FileOutputStream fileOutputStream = new FileOutputStream(getFilePath(filePath, sheetName))) {
            workbook.write(fileOutputStream);
            System.out.println(getFilePath(filePath, sheetName));
        } finally {
            workbook.close(); // Всегда закрывайте книгу для освобождения ресурсов
        }
    }

    private static String getFilePath(String filePath, String sheetName) {
        return filePath + "\\" + sheetName + ".xlsx"; // Правильный способ добавления обратного слэша
    }

    private static CellStyle setDefaultStyle(Workbook workbook) {
        CellStyle headerStyle = workbook.createCellStyle();

        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);

        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);

        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14); // Увеличиваем размер шрифта заголовка
        headerStyle.setFont(headerFont);

        return headerStyle;
    }


    private static <T extends ExcelExportable> void setDataInRow(List<T> data, int[] columns, Sheet sheet, int rowCount) {
        for (T obj : data) {
            Row row = sheet.createRow(rowCount++);
            for (int i = 0; i < columns.length; i++) {
                Cell cell = row.createCell(i);
                Object value = obj.getValueByColumnIndex(columns[i]);
                if (value instanceof Number) {
                    cell.setCellValue(((Number) value).doubleValue());
                } else if (value instanceof Boolean) {
                    cell.setCellValue((Boolean) value);
                } else {
                    cell.setCellValue(value != null ? value.toString() : "");
                }
            }
        }
    }
}
