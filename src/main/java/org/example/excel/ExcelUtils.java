package org.example.excel;

import lombok.experimental.UtilityClass;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.data.AbstractExcelExportable;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

@UtilityClass
public class ExcelUtils {

    public static <T extends AbstractExcelExportable> void writeDataToExcel(String filePath, String sheetName, T data) throws IOException {
        if (data == null) {
            throw new IllegalArgumentException("Данные не должны быть null");
        }
        writeDataToExcel(filePath, sheetName, List.of(data));
    }

    public  <T extends AbstractExcelExportable> void writeDataToExcel(String filePath, String sheetName, List<T> data) throws IOException {
        if (data == null) {
            throw new IllegalArgumentException("Данные не должны быть null");
        }

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(sheetName);

        int rowCount = 0;
        CellStyle headerStyle = setDefaultStyle(workbook);
        Row headerRow = sheet.createRow(rowCount++);
        headerRow.setHeight((short) (20 * 20)); // Настройка высоты header

        // Получаем количество колонок от первого объекта
        int columnCount = data.get(0).getColumnCount();
        List<String> columnName = data.get(0).getColumnNames();


        // Создаем заголовки столбцов
        for (int i = 1; i <= columnCount; i++) {
            Cell cell = headerRow.createCell(i - 1);
            cell.setCellValue(columnName.get(i-1)); // Или используйте реальные имена столбцов
            cell.setCellStyle(headerStyle);
        }

        // Заполняем данные
        setDataInRow(data, sheet, rowCount, columnCount); // Передаем стиль для данных

        // Автоматическая подгонка ширины столбцов
        for (int i = 0; i < columnCount; i++) {
            sheet.autoSizeColumn(i);
        }

        // Сохраняем файл по указанному пути
        try (FileOutputStream fileOutputStream = new FileOutputStream(getFilePath(filePath, sheetName))) {
            workbook.write(fileOutputStream);
            System.out.println(getFilePath(filePath, sheetName));
        } finally {
            workbook.close(); // Всегда закрывайте книгу для освобождения ресурсов
        }
    }

    private String getFilePath(String filePath, String sheetName) {
        return filePath + "\\" + sheetName + ".xlsx"; // Правильный способ добавления обратного слэша
    }

    private CellStyle setDefaultStyle(Workbook workbook) {
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

    private <T extends AbstractExcelExportable> void setDataInRow(List<T> data, Sheet sheet, int rowCount, int columnCount) {
        for (T obj : data) {
            Row row = sheet.createRow(rowCount++);
            for (int i = 1; i <= columnCount; i++) {
                Cell cell = row.createCell(i - 1);
                Object value = obj.getValue(i); // Получаем значение по индексу
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