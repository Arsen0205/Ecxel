package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class App
{
    public static void main(String[] args) {
        try {
            // --- Выбор файлов ---
            JFileChooser fc = new JFileChooser();
            fc.setDialogTitle("Выберите файл ORIGIN (актуальные данные)");
            if (fc.showOpenDialog(null) != JFileChooser.APPROVE_OPTION) return;
            File originFile = fc.getSelectedFile();

            fc.setDialogTitle("Выберите файл MAIN (основные данные для обновления)");
            if (fc.showOpenDialog(null) != JFileChooser.APPROVE_OPTION) return;
            File mainFile = fc.getSelectedFile();

            // --- Читаем оба файла ---
            List<List<String>> originData = readExcel(originFile);
            List<List<String>> mainData = readExcel(mainFile);

            // --- Сопоставляем по номеру рейса (A = индекс 0) ---
            int keyCol = 0;     // рейс
            int carCol = 1;     // госномер
            int driverCol = 2;  // водитель

            // Индексируем Origin по номеру рейса
            Map<String, List<String>> originMap = new HashMap<>();
            for (List<String> row : originData) {
                if (row.size() > keyCol) {
                    originMap.put(row.get(keyCol).trim(), row);
                }
            }

            int updatedCount = 0;

            // --- Сравнение и обновление ---
            for (List<String> mainRow : mainData) {
                if (mainRow.size() <= keyCol) continue;
                String flight = mainRow.get(keyCol).trim();

                if (originMap.containsKey(flight)) {
                    List<String> originRow = originMap.get(flight);

                    String mainCar = safeGet(mainRow, carCol);
                    String mainDriver = safeGet(mainRow, driverCol);
                    String originCar = safeGet(originRow, carCol);
                    String originDriver = safeGet(originRow, driverCol);

                    boolean changed = false;

                    if (!mainCar.equalsIgnoreCase(originCar)) {
                        mainRow.set(carCol, originCar);
                        changed = true;
                    }

                    if (!mainDriver.equalsIgnoreCase(originDriver)) {
                        mainRow.set(driverCol, originDriver);
                        changed = true;
                    }

                    if (changed) updatedCount++;
                }
            }

            // --- Сохраняем результат ---
            File updatedFile = new File(mainFile.getParentFile(), "Main_updated.xlsx");
            writeExcel(mainData, updatedFile);

            JOptionPane.showMessageDialog(null,
                    "Готово!\nОбновлено строк: " + updatedCount +
                            "\nФайл сохранён: " + updatedFile.getAbsolutePath(),
                    "Excel Comparator", JOptionPane.INFORMATION_MESSAGE);

        } catch (Exception e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null, "Ошибка: " + e.getMessage());
        }
    }

    // ---- Вспомогательные методы ----
    private static List<List<String>> readExcel(File file) throws IOException {
        List<List<String>> rows = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(file);
             Workbook wb = new XSSFWorkbook(fis)) {
            Sheet sheet = wb.getSheetAt(0);
            for (Row row : sheet) {
                List<String> values = new ArrayList<>();
                for (int c = 0; c < row.getLastCellNum(); c++) {
                    Cell cell = row.getCell(c, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cell.setCellType(CellType.STRING);
                    values.add(cell.getStringCellValue().trim());
                }
                rows.add(values);
            }
        }
        return rows;
    }

    private static void writeExcel(List<List<String>> data, File file) throws IOException {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Updated");
            for (int i = 0; i < data.size(); i++) {
                Row row = sheet.createRow(i);
                for (int j = 0; j < data.get(i).size(); j++) {
                    row.createCell(j).setCellValue(data.get(i).get(j));
                }
            }
            try (FileOutputStream fos = new FileOutputStream(file)) {
                wb.write(fos);
            }
        }
    }

    private static String safeGet(List<String> row, int idx) {
        return (idx < row.size()) ? row.get(idx).trim() : "";
    }

}
