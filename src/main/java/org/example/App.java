package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class App
{
    public static void main( String[] args ) throws Exception {
        String file1 = "C:\\Users\\arsen\\OneDrive\\Рабочий стол\\Origin.xlsx";
        String file2 = "C:\\Users\\arsen\\OneDrive\\Рабочий стол\\Main.xlsx";
        String outputFile = "C:\\Users\\arsen\\OneDrive\\Рабочий стол\\work_updated.xlsx";

        // 🔹 Читаем оба файла
        List<List<String>> workData = readExcel(file1);
        List<List<String>> mainData = readExcel(file2);

        // 🔹 Укажи индексы столбцов (0 — первый столбец)
        int flightCol = 0; // номер рейса
        int carCol = 1;    // гос номер машины
        int driverCol = 2; // водитель

        // 🔹 Создаём карту рейсов из main.xlsx
        Map<String, List<String>> mainMap = new HashMap<>();
        for (List<String> row : mainData) {
            if (row.size() > flightCol) {
                String flight = row.get(flightCol).trim();
                mainMap.put(flight, row);
            }
        }

        // 🔹 Обновляем work.xlsx данными из main.xlsx
        for (List<String> workRow : workData) {
            if (workRow.size() > flightCol) {
                String flight = workRow.get(flightCol).trim();
                if (mainMap.containsKey(flight)) {
                    List<String> mainRow = mainMap.get(flight);

                    // если госномер или водитель отличаются — обновляем
                    if (mainRow.size() > carCol && mainRow.size() > driverCol) {
                        String mainCar = mainRow.get(carCol);
                        String mainDriver = mainRow.get(driverCol);

                        if (workRow.size() > carCol) workRow.set(carCol, mainCar);
                        if (workRow.size() > driverCol) workRow.set(driverCol, mainDriver);
                    }
                }
            }
        }

        // 🔹 Записываем обновлённый файл
        writeExcel(workData, outputFile);

        System.out.println("✅ Файл обновлён: " + new File(outputFile).getAbsolutePath());
    }

    // === ЧТЕНИЕ ===
    private static List<List<String>> readExcel(String filePath) throws IOException {
        List<List<String>> rows = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                List<String> rowData = new ArrayList<>();
                for (Cell cell : row) {
                    cell.setCellType(CellType.STRING);
                    rowData.add(cell.getStringCellValue().trim());
                }
                rows.add(rowData);
            }
        }
        return rows;
    }

    // === ЗАПИСЬ ===
    private static void writeExcel(List<List<String>> data, String filePath) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Updated");
            for (int i = 0; i < data.size(); i++) {
                Row row = sheet.createRow(i);
                List<String> rowData = data.get(i);
                for (int j = 0; j < rowData.size(); j++) {
                    row.createCell(j).setCellValue(rowData.get(j));
                }
            }
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }
        }
    }

}
