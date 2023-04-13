package org.example;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class AccountSettlement {
    public static void main(String[] args) {

        // Reading Old Excel

        try {
            // Чтение данных из файла
            FileInputStream file = new FileInputStream("D:\\JavaProjects\\AccountSettlement\\src\\main\\java\\org\\example\\rawExcelFile.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(0);
            // XSSFRow row = sheet.getRow(0);
            // XSSFCell cell = row.getCell(0);
            // String value = cell.getStringCellValue();

            // Чтение значения ячейки
            Map<String, Double> states = new HashMap<String, Double>();
            ArrayList<String> arrList = new ArrayList<>();
            for(Row r : sheet) {
                Cell cStr = r.getCell(3);
                Cell cInt = r.getCell(4);
                if(cStr != null & cInt != null) {
                    if(cStr.getCellType() == CellType.STRING & cInt.getCellType() == CellType.NUMERIC) {
                        if (states.containsKey(cStr.getStringCellValue())) {
                            arrList.add(cStr.getStringCellValue());
                            states.put(cStr.getStringCellValue(),states.get(cStr.getStringCellValue()) + cInt.getNumericCellValue());
                        } else {
                            states.put(cStr.getStringCellValue(), cInt.getNumericCellValue());
                        }
                    } else if(cStr.getCellType() == CellType.FORMULA && cStr.getCachedFormulaResultType() == CellType.NUMERIC && cInt.getCellType() == CellType.FORMULA && cInt.getCachedFormulaResultType() == CellType.NUMERIC) {
//                        values.add(c.getNumericCellValue());
                        states.put(cStr.getStringCellValue(), cInt.getNumericCellValue());
                    }
                }
            }

            int size = 0;

            for (String elem : arrList) {
//                elem.contains("KENDALA IMPEX TOO") ? size++;
                if (elem.contains("KENDALA IMPEX ТОО")) {
                    size++;
                }
            }
            System.out.println("Suppliers found by name: " + size);
            System.out.println("Sum: " + states.get("KENDALA IMPEX ТОО"));
            // Вывод данных в консоль
//            System.out.println("Значение ячейки A1: " + value);
//            System.out.println(values);
            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        // Creating New Excel
//        try {
//            // Создание нового Excel-документа
//            XSSFWorkbook workbook = new XSSFWorkbook();
//            XSSFSheet sheet = workbook.createSheet("Новый лист");
//            // Создание формулы в ячейке B1
//            XSSFRow row = sheet.createRow(0);
//            XSSFCell cellA1 = row.createCell(0);
//            XSSFCell cellB1 = row.createCell(1);
//            cellA1.setCellValue(10);
//            cellB1.setCellFormula("A1*2");
//            // Запись данных в файл
//            FileOutputStream outputStream = new FileOutputStream("Результат.xlsx");
//            workbook.write(outputStream);
//            workbook.close();
//            System.out.println("Результат записан в Excel-документ");
//        } catch (Exception e) {
//            e.printStackTrace();
//        }

    }
}

