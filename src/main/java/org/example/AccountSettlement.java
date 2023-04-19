package org.example;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class AccountSettlement {

//    public String filePath;
//
//    public AccountSettlement(String filePath) {
//        this.filePath = filePath;
//    }

    private XSSFSheet sheet;
    //    private String nameSupp;
    private HashMap<Character, Integer> columns = new HashMap<>();
    private HashMap<String, Double> states = new HashMap<>();
    private ArrayList<String> arrList = new ArrayList<>();

    // Add method that accepts Excel file as parameter | String readExcel()
    public void readExcel(String filePath, int sheet) {
        try {
            FileInputStream file = new FileInputStream(filePath);
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            this.sheet = workbook.getSheetAt(sheet - 1);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Get name of the supplier
    public void getNameSupp() {
        System.out.println("Enter name of the supplier: ");
        Scanner sc = new Scanner(System.in);
//        this.nameSupp = sc.nextLine();
    }

    // I can add 1-sheet, 2-columns | parameter
    public void getData() {
//        getNameSupp();
        for (Row r : sheet) {
            // Add method that accepts letter as number of column | char getColumn()
            Cell cStr = r.getCell(3);
            Cell cInt = r.getCell(4);
            Cell cData = r.getCell(5);
            Cell cStatus = r.getCell(2);

            if (cStr != null & cInt != null & cStatus.getStringCellValue().contains("Принят от поставщика")) {
                if (cStr.getCellType() == CellType.STRING & cInt.getCellType() == CellType.NUMERIC) {
                    if (states.containsKey(cStr.getStringCellValue())) {
//                        System.out.println(cStr.getStringCellValue() + " " + cInt.getNumericCellValue() + " | " + cData.getStringCellValue());
                        arrList.add(cStr.getStringCellValue());
                        states.put(cStr.getStringCellValue(), states.get(cStr.getStringCellValue()) + cInt.getNumericCellValue());
                    } else {
//                        System.out.println(cStr.getStringCellValue() + " " + cInt.getNumericCellValue() + " | " + cData.getStringCellValue());
                        states.put(cStr.getStringCellValue(), cInt.getNumericCellValue());
                    }
                } else if (cStr.getCellType() == CellType.FORMULA && cStr.getCachedFormulaResultType() == CellType.NUMERIC && cInt.getCellType() == CellType.FORMULA && cInt.getCachedFormulaResultType() == CellType.NUMERIC) {
                    states.put(cStr.getStringCellValue(), cInt.getNumericCellValue());
                }
            }
        }
    }

    public void getSum() {
        int size = 1;

        for (String elem : arrList) {
//            elem.contains("KENDALA IMPEX TOO") ? size++;
            size++;
//            if (elem.contains(nameSupp)) {
//                size++;
//            }
        }
        System.out.println("=========================================\n");
        System.out.println("Suppliers found by name: " + size);
        System.out.println("Sum: " + states);
    }

    public void createExcel() {
        try {

            Set<String> keys = states.keySet();
            Collection<Double> values = states.values();

//             Создание нового Excel-документа
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("Sheet 1");
            Row headerRow = sheet.createRow(0);

            headerRow.createCell(0).setCellValue("Контрагент");
            headerRow.createCell(1).setCellValue("Сумма");

            ArrayList<String> keyList = new ArrayList<>();
            ArrayList<Double> valuesList = new ArrayList<>();
            for (int i = 1; i <= keys.size(); i++) {
                keyList.add((String) keys.toArray()[i - 1]);
                valuesList.add((Double) values.toArray()[i - 1]);
                Row row = sheet.createRow(i);

                row.createCell(0).setCellValue(keyList.get(i - 1));
                row.createCell(1).setCellValue(valuesList.get(i - 1));
            }
            sheet.autoSizeColumn(0);
            sheet.autoSizeColumn(1);
//             Запись данных в файл
            FileOutputStream outputStream = new FileOutputStream("Результат.xlsx");
            workbook.write(outputStream);
            workbook.close();
            System.out.println("Результат записан в Excel-документ");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}


