package org.example;

import org.apache.poi.xssf.usermodel.XSSFSheet;

public class Main {

    public static void main(String[] args) {
        AccountSettlement n = new AccountSettlement();
        n.readExcel("D:\\JavaProjects\\AccountSettlement\\src\\main\\java\\org\\example\\DOC-20230320-WA0015.xlsx", 1);
        n.getData();
        n.getSum();
    }

}
