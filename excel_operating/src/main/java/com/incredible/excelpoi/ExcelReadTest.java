package com.incredible.excelpoi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ExcelReadTest {

    public static void main(String[] args) throws IOException {
        readExcel03Test();
        readExcel07Test();
    }

    private static void readExcel03Test() throws IOException {

        FileInputStream fis = new FileInputStream("test03.xls");
        // 读取xls文件为workboot对象
        Workbook workbook = new HSSFWorkbook(fis);
        // 得到第一个sheet页
        Sheet sheet = workbook.getSheetAt(0);
        // 得到第一行
        Row row = sheet.getRow(0);

        Cell cell = row.getCell(0);
        // 读取值的时候注意值的类型
        System.out.println("03版本："+cell.getStringCellValue());
        workbook.close();
        fis.close();

    }

    private static void readExcel07Test() throws IOException {

        FileInputStream fis = new FileInputStream("test07.xlsx");
        // 读取xls文件为workboot对象
        Workbook workbook = new XSSFWorkbook(fis);
        // 得到第一个sheet页
        Sheet sheet = workbook.getSheetAt(0);
        // 得到第一行
        Row row = sheet.getRow(0);

        Cell cell = row.getCell(0);
        // 读取值的时候注意值的类型
        System.out.println("07版本:"+cell.getStringCellValue());
        workbook.close();
        fis.close();

    }

}
