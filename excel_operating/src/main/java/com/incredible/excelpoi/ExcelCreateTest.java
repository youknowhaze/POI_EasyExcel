package com.incredible.excelpoi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 * 1、03版本对应xls对应HSSFWorkbook
 * 2/07版本对应xlsx对应XSSFWorkbook
 */
import java.io.FileOutputStream;

public class ExcelCreateTest {

    public static void main(String[] args) throws Exception{

        excel03Test();

        excel07Test();

    }


    private static void excel03Test() throws Exception{
        // 创建工作表
        HSSFWorkbook workbook = new HSSFWorkbook();
        // 创建sheet页
        Sheet sheet = workbook.createSheet("测试03版本");
        // 创建第一行数据
        Row row = sheet.createRow(0);
        // 第一行创建10个cell并插入数据
        for (int i = 0; i < 10; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue("测试03第一行cell："+i);
        }

        Row row1 = sheet.createRow(1);
        for (int i = 0; i < 10; i++) {
            Cell cell = row1.createCell(i);
            cell.setCellValue("测试03第二行cell："+i);
        }

        // 03版本的excel为xsl格式文件
        FileOutputStream fileOutputStream = new FileOutputStream("test03.xls");

        workbook.write(fileOutputStream);
        fileOutputStream.close();
        workbook.close();
    }

    private static void excel07Test() throws Exception{
        // 创建工作表
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 创建sheet页
        Sheet sheet = workbook.createSheet("测试07版本");
        // 创建第一行数据
        Row row = sheet.createRow(0);
        // 第一行创建10个cell并插入数据
        for (int i = 0; i < 10; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue("测试07第一行cell："+i);
        }

        Row row1 = sheet.createRow(1);
        for (int i = 0; i < 10; i++) {
            Cell cell = row1.createCell(i);
            cell.setCellValue("测试07第二行cell："+i);
        }

        // 03版本的excel为xsl格式文件
        FileOutputStream fileOutputStream = new FileOutputStream("test07.xlsx");

        workbook.write(fileOutputStream);
        fileOutputStream.close();
        workbook.close();
    }

    /**
     * 07版本写入大量数据会比较慢，所以提供了SXSSFWorkbook来加快速度，
     * 这个操作会产生临时文件，所以在操作结束后需要关闭临时文件、
     * ((SXSSFWorkbook) workbook).dispose();
     * @throws Exception
     */
    private static void excel07BigDataTest() throws Exception{
        // 创建工作表
        Workbook workbook = new SXSSFWorkbook();
        // 创建sheet页
        Sheet sheet = workbook.createSheet("测试07版本");
        // 创建第一行数据
        Row row = sheet.createRow(0);
        // 第一行创建10个cell并插入数据
        for (int i = 0; i < 10; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue("测试07第一行cell："+i);
        }

        Row row1 = sheet.createRow(1);
        for (int i = 0; i < 10; i++) {
            Cell cell = row1.createCell(i);
            cell.setCellValue("测试07第二行cell："+i);
        }

        // 03版本的excel为xsl格式文件
        FileOutputStream fileOutputStream = new FileOutputStream("test07.xlsx");

        workbook.write(fileOutputStream);
        fileOutputStream.close();
        workbook.close();
        ((SXSSFWorkbook) workbook).dispose();
    }

}
