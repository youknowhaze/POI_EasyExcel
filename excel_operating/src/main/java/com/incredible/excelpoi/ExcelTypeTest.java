package com.incredible.excelpoi;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;

import static com.sun.org.apache.xalan.internal.xsltc.compiler.Constants.STRING;

/**
 * excel的分类型读写
 */
public class ExcelTypeTest {
    public static void main(String[] args) throws IOException {
        excelType03Test();
    }

    private static void excelType03Test() throws IOException {

        FileInputStream fis = new FileInputStream("test03.xls");

        Workbook workbook = new HSSFWorkbook(fis);

        Sheet sheet = workbook.getSheetAt(0);
        // 获取标题行
        Row titleRow = sheet.getRow(0);
        //得到这行的cell数量
        int cellNum = titleRow.getPhysicalNumberOfCells();
        //变量标题行
        for (int i = 0; i < cellNum; i++) {
            Cell cell = titleRow.getCell(i);
            if (cell != null){
                System.out.print(cell.getStringCellValue() + "  |  ");
            }
        }

        // 得到row有多少行
        int rowNum = sheet.getPhysicalNumberOfRows();
        // 从第二行开始，第一行是标题
        for (int i = 1; i < rowNum; i++) {
            Row valueRow = sheet.getRow(i);
            int valueCellNum = valueRow.getPhysicalNumberOfCells();
            for (int j = 0; j < valueCellNum; j++) {
                Cell valueCell = valueRow.getCell(j);
                if (valueCell != null){
                    // 得到数据的类型
                   CellType cellType = valueCell.getCellType();
                   String cellValue = "";
                   switch (cellType){
                       case _NONE:
                           break;
                       case NUMERIC:
                           System.out.printf("【Number】:");
                           if (DateUtil.isCellDateFormatted(valueCell)){
                               DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
                               cellValue = dateFormat.format(valueCell.getDateCellValue());
                           }else {
                               System.out.println("数字转换为字符类型");
                               valueCell.setCellFormula(STRING);
                               cellValue = String.valueOf(valueCell.getNumericCellValue());
                           }
                           break;
                       case STRING :
                           System.out.printf("【String】:");
                           cellValue = valueCell.getStringCellValue();
                           break;
                       case FORMULA:
                           System.out.printf("【FORMULA】:");
                           break;
                       case BLANK:
                           System.out.printf("【BLANK】:");
                           break;
                       case BOOLEAN:
                           System.out.printf("【BOOLEAN】:");
                           cellValue = String.valueOf(valueCell.getBooleanCellValue());
                           break;
                       case ERROR:
                           System.out.printf("【ERROR】:");
                           break;
                   }
                    System.out.println(" ----- "+cellValue);

                }
            }

        }


    }
}
