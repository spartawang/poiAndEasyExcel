package com.zhang.read;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

/**
 * @Author zhangxinrun(OS - > zhang)
 * @Date 2021/4/18 14:38
 * @Version 1.0
 * @Description 03版本读取
 */
public class ExcelRead03Version {

    /**
     * 当前项目的绝对路径
     *  -   例如：D:\poiAndEasyExcel\apache-poi-writeFile\excelDirection
     */
    private static final String PATH = System.getProperty("user.dir") + "\\apache-poi-writeFile\\excelDirection\\";

    public static void main(String[] args) {

        FileInputStream fileInputStream = null;
        try {
            fileInputStream = new FileInputStream(PATH + "写入的03版excel.xls");

            Workbook workbook = new HSSFWorkbook(fileInputStream);
            // 读取第一张工作簿
            Sheet sheetAt = workbook.getSheetAt(0);

            // 读取第一行第一列
            Row row = sheetAt.getRow(0);
            Cell cell = row.getCell(0);

            Cell cell1 = row.getCell(3);

            System.out.println(cell.getStringCellValue());
            System.out.println(cell1.getStringCellValue());

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                // 关闭流，释放资源
                fileInputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
