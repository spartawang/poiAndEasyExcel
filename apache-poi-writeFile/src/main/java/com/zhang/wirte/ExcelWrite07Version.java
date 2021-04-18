package com.zhang.wirte;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @Author zhangxinrun(OS - > zhang)
 * @Date 2021/4/18 14:10
 * @Version 1.0
 * @Description 测试excel（07版）的写入操作，主要实现类为---->XSSFWorkbook
 */
public class ExcelWrite07Version {

    /**
     * 当前项目的绝对路径
     *  -   例如：D:\poiAndEasyExcel\apache-poi-writeFile\excelDirection
     */
    private static final String PATH = System.getProperty("user.dir") + "\\apache-poi-writeFile\\excelDirection\\";

    public static void main(String[] args) {

        // 1、创建工作簿，与03版本的区别就是实现类替换了
        Workbook workbook = new XSSFWorkbook();

        // 2、创建一个名为“写入的07版excel”的工作表
        Sheet sheet = workbook.createSheet("写入的07版excel");

        // 3、创建行
        // 第一行
        {
            Row row1 = sheet.createRow(0);

            // 4、创建单元格
            Cell cell11 = row1.createCell(0);
            cell11.setCellValue("序号");

            Cell cell12 = row1.createCell(1);
            cell12.setCellValue("姓名");

            Cell cell13 = row1.createCell(2);
            cell13.setCellValue("年龄");

            Cell cell14 = row1.createCell(3);
            cell14.setCellValue("生日");
        }

        // 第二行
        {
            Row row2 = sheet.createRow(1);

            Cell cell21 = row2.createCell(0);
            cell21.setCellValue("1");

            Cell cell22 = row2.createCell(1);
            cell22.setCellValue("xiaozhang");

            Cell cell23 = row2.createCell(2);
            cell23.setCellValue("21");

            Cell cell24 = row2.createCell(3);
            String dateTime = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
            cell24.setCellValue(dateTime);
        }

        // 5、新建一个输出文件流（注意：要先创建文件夹）
        FileOutputStream fileOutputStream = null;
        try {
            // 注意：07版本的excel文件名后缀为xlsx，无最大存储
            fileOutputStream = new FileOutputStream(PATH + "\\写入的07版excel.xlsx");
            workbook.write(fileOutputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                fileOutputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        System.out.println("您需要的文件已经生成成功，请查看！！！");
    }
}
