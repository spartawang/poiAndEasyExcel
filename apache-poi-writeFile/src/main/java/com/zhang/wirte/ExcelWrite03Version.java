package com.zhang.wirte;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.joda.time.DateTime;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @Author zhangxinrun(OS - > zhang)
 * @Date 2021/4/18 13:36
 * @Version 1.0
 * @Description 测试excel（03版）的写入操作，主要实现类为---->HSSFWorkbook
 */
public class ExcelWrite03Version {

    /**
     * 当前项目的绝对路径
     *  -   例如：D:\poiAndEasyExcel\apache-poi-writeFile\excelDirection
     */
    private static final String PATH = System.getProperty("user.dir") + "\\apache-poi-writeFile\\excelDirection\\";

    public static void main(String[] args) {

        // 1、常见新的Excel工作簿
        Workbook workbook = new HSSFWorkbook();

        // 2、在Excel工作簿中创建一张工作表，缺省名为Sheet0
        // Sheet sheet = workbook.createSheet();

        // 2、创建一个名为“写入的03版excel”的工作表
        Sheet sheet = workbook.createSheet("写入的03版excel");

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
            // 注意：03版本的excel文件名后缀为xls，最大存储65536行记录
            fileOutputStream = new FileOutputStream(PATH + "\\写入的03版excel.xls");
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
