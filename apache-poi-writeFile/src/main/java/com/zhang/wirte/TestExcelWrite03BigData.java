package com.zhang.wirte;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @Author zhangxinrun(OS - > zhang)
 * @Date 2021/4/18 14:15
 * @Version 1.0
 * @Description 测试03版插入大数据量时的消耗时间
 */
public class TestExcelWrite03BigData {

    /**
     * 当前项目的绝对路径
     *  -   例如：D:\poiAndEasyExcel\apache-poi-writeFile\excelDirection
     */
    private static final String PATH = System.getProperty("user.dir") + "\\apache-poi-writeFile\\excelDirection\\";

    public static void main(String[] args) {
        // 记录开始时间
        long begin = System.currentTimeMillis();

        Workbook workbook = new HSSFWorkbook();

        Sheet sheet = workbook.createSheet("03版大数据量写入");

        //xls文件最大支持65536行
        for(int rowNumber = 0; rowNumber < 65536; rowNumber++){
            Row row = sheet.createRow(rowNumber);
            for(int cellNumber = 0; cellNumber < 10; cellNumber++){
                Cell cell = row.createCell(cellNumber);
                cell.setCellValue(cellNumber);
            }
        }

        System.out.println("done");

        FileOutputStream fileOutputStream = null;
        try {
            fileOutputStream = new FileOutputStream(PATH + "\\03版大数据量写入.xls" );
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

        System.out.println("写入完成，请查看！！！");

        // 记录结束时间
        long end = System.currentTimeMillis();
        // 消耗总时长
        System.out.println((double)(end - begin) / 1000 + "秒；" + (end - begin) + "毫秒");
    }
}
