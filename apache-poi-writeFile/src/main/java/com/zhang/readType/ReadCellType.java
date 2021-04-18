package com.zhang.readType;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.joda.time.DateTime;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Date;

/**
 * @Author zhangxinrun(OS - > zhang)
 * @Date 2021/4/18 14:55
 * @Version 1.0
 * @Description 读取商品类型数据
 */
public class ReadCellType {
    /**
     * 当前项目的绝对路径
     *  -   例如：D:\poiAndEasyExcel\apache-poi-writeFile\excelDirection
     */
    private static final String PATH = System.getProperty("user.dir") + "\\apache-poi-writeFile\\excelDirection\\";

    public static void main(String[] args) {
        FileInputStream fileInputStream = null;
        try {
            fileInputStream = new FileInputStream(PATH + "会员消费商品明细表.xls");

            Workbook workbook = new HSSFWorkbook(fileInputStream);

            Sheet sheetAt = workbook.getSheetAt(0);

            // 读取标题所有内容
            Row rowTitle = sheetAt.getRow(0);
            // 判断行不为空
            if(rowTitle != null){
                // 读取cell
                int cellCount = rowTitle.getPhysicalNumberOfCells();
                for (int cellNumber = 0; cellNumber < cellCount; cellNumber++) {
                    Cell cell = rowTitle.getCell(cellNumber);
                    if(cell != null){
                        String cellValue = cell.getStringCellValue();
                        System.out.print(cellValue + "|");
                    }
                }
                System.out.println();
            }

            // 读取商品列表信息
            int rowCount = sheetAt.getPhysicalNumberOfRows();
            if(rowCount > 1){
                for (int rowNumber = 0; rowNumber < rowCount; rowNumber++) {
                    Row rowData = sheetAt.getRow(rowNumber);
                    if(rowData != null){
                        int cellCount = rowTitle.getPhysicalNumberOfCells();
                        for (int cellNumber = 0; cellNumber < cellCount; cellNumber++) {
                            Cell cell = rowData.getCell(cellNumber);
                            if(cell != null){
                                int cellType = cell.getCellType();
                                // 判断单元格类型
                                String cellValue = "";
                                switch (cellType){
                                    case HSSFCell.CELL_TYPE_STRING: // 字符串
                                        System.out.print("【STRING】");
                                        cellValue = cell.getStringCellValue();
                                        break;
                                    case HSSFCell.CELL_TYPE_BOOLEAN: // 布尔类型
                                        System.out.print("【BOOLEAN】");
                                        cellValue = String.valueOf(cell.getBooleanCellValue());
                                        break;
                                    case HSSFCell.CELL_TYPE_BLANK: // 空
                                        System.out.print("【BLANK】");
                                        break;
                                    case HSSFCell.CELL_TYPE_NUMERIC: // 数字（纯数字 or 日期）
                                        System.out.print("【NUMERIC】");
                                        // 日期类型
                                        if(HSSFDateUtil.isCellDateFormatted(cell)){
                                            System.out.print("DATE");
                                            Date date = cell.getDateCellValue();
                                            cellValue = new DateTime(date).toString("yyyy-MM-dd");
                                        } else {
                                            // 不是日期类型。则防止当数字过长时以科学计数法显示
                                            System.out.print("NUMBER");
                                            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                            cellValue = cell.toString();
                                        }
                                        break;
                                    case HSSFCell.CELL_TYPE_ERROR: // 错误
                                        System.out.println("【ERROR】");
                                        break;
                                }
                                System.out.println(cellCount);
                            }
                        }
                    }
                }
            }
            System.out.println("finish...");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                fileInputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
