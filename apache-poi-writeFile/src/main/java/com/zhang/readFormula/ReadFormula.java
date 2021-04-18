package com.zhang.readFormula;

import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

/**
 * @Author zhangxinrun(OS - > zhang)
 * @Date 2021/4/18 15:19
 * @Version 1.0
 * @Description 测试计算公式
 */
public class ReadFormula {

    /**
     * 当前项目的绝对路径
     *  -   例如：D:\poiAndEasyExcel\apache-poi-writeFile\excelDirection
     */
    private static final String PATH = System.getProperty("user.dir") + "\\apache-poi-writeFile\\excelDirection\\";

    public static void main(String[] args) {

        InputStream inputStream = null;
        try {
            inputStream = new FileInputStream(PATH + "计算公式.xls");

            Workbook workbook = new HSSFWorkbook(inputStream);
            Sheet sheet = workbook.getSheetAt(0);

            // 获取第五行的第一列
            Row row = sheet.getRow(4);
            Cell cell = row.getCell(0);

            // 公式计算器
            FormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook) workbook);

            // 输出单元格内容
            int cellType = cell.getCellType();
            switch (cellType) {
                case Cell.CELL_TYPE_FORMULA:
                    String formula = cell.getCellFormula();
                    System.out.println("计算公式是：" + formula);
                    CellValue evaluate = formulaEvaluator.evaluate(cell);
                    String cellValue = evaluate.formatAsString();
                    System.out.println("值为：" + cellValue);
                    break;
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }
}
