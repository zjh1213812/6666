package com.javasm;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author:JAVASM
 * @classname:ExcelTest
 * @description:
 * @date:2023/1/10 20:04
 * @version:1.0
 * @since:11
 */
public class ExcelTest {
    @Test
    public void fun1() throws IOException {
        //创建工作簿对象
        Workbook workbook = new XSSFWorkbook();
        //创建页
        Sheet sheet = workbook.createSheet();
        //创建行
        //参数表示索引 从0开始
        Row row = sheet.createRow(3);
        //创建单元格
        //参数表示索引 从0开始
        Cell cell = row.createCell(3);
        //给单元格设置值
        cell.setCellValue("大厦北");
        //将工作簿写出本地
        workbook.write(new FileOutputStream("D:\\work.xlsx"));
        workbook.close();
    }

    @Test
    public void fun2() throws IOException {
        //创建工作簿对象
        Workbook workbook = new XSSFWorkbook("D:\\work.xlsx");
        //设置索引值 查询页 从0开始
        Sheet sheetAt = workbook.getSheetAt(0);
        //查找总行数
        int lastRowNum = sheetAt.getLastRowNum();
        for (int i = 0; i < lastRowNum; i++) {
            Row row = sheetAt.getRow(i);
            short lastCellNum = row.getLastCellNum();
            for (int i1 = 0; i1 < (int) lastCellNum; i1++) {
                Cell cell = row.getCell(i1);
                if (cell!=null) {
                    Object value = getValue(cell);
                    System.out.println(value);
                }
            }
        }
    }

    /**
     * 根据单元格获取单元格的值
     */
    public Object getValue(Cell cell){
        // 获取单元格的类型
        CellType cellType = CellType.forInt(cell.getCellType());
        Object obj = null;
        switch (cellType){
            case STRING:// 字符串类型
                obj = cell.getStringCellValue();
                break;
            case BOOLEAN:// 布尔类型
                obj = cell.getBooleanCellValue();
                break;
            case NUMERIC:// 数字和日期类型
                if (DateUtil.isCellDateFormatted(cell)){
                    // 如果是日期类型
                    obj = cell.getDateCellValue();
                }else {
                    // 数字类型
                    obj = cell.getNumericCellValue();
                }
                break;
            case FORMULA://如果是 公式 类型
                obj = cell.getCellFormula();
                break;
        }
        return obj;
    }
}
