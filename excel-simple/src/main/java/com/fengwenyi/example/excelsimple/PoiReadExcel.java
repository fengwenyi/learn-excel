package com.fengwenyi.example.excelsimple;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.IOException;

/**
 * POI导入Excel文件
 * @author Wenyi Feng
 * @since 2018-11-24
 */
public class PoiReadExcel {

    public static void main(String[] args) {
        // 需要解析的Excel文件
        File file = new File("d:/tmp/excel/poi_test.xls");
        try {
            // 创建Excel，读取文件内容
            HSSFWorkbook workbook = new HSSFWorkbook(FileUtils.openInputStream(file));

            // 获取第一个工作表
            // HSSFSheet sheet = workbook.getSheet("Sheet0");
            HSSFSheet sheet = workbook.getSheetAt(0);
            int firstRowNum = 0;
            // 获取sheet最后一行的行号
            int lastRowNum = sheet.getLastRowNum();
            for (int i = firstRowNum; i <= lastRowNum; i++) {
                HSSFRow row = sheet.getRow(i);
                // 获取当前行最后单元格列号
                int lastCellNum = row.getLastCellNum();
                for (int j = 0; j < lastCellNum; j++) {
                    HSSFCell cell = row.getCell(j);
                    String value = cell.getStringCellValue();
                    System.out.print(value + " ");
                }
                System.out.println();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
