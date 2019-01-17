package com.fengwenyi.example.excelsimple;

import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * POI导出高版本的Excel
 * @author Wenyi Feng
 * @since 2018-11-24
 */
public class PoiExpExcelXslx {

    public static void main(String[] args) {

        String [] titles = {"id", "name", "sex"};

        // 创建Excel工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();

        // 创建一个工作表sheet
        XSSFSheet sheet = workbook.createSheet();

        // 创建第一行
        XSSFRow row = sheet.createRow(0);

        XSSFCell cell;

        // 插入第一行数据（标题）
        for (int i = 0; i < titles.length; i++) {
            cell = row.createCell(i);
            cell.setCellValue(titles[i]);
        }

        // 追加数据
        for (int i = 1; i < 11; i++) {
            XSSFRow nextRow = sheet.createRow(i);
            XSSFCell nextCell = nextRow.createCell(0);
            nextCell.setCellValue("a" + i);
            nextCell = nextRow.createCell(1);
            nextCell.setCellValue("b" + i);
            nextCell = nextRow.createCell(2);
            nextCell.setCellValue("c" + i);
        }

        // 创建一个文件
        File file = new File("d:/tmp/excel/poi_test.xlsx");
        try {
            file.createNewFile();

            // 将excel的内容写入到文件中
            FileOutputStream fos = FileUtils.openOutputStream(file);
            workbook.write(fos);
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
