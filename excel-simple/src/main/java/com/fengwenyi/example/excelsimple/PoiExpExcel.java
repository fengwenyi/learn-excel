package com.fengwenyi.example.excelsimple;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * POI导出Excel
 * @author Wenyi Feng
 * @since 2018-11-24
 */
public class PoiExpExcel {

    public static void main(String[] args) {

        String [] titles = {"id", "name", "sex"};

        // 创建Excel工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();

        // 创建一个工作表sheet
        HSSFSheet sheet = workbook.createSheet();

        // 创建第一行
        HSSFRow row = sheet.createRow(0);

        HSSFCell cell;

        // 插入第一行数据（标题）
        for (int i = 0; i < titles.length; i++) {
            cell = row.createCell(i);
            cell.setCellValue(titles[i]);
        }

        // 追加数据
        for (int i = 1; i < 11; i++) {
            HSSFRow nextRow = sheet.createRow(i);
            HSSFCell nextCell = nextRow.createCell(0);
            nextCell.setCellValue("a" + i);
            nextCell = nextRow.createCell(1);
            nextCell.setCellValue("b" + i);
            nextCell = nextRow.createCell(2);
            nextCell.setCellValue("c" + i);
        }

        // 创建一个文件
        File file = new File("d:/tmp/excel/poi_test.xls");
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
