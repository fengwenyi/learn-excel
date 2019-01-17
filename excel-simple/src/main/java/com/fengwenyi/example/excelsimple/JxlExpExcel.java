package com.fengwenyi.example.excelsimple;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import java.io.File;

/**
 * JXL导出Excel
 * @author Wenyi Feng
 * @since 2018-11-24
 */
public class JxlExpExcel {

    public static void main(String[] args) {
        String [] titles = {"id", "name", "sex"};
        File file = new File("d:/tmp/excel/jxl_test.xls");
        try {
            file.createNewFile();
            // 创建工作簿
            WritableWorkbook workbook = Workbook.createWorkbook(file);
            // 创建sheet
            WritableSheet sheet = workbook.createSheet("sheet1", 0);
            Label label;

            // 第一行设置列名
            for (int i = 0; i < titles.length; i++) {
                label = new Label(i, 0, titles[i]);
                sheet.addCell(label);
            }

            // 设置数据
            for (int i = 1; i < 11; i++) {
                label = new Label(0, i, "a" + i);
                sheet.addCell(label);
                label = new Label(1, i, "b" + i);
                sheet.addCell(label);
                label = new Label(2, i, "c" + i);
                sheet.addCell(label);
            }

            // 写入数据
            workbook.write();

            // 关闭流
            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
