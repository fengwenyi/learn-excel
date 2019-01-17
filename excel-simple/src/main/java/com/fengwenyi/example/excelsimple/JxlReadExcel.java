package com.fengwenyi.example.excelsimple;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import java.io.File;
import java.io.IOException;

/**
 * JXL解析Excel
 * @author Wenyi Feng
 * @since 2018-11-24
 */
public class JxlReadExcel {

    public static void main(String[] args) throws BiffException {
        try {
            // 创建Workbook
            Workbook workbook = Workbook.getWorkbook(new File("d:/tmp/excel/jxl_test.xls"));

            // 获取工作表sheet
            Sheet sheet = workbook.getSheet(0);

            // 获取数据
            for (int i = 0; i < sheet.getRows(); i++) {
                for (int j = 0; j < sheet.getColumns(); j++) {
                    Cell cell = sheet.getCell(j, i);
                    System.out.print(cell.getContents() + " ");
                }
                System.out.println();
            }
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
