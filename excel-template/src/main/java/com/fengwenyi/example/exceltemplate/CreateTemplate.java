package com.fengwenyi.example.exceltemplate;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.jdom.Attribute;
import org.jdom.Document;
import org.jdom.Element;
import org.jdom.JDOMException;
import org.jdom.input.SAXBuilder;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

/**
 * 创建模板文件
 * @author Wenyi Feng
 * @since 2018-11-24
 */
public class CreateTemplate {

    public static void main(String[] args) throws IOException {
        // 获取解析xml文件路径
        String path = System.getProperty("user.dir") + "/excel-template/src/main/resources/student.xml";
        File file = new File(path);
        SAXBuilder builder = new SAXBuilder();
        try {
            // 解析xml文件
            Document parse = builder.build(file);
            // 创建Excel
            HSSFWorkbook workbook = new HSSFWorkbook();
            // 创建sheet
            HSSFSheet sheet = workbook.createSheet("Sheet0");

            // 获取xml文件根节点
            Element root = parse.getRootElement();
            // 获取模板名称
            String templateName = root.getAttribute("name").getValue();

            int rowNum = 0;
            int columnNum = 0;

            // 设置列宽
            Element colGroup = root.getChild("colgroup");
            setColumnWidth(sheet, colGroup);

            // 设置标题
            Element title = root.getChild("title");
            List<Element> trs = title.getChildren("tr");
            for (int i = 0; i < trs.size(); i++) {
                Element tr = trs.get(i);
                List<Element> tds = tr.getChildren("td");
                HSSFRow row = sheet.createRow(rowNum);

                // 样式
                HSSFCellStyle cellStyle = workbook.createCellStyle();
                cellStyle.setAlignment(HorizontalAlignment.CENTER);

                for (columnNum = 0; columnNum < tds.size(); columnNum++) {
                    Element td = tds.get(columnNum);
                    HSSFCell cell = row.createCell(columnNum);
                    Attribute rowSpan = td.getAttribute("rowspan");
                    Attribute colSpan = td.getAttribute("colspan");
                    Attribute value = td.getAttribute("value");
                    if (value != null) {
                        String val = value.getValue();
                        cell.setCellValue(val);

                        int rspan = rowSpan.getIntValue() - 1;
                        int cspan = colSpan.getIntValue() - 1;

                        // 设置字体
                        HSSFFont font = workbook.createFont();
                        font.setFontName("仿宋_GB2312");
                        font.setBold(true);
                        //font.setFontHeight((short) 12);
                        font.setFontHeightInPoints((short) 12);
                        cellStyle.setFont(font);
                        cell.setCellStyle(cellStyle);

                        // 合并单元格
                        sheet.addMergedRegion(new CellRangeAddress(rspan, rspan, 0, cspan));
                    }
                }
                rowNum++;
            }

            // 设置表头信息
            Element thead = root.getChild("thead");
            trs = thead.getChildren("tr");
            for (int i = 0; i < trs.size(); i++) {
                Element tr = trs.get(i);
                HSSFRow row = sheet.createRow(rowNum);
                List<Element> ths = tr.getChildren("th");
                for (columnNum = 0; columnNum < ths.size(); columnNum++) {
                    Element th = ths.get(columnNum);
                    Attribute valueAttr = th.getAttribute("value");
                    HSSFCell cell = row.createCell(columnNum);
                    if (valueAttr != null) {
                        String value = valueAttr.getValue();
                        cell.setCellValue(value);
                    }
                }
                rowNum++;
            }

            // 设置数据区样式
            Element tbody = root.getChild("tbody");
            Element tr = tbody.getChild("tr");
            int repeat = tr.getAttribute("repeat").getIntValue();

            List<Element> tds = tr.getChildren("td");
            for (int i = 0; i < repeat; i++) {
                HSSFRow row = sheet.createRow(rowNum);
                for (columnNum = 0; columnNum < tds.size(); columnNum++) {
                    Element td = tds.get(columnNum);
                    HSSFCell cell = row.createCell(columnNum);
                    setType(workbook, cell, td);
                }
                rowNum++;
            }

            // 生成Excel导入模板
            File tempFile = new File("D:\\tmp\\excel/" + templateName + ".xls");
            tempFile.delete();
            tempFile.createNewFile();
            FileOutputStream stream = FileUtils.openOutputStream(tempFile);
            workbook.write(stream);
            stream.close();
            workbook.close();

        } catch (JDOMException e) {
            e.printStackTrace();
        }
    }

    /**
     * 设置单元格样式
     * @param workbook
     * @param cell
     * @param td
     */
    private static void setType(HSSFWorkbook workbook, HSSFCell cell, Element td) {
        Attribute typeAttr = td.getAttribute("type");
        String type = typeAttr.getValue();
        HSSFDataFormat format = workbook.createDataFormat();
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        if ("NUMERIC".equalsIgnoreCase(type)) {
            cell.setCellType(CellType.NUMERIC);
            Attribute formatAttr = td.getAttribute("format");
            String formatValue = formatAttr.getValue();
            formatValue = StringUtils.isNoneBlank(formatValue) ? formatValue : "#,##0.00";
            cellStyle.setDataFormat(format.getFormat(formatValue));
        } else if ("STRING".equalsIgnoreCase(type)) {
            cell.setCellValue("");
            cell.setCellType(CellType.STRING);
            cellStyle.setDataFormat(format.getFormat("@"));
        } else if ("DATE".equalsIgnoreCase(type)) {
            cell.setCellType(CellType.NUMERIC);
            cellStyle.setDataFormat(format.getFormat("yyyy-MM-dd"));
        } else if ("ENUM".equalsIgnoreCase(type)) {
            CellRangeAddressList regions = new CellRangeAddressList(cell.getRowIndex(), cell.getRowIndex(),
                    cell.getColumnIndex(), cell.getColumnIndex());
            Attribute enumAttr = td.getAttribute("format");
            String enumValue = enumAttr.getValue();
            // 加载下拉列表内容
            DVConstraint constraint = DVConstraint.createExplicitListConstraint(enumValue.split(","));
            // 数据有效性对象
            HSSFDataValidation dataValidation = new HSSFDataValidation(regions, constraint);
            workbook.getSheetAt(0).addValidationData(dataValidation);
            cell.setCellValue(enumValue.split(",")[0]);
        }
        cell.setCellStyle(cellStyle);
    }

    /**
     * 设置列宽
     * @param sheet
     * @param colGroup
     */
    private static void setColumnWidth(HSSFSheet sheet, Element colGroup) {
        List<Element> cols = colGroup.getChildren();
        for (int i = 0; i < cols.size(); i++) {
            Element col = cols.get(i);
            Attribute width = col.getAttribute("width");
            String unit = width.getValue().replaceAll("[0-9,\\.]", "");
            String value = width.getValue().replace(unit, "");
            int v = 0;
            if (StringUtils.isBlank(unit) || "px".endsWith(unit)) {
                v = Math.round(Float.parseFloat(value) * 37F);
            } else if ("em".equals(unit)) {
                v = Math.round(Float.parseFloat(value) * 267.5F);
            }
            sheet.setColumnWidth(i, v);
        }
    }

}
