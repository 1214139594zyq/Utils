package Utils.Poi;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.awt.Color;
import java.io.*;

public class PoiTest {
    public static void main(String[] args) {
        try {
            XSSFWorkbook xwb = new XSSFWorkbook();//这是创建表
            Sheet sheet = xwb.createSheet("测试poi");
            Row row = null;
            Cell cell = null;
            sheet.protectSheet("edit");
            XSSFCellStyle cellStyle5 = initColumnBorderStyle(xwb); // 设置边框
            XSSFCellStyle title1CellStyle = inittitle1CellStyle(xwb); // 第一行样式
            XSSFCellStyle cellStyle = inittitle2_3cellStyle(xwb); // 第二行和第三行样式
            XSSFCellStyle cellStyle2 = initMergeStyle(xwb); // 合并行的 样式
            XSSFCellStyle cellStyle3 = contentStyle(xwb);
            for (int i = 0; i < 20; i++) {//这里是先创建好所有的行
                row = sheet.createRow(i);
                for (int j = 0; j < 20; j++) {//指定了所有的列
                    cell = row.createCell(j);
                    cell.setCellStyle(cellStyle5);
                }
            }
            cell = sheet.getRow(0).getCell(0);//在已经建立好的单元格基础上获取到第一行第一个单元格
            cell.setCellValue("测试表格写入");
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 19));//第一行开始到第一行结束，第一个单元格开始到19单元格结束
            cell.setCellStyle(title1CellStyle);

            row = sheet.getRow(0);
            row.setHeight((short) (30 * 25));
            sheet.setColumnWidth(3, 256 * 30 + 184);
            sheet.setColumnWidth(4, 256 * 12 + 184);
            sheet.setColumnWidth(7, 256 * 12 + 184);
            sheet.setColumnWidth(13, 256 * 12 + 184);
            sheet.setColumnWidth(14, 256 * 12 + 184);
            sheet.setColumnWidth(17, 256 * 12 + 184);
            sheet.setColumnWidth(18, 256 * 12 + 184);

            cell = sheet.getRow(1).getCell(0);
            cell.setCellValue("到达");
            sheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 3));
            cell = sheet.getRow(1).createCell(4);
            cell.setCellValue("操作费");
            sheet.addMergedRegion(new CellRangeAddress(1, 2, 4, 4));
            cell = sheet.getRow(1).getCell(5);
            cell.setCellValue("干线");
            sheet.addMergedRegion(new CellRangeAddress(1, 2, 5, 11));
            cell = sheet.getRow(1).getCell(12);
            cell.setCellValue("派送");
            sheet.addMergedRegion(new CellRangeAddress(1, 1, 12, 19));

            row = sheet.getRow(2);
            row.createCell(12).setCellValue("送货");
            sheet.addMergedRegion(new CellRangeAddress(2, 2, 12, 15));
            row.createCell(16).setCellValue("自提");
            sheet.addMergedRegion(new CellRangeAddress(2, 2, 16, 19));
            XSSFCell nowCell1 = (XSSFCell) sheet.getRow(1).getCell(0);
            XSSFCell nowCell2 = (XSSFCell) sheet.getRow(1).getCell(4);
            XSSFCell nowCell3 = (XSSFCell) sheet.getRow(1).getCell(5);
            XSSFCell nowCell4 = (XSSFCell) sheet.getRow(1).getCell(12);
            XSSFCell nowCell5 = (XSSFCell) sheet.getRow(2).getCell(12);
            XSSFCell nowCell6 = (XSSFCell) sheet.getRow(2).getCell(16);
            nowCell1.setCellStyle(cellStyle);
            nowCell2.setCellStyle(cellStyle);
            nowCell3.setCellStyle(cellStyle);
            nowCell4.setCellStyle(cellStyle);
            nowCell5.setCellStyle(cellStyle);
            nowCell6.setCellStyle(cellStyle);

            // 创建表头 第四行
            row = sheet.getRow(3);
            for (int i = 0; i < 20; i++) {
                cell = sheet.getRow(3).getCell(i);
                cell.setCellStyle(cellStyle);
            }
            for (int z = 0; z < 3; z++) {
                row = sheet.createRow(16 + 4);
                row.createCell(z).setCellValue("");
            }


            FileOutputStream fs = new FileOutputStream(new File("D:/test.xlsx"));
            xwb.write(fs);
            fs.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    private static XSSFCellStyle contentStyle(XSSFWorkbook wb) {
        XSSFCellStyle cellStyle = wb.createCellStyle(); // 内容体样式
        cellStyle.setAlignment(HorizontalAlignment.RIGHT);
        cellStyle.setLeftBorderColor(new XSSFColor(new Color(0, 0, 0)));// 左边框的颜色
        cellStyle.setBorderLeft(BorderStyle.THIN);// 边框的大小
        cellStyle.setRightBorderColor(new XSSFColor(new Color(0, 0, 0)));// 右边框的颜色
        cellStyle.setBorderRight(BorderStyle.THIN);// 边框的大小
        cellStyle.setBorderBottom(BorderStyle.THIN); // 设置单元格的边框为粗体
        cellStyle.setBottomBorderColor(new XSSFColor(new Color(0, 0, 0))); // 设置单元格的边框颜色
        cellStyle.setLocked(true);
        return cellStyle;
    }


    private static XSSFCellStyle initMergeStyle(XSSFWorkbook wb) {
        // 设置合并行样式
        XSSFCellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.LEFT); // 靠左
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);// 垂直
        cellStyle.setLeftBorderColor(new XSSFColor(new Color(0, 0, 0)));// 左边框的颜色
        cellStyle.setBorderLeft(BorderStyle.THIN);// 边框的大小
        cellStyle.setRightBorderColor(new XSSFColor(new Color(0, 0, 0)));// 右边框的颜色
        cellStyle.setBorderRight(BorderStyle.THIN);// 边框的大小
        cellStyle.setBorderBottom(BorderStyle.THIN); // 设置单元格的边框为粗体
        cellStyle.setBottomBorderColor(new XSSFColor(new Color(0, 0, 0))); // 设置单元格的边框颜色
        cellStyle.setLocked(true);
        cellStyle.setLocked(true);
        return cellStyle;
    }


    private static XSSFCellStyle inittitle2_3cellStyle(XSSFWorkbook wb) {
        // 设置第二行到第四行样式
        XSSFCellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER); // 居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);// 垂直
        cellStyle.setFillForegroundColor(new XSSFColor(new Color(198, 226, 164)));
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setLeftBorderColor(new XSSFColor(new Color(0, 0, 0)));// 左边框的颜色
        cellStyle.setBorderLeft(BorderStyle.THIN);// 边框的大小
        cellStyle.setRightBorderColor(new XSSFColor(new Color(0, 0, 0)));// 右边框的颜色
        cellStyle.setBorderRight(BorderStyle.THIN);// 边框的大小
        cellStyle.setBorderBottom(BorderStyle.THIN); // 设置单元格的边框为粗体
        cellStyle.setBottomBorderColor(new XSSFColor(new Color(0, 0, 0))); // 设置单元格的边框颜色
        XSSFFont font1 = wb.createFont();
        font1.setBold(true); // 字体加粗
        cellStyle.setFont(font1);
        cellStyle.setLocked(true);
        return cellStyle;
    }


    private static XSSFCellStyle inittitle1CellStyle(XSSFWorkbook wb) {
        // 设置第一行样式
        XSSFCellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER); // 居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);// 垂直

        cellStyle.setFillForegroundColor(new XSSFColor(new Color(198, 226, 164)));
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        cellStyle.setLeftBorderColor(new XSSFColor(new Color(0, 0, 0)));// 左边框的颜色
        cellStyle.setBorderLeft(BorderStyle.THIN);// 边框的大小
        cellStyle.setRightBorderColor(new XSSFColor(new Color(0, 0, 0)));// 右边框的颜色
        cellStyle.setBorderRight(BorderStyle.THIN);// 边框的大小
        cellStyle.setBorderBottom(BorderStyle.THIN); // 设置单元格的边框为粗体
        cellStyle.setBottomBorderColor(new XSSFColor(new Color(0, 0, 0))); // 设置单元格的边框颜色

        XSSFFont font = wb.createFont();
        font.setFontHeightInPoints((short) 20);// 设置字体大小
        font.setBold(true); // 字体加粗
        cellStyle.setFont(font);
        cellStyle.setLocked(true);
        return cellStyle;
    }


    private static XSSFCellStyle initColumnBorderStyle(XSSFWorkbook wb) {
        // 设置边框
        XSSFCellStyle cellStyle = wb.createCellStyle();
        cellStyle.setLeftBorderColor(new XSSFColor(new Color(0, 0, 0)));// 左边框的颜色
        cellStyle.setBorderLeft(BorderStyle.THIN);// 边框的大小
        cellStyle.setRightBorderColor(new XSSFColor(new Color(0, 0, 0)));// 右边框的颜色
        cellStyle.setBorderRight(BorderStyle.THIN); // 边框的大小
        cellStyle.setBorderBottom(BorderStyle.THIN);// 设置单元格的边框为粗体
        cellStyle.setBottomBorderColor(new XSSFColor(new Color(0, 0, 0))); // 设置单元格的边框颜色
        return cellStyle;
    }


    private void mergeCell(int size, XSSFCellStyle cellStyle, XSSFSheet sheet, int a) {
        int currnetRow = 4;// 开始查找的行
        int totalRow = size + 4;
        for (int p = 4; p < totalRow; p++) {// totalRow 总行数
            XSSFCell currentCell = sheet.getRow(p).getCell(a);
            String current = getStringCellValue(currentCell);
            XSSFCell nextCell = null;
            String next = "";
            if (p < totalRow + 1) {
                XSSFRow nowRow = sheet.getRow(p + 1);
                if (nowRow != null) {
                    nextCell = nowRow.getCell(a);
                    next = getStringCellValue(nextCell);
                } else {
                    next = "";
                }
            } else {
                next = "";
            }
            if (current.equals(next)) {// 比对是否相同
                currentCell.setCellValue("");
                continue;
            } else if (!current.equals(next)) {
                if ((p) - currnetRow > 0 && current != "") {
                    sheet.addMergedRegion(new CellRangeAddress(currnetRow, p, a, a));// 合并单元格
                    XSSFCell nowCell = sheet.getRow(currnetRow).getCell(a);
                    nowCell.setCellValue(current);
                    nowCell.setCellStyle(cellStyle);
                }
                currnetRow = p + 1;
            }
        }
    }

    @SuppressWarnings({"deprecation", "unused"})
    private String getStringCellValue(XSSFCell cell) {
        String strCell = "";
        if (cell != null) {
            switch (cell.getCellType()) {
                case XSSFCell.CELL_TYPE_STRING:
                    strCell = cell.getStringCellValue();
                    break;
                case XSSFCell.CELL_TYPE_NUMERIC:
                    strCell = String.valueOf(cell.getNumericCellValue());
                    break;
                case XSSFCell.CELL_TYPE_BOOLEAN:
                    strCell = String.valueOf(cell.getBooleanCellValue());
                    break;
                case XSSFCell.CELL_TYPE_BLANK:
                    strCell = "";
                    break;
                default:
                    strCell = "";
                    break;
            }
            if (strCell.equals("") || strCell == null) {
                return "";
            }
            if (cell == null) {
                return "";
            }
        }
        return strCell;
    }
}
