package com.excel.Util;


import org.apache.poi.ss.usermodel.*;

public class HeaderAndSetValueInCol {

    public static  void createCellAndSetValue(Row row, int col, String value, CellStyle boldStyle) {
        Cell cell = row.createCell(col);
        cell.setCellValue(value==null || value.equals("") ? "" :value);
        if(boldStyle!=null)
            cell.setCellStyle(boldStyle);
    }
    public static  Sheet setHeaders(String sheetName , String[] headers, Workbook wb) {
        Sheet sh = wb.createSheet(sheetName);

        CellStyle fillStyle = wb.createCellStyle();
        fillStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        fillStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        CellStyle boldStyle = wb.createCellStyle();
        int rownum = 0;
        Row row = sh.createRow(rownum);
        int pos = 0;
        for (int i = 0; i < headers.length; i++) {
            sh.setColumnWidth(pos, 25 * 256);
            Cell cell = row.createCell(pos);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(boldStyle);
            pos++;
        }
        return sh;
    }
}
