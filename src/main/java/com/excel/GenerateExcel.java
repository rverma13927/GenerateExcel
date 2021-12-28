package com.excel;

import com.excel.Util.HeaderAndSetValueInCol;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.ParseException;
import java.util.List;

public class GenerateExcel {

    public ByteArrayOutputStream getExcel(String sheetName,String className, List<?> object, int size,String headers[],String methods[]) throws IOException, ParseException, ClassNotFoundException, NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        Workbook wb = new XSSFWorkbook();
        XSSFSheet sh = null;
        CellStyle boldStyle = wb.createCellStyle();
        boldStyle.setLocked(false);
        sh = (XSSFSheet) HeaderAndSetValueInCol.setHeaders(sheetName, headers, wb);
        int col = 0;
        Class<?> c = Class.forName(className);
        for (int i = 0; i < size; i++) {
            Row row = sh.createRow(i + 1);
            col = 0;
            Object obj = object.get(i);
            for(int j=0;j<headers.length;j++) {
                Method method = c.getDeclaredMethod(methods[j]);
                HeaderAndSetValueInCol.createCellAndSetValue(row, col++, String.valueOf(method.invoke(c.cast(obj))), null);
            }
        }
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        wb.write(out);
        return out;
    }

}
