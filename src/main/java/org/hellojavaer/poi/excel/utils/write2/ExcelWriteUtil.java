package org.hellojavaer.poi.excel.utils.write2;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;

/**
 * 大致初步的用法：
 * ExcelWriteFieldMapping fieldMapping = new ExcelWriteFieldMapping();
 * fieldMapping.put(fieldName).setHead(str);  //
 * ...
 * <p>
 * SheetProcessor sheetProcessor = new SheetProcessor();
 * sheetProcessor.setFieldMapping(fieldMapping);
 * sheetProcessor.set...
 * ...
 * sheetProcessor.setDataList(dataList);
 * <p>
 * WriteContext context = new WriteContext();
 * context.addSheetProcessor(sheetProcessor);
 * <p>
 * ExcelWriteUtil.write2(output, context); // or instead of context.write(output);
 * <p>
 * Created by luzy on 2017/10/17.
 */
public class ExcelWriteUtil {


    public static void write(OutputStream output, SheetProcessor... sheetProcessors) {
        Workbook workbook = new XSSFWorkbook();
        for (SheetProcessor sheetProcessor : sheetProcessors) {

            String sheetName = sheetProcessor.getSheetName();
            Integer sheetIndex = sheetProcessor.getSheetIndex();
            Sheet sheet = null;
            if (sheetName != null) {

            } else if (sheetIndex != null) {
                sheet = workbook.createSheet();
                workbook.setSheetOrder(sheet.getSheetName(), sheetIndex);
            }
            if (sheet == null) return; // or exception

            WriteContext context = new WriteContext();
            context.setCurSheet(sheet);
            context.setCurSheetIndex(sheetIndex);
            context.setCurSheetName(sheet.getSheetName());
            context.setCurRow(null);
            context.setCurRowIndex(null);
            context.setCurCell(null);
            context.setCurColIndex(null);

            sheetProcessor.setRowProcessor(new RowProcessor());
            sheetProcessor.process(context);
        }

        try {
            workbook.write(output);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
