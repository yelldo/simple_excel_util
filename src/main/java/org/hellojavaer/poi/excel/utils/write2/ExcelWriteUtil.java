package org.hellojavaer.poi.excel.utils.write2;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.OutputStream;

/**
 * 大致初步的用法：
 * WriteFieldMapping fieldMapping = new WriteFieldMapping();
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

    public static void write(OutputStream output, SheetProcessor sheetProcessor) {
        Workbook workbook = new XSSFWorkbook();
        Integer sheetIndex = sheetProcessor.getSheetIndex();
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        if (sheet == null) return; // or exception
        WriteContext context = new WriteContext();
        context.setCurSheet(sheet);
        context.setCurSheetIndex(sheetIndex);
        context.setCurSheetName(sheet.getSheetName());
        context.setCurRow(null);
        context.setCurRowIndex(null);
        context.setCurCell(null);
        context.setCurColIndex(null);
        sheetProcessor.setWriteContext(context);

        // sheet process
        /*sheetProcessor.setSheet(sheet).process(new RowProcessor() {
            @Override
            void process(WriteProcessor processor, WriteContext context) {
                // to do something
            }
        }, context, workbook);*/
        RowProcessor rowProcessor = null;
        WriteProcessor writeProcessor = sheetProcessor.getWriteProcessor();
        if (writeProcessor != null && writeProcessor instanceof RowProcessor) {
            rowProcessor = (RowProcessor) writeProcessor;
        } else {
            rowProcessor = new RowProcessor();
        }
        sheetProcessor.setSheet(sheet).process(rowProcessor);


    }

    public static void write2(OutputStream output, WriteContext context) {
        // use this method finally !!!
    }
}
