package org.hellojavaer.poi.excel.utils.write2;

import lombok.Data;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.List;
import java.util.Map;

/**
 * sheet 工作表处理器
 *
 * Created by luzy on 2017/10/17.
 */
@Data
public class SheetProcessor extends WriteProcessor {

    private Sheet                          sheet;
    private Integer                       sheetIndex;
    private String                        sheetName;
    private int                            startRowIndex = 0;
    private Integer                       templateStartRowIndex;
    private Integer                       templateEndRowIndex;
    private Integer                       headRowIndex;
    private Integer                       theme;
    private WriteFieldMapping            fieldMapping;
    private List<T>                        dataList;
    //private boolean                       trimSpace     = false;

    public SheetProcessor setSheet(Sheet sheet){
        this.sheet = sheet;
        return this;
    }

    @Override
    public void process(WriteProcessor rowProcessor) {
        this.writeProcessor = rowProcessor;
        writeHead();
        writeContent();
    }

    public void writeContent(){
        int writeRowIndex = startRowIndex;
        for (Object rowData : dataList) {
            if(rowData == null) continue;
            Row row = sheet.getRow(writeRowIndex);
            if(row == null) row = sheet.createRow(writeRowIndex);

            writeContext.setCurRow(row);
            writeContext.setCurRowIndex(writeRowIndex);
            writeContext.setCurCell(null);
            writeContext.setCurColIndex(null);

            if (this.writeProcessor == null) continue; // exception

            CellProcessor cellProcessor = null;
            RowProcessor rowProcessor = (RowProcessor) writeProcessor;
            WriteProcessor processor = rowProcessor.getWriteProcessor();
            if (processor != null && processor instanceof CellProcessor) {
                cellProcessor = (CellProcessor) writeProcessor;
            }else{
                cellProcessor = new CellProcessor();
            }
            rowProcessor.process(cellProcessor);
            writeRowIndex++;

            // 未完待续 。。。
        }
    }

    public void writeHead(){
        if(headRowIndex == null) return; // or exception
        Row row = sheet.getRow(headRowIndex);
        if (row == null) {
            row = sheet.createRow(headRowIndex);
        }
        for (Map.Entry<String, Map<Integer, WriteFieldMapping.ValueAttribute>> entry : getFieldMapping().export().entrySet()) {
            Map<Integer, WriteFieldMapping.ValueAttribute> map = entry.getValue();
            if (map != null) {
                for (Map.Entry<Integer, WriteFieldMapping.ValueAttribute> entry2 : map.entrySet()) {
                    String head = entry2.getValue().getHead();
                    Integer colIndex = entry2.getKey();
                    Cell cell = row.getCell(colIndex);
                    if (cell == null) {
                        cell = row.createCell(colIndex);
                    }
                    /*// use theme
                    if (!useTemplate && sheetProcessor.getTheme() != null) {
                        cell.setCellStyle(style);

                    }*/
                    cell.setCellValue(head);
                }
            }
        }
    }
}
