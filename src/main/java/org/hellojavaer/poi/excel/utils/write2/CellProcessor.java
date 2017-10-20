package org.hellojavaer.poi.excel.utils.write2;

import lombok.Data;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.format.CellDateFormatter;
import org.apache.poi.ss.usermodel.*;

import java.util.Date;

/**
 * Cell 处理器
 * <p>
 * Created by luzy on 2017/10/17.
 */
@Data
public class CellProcessor extends WriteProcessor {

    private Cell cell;
    private Object cellValue;

    public CellProcessor setCell(Cell cell) {
        this.cell = cell;
        return this;
    }

    public CellProcessor setCellValue(Object value) {
        this.cellValue = value;
        return this;
    }

    @Override
    public void process(WriteContext context) {
        writeContent();
    }

    private void writeContent() {
        if (cellValue == null) {
            cell.setCellValue((String) null);
            return;
        }
        if (cellValue instanceof Short) {
            Short temp = (Short) cellValue;
            cell.setCellValue((double) temp.shortValue());
        } else if (cellValue instanceof Integer) {
            Integer temp = (Integer) cellValue;
            cell.setCellValue((double) temp.intValue());
        } else if (cellValue instanceof Long) {
            Long temp = (Long) cellValue;
            cell.setCellValue((double) temp.longValue());
        } else if (cellValue instanceof Float) {
            Float temp = (Float) cellValue;
            cell.setCellValue((double) temp.floatValue());
        } else if (cellValue instanceof Double) {
            Double temp = (Double) cellValue;
            cell.setCellValue((double) temp.doubleValue());
        } else if (cellValue instanceof Date) {
            Date dateVal = (Date) cellValue;
            cell.setCellValue(dateVal);
            Workbook wb = cell.getRow().getSheet().getWorkbook();
            CellStyle cellStyle = cell.getCellStyle();
            DataFormat dataFormat = wb.getCreationHelper().createDataFormat();

            cellStyle.setDataFormat(dataFormat.getFormat("yyyy/MM/dd HH:mm:ss"));
            cell.setCellStyle(cellStyle);

            //日期格式转换
            String value = "";
            //if (DateUtil.isCellDateFormatted(cell)) {
            double val = cell.getNumericCellValue();
            Date date = HSSFDateUtil.getJavaDate(val);
            String dateFmt = null;

            if (cell.getCellStyle().getDataFormat() == 14) {
                dateFmt = "dd/mm/yyyy";
                value = new CellDateFormatter(dateFmt).format(date);
            } else {
                DataFormatter fmt = new DataFormatter();
                String valueAsInExcel = fmt.formatCellValue(cell);
                value = valueAsInExcel;
            }
            //}
            cell.setCellValue(value);
        } else if (cellValue instanceof Boolean) {
            cell.setCellValue(((Boolean) cellValue).booleanValue());
        } else { // String
            cell.setCellValue(cellValue.toString());
        }
    }
}
