package org.hellojavaer.poi.excel.utils.write2;

import lombok.Data;
import org.apache.poi.ss.usermodel.*;

import java.util.Calendar;
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
        System.out.println("CellProcessor,writeContent...");
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
        } else if (cellValue instanceof Date) {// Date
            Date dateVal = (Date) cellValue;
            long time = dateVal.getTime();
            // read is based on 1899/12/31 but DateUtil.getExcelDate is base on
            // 1900/01/01
            if (time >= TIME_1899_12_31_00_00_00_000 && time < TIME_1900_01_01_00_00_00_000) {
                Date incOneDay = new Date(time + 24 * 60 * 60 * 1000);
                double d = DateUtil.getExcelDate(incOneDay);
                cell.setCellValue(d - 1);
            } else {
                Workbook wb = cell.getRow().getSheet().getWorkbook();
                CellStyle cellStyle = cell.getCellStyle();
                if (cellStyle == null) {
                    cellStyle = wb.createCellStyle();
                }
                DataFormat dataFormat = wb.getCreationHelper().createDataFormat();
                // @see #BuiltinFormats
                // 0xe, "m/d/yy"
                // 0x14 "h:mm"
                // 0x16 "m/d/yy h:mm"
                // {@linke https://en.wikipedia.org/wiki/Year_10,000_problem}
                /** [1899/12/31 00:00:00:000~1900/01/01 00:00:000) */
                if (time >= TIME_1899_12_31_00_00_00_000 && time < TIME_1900_01_02_00_00_00_000) {
                    cellStyle.setDataFormat(dataFormat.getFormat("h:mm"));
                    // cellStyle.setDataFormat(dataFormat.getFormat("m/d/yy h:mm"));
                } else {
                    // if ( time % (24 * 60 * 60 * 1000) == 0) {//for time
                    // zone,we can't use this way.
                    Calendar calendar = Calendar.getInstance();
                    calendar.setTime(dateVal);
                    int hour = calendar.get(Calendar.HOUR_OF_DAY);
                    int minute = calendar.get(Calendar.MINUTE);
                    int second = calendar.get(Calendar.SECOND);
                    int millisecond = calendar.get(Calendar.MILLISECOND);
                    if (millisecond == 0 && second == 0 && minute == 0 && hour == 0) {
                        cellStyle.setDataFormat(dataFormat.getFormat("m/d/yy"));
                    } else {
                        cellStyle.setDataFormat(dataFormat.getFormat("m/d/yy h:mm"));
                    }
                }
                cell.setCellStyle(cellStyle);
                cell.setCellValue(dateVal);
            }
        } else if (cellValue instanceof Boolean) {
            cell.setCellValue(((Boolean) cellValue).booleanValue());
        } else { // String
            cell.setCellValue(cellValue.toString());
        }
    }
}
