package purepoi.temlate;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.format.CellDateFormatter;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Date;
import java.util.List;

/**
 * Created by luzy on 2017/10/26.
 */
public abstract class ExcelTemplate {

    public static final String DYNAMIC_ROW_COL_TEMPLATE = "excel.template/SettleDetailsSummary.xlsx";

    abstract Workbook build(String templatePath, String title, List<Object[]> data) throws IOException;

    abstract void write(OutputStream output, Workbook wb) throws IOException;

    protected void setCellValue(Cell cell, Object value){
        if (value == null) {
            cell.setCellValue((String) null);
            return;
        }
        if (value instanceof Short) {
            Short temp = (Short) value;
            cell.setCellValue((double) temp.shortValue());
        } else if (value instanceof Integer) {
            Integer temp = (Integer) value;
            cell.setCellValue((double) temp.intValue());
        } else if (value instanceof Long) {
            Long temp = (Long) value;
            cell.setCellValue((double) temp.longValue());
        } else if (value instanceof Float) {
            Float temp = (Float) value;
            cell.setCellValue((double) temp.floatValue());
        } else if (value instanceof Double) {
            Double temp = (Double) value;
            cell.setCellValue((double) temp.doubleValue());
        } else if (value instanceof Date) {
            Date dateVal = (Date) value;
            cell.setCellValue(dateVal);
            Workbook wb = cell.getRow().getSheet().getWorkbook();
            CellStyle cellStyle = cell.getCellStyle();
            DataFormat dataFormat = wb.getCreationHelper().createDataFormat();

            cellStyle.setDataFormat(dataFormat.getFormat("yyyy/MM/dd HH:mm:ss"));
            cell.setCellStyle(cellStyle);

            //日期格式转换
            String valStr;
            //if (DateUtil.isCellDateFormatted(cell)) {
            double valD = cell.getNumericCellValue();
            Date date = HSSFDateUtil.getJavaDate(valD);
            String dateFmt = null;

            if (cell.getCellStyle().getDataFormat() == 14) {
                dateFmt = "dd/mm/yyyy";
                valStr = new CellDateFormatter(dateFmt).format(date);
            } else {
                DataFormatter fmt = new DataFormatter();
                String valueAsInExcel = fmt.formatCellValue(cell);
                valStr = valueAsInExcel;
            }
            //}
            cell.setCellValue(valStr);
        } else if (value instanceof Boolean) {
            cell.setCellValue(((Boolean) value).booleanValue());
        } else { // String
            cell.setCellValue(value.toString());
        }
    }
}
