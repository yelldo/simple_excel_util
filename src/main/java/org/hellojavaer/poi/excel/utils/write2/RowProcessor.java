package org.hellojavaer.poi.excel.utils.write2;

import lombok.Data;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.beans.BeanUtils;

import java.beans.PropertyDescriptor;
import java.util.Map;

/**
 * Row 处理器
 *
 * Created by luzy on 2017/10/17.
 */
@Data
public class RowProcessor extends WriteProcessor {

    private SheetProcessor sheetProcessor;
    private CellProcessor cellProcessor;
    private Row row;
    private Object rowData;

    public RowProcessor setRow(Row row) {
        this.row = row;
        return this;
    }

    public RowProcessor setRowData(Object rowData) {
        this.rowData = rowData;
        return this;
    }

    @Override
    void process(WriteContext context) {
        writeContent(context);
    }

    public void writeContent(WriteContext context){
        WriteFieldMapping fieldMapping = sheetProcessor.getFieldMapping();
        for (Map.Entry<String, Map<Integer, WriteFieldMapping.ValueAttribute>> entry : fieldMapping.export().entrySet()) {
            String fieldName = entry.getKey();
            Map<Integer, WriteFieldMapping.ValueAttribute> map = entry.getValue();
            for (Map.Entry<Integer, WriteFieldMapping.ValueAttribute> attributeEntry : map.entrySet()) {
                Integer colIndex = attributeEntry.getKey();
                WriteFieldMapping.ValueAttribute attribute = attributeEntry.getValue();
                if(rowData == null){
                    // something to do
                }
                // rowData != null
                Object val = getFieldValue(rowData, fieldName, false);

                // proc cell
                Cell cell = row.getCell(colIndex);
                if (cell == null) {
                    cell = row.createCell(colIndex);
                }
                context.setCurColIndex(colIndex);
                context.setCurCell(cell);

                cellProcessor.setCell(cell);
                cellProcessor.setCellValue(val);
                cellProcessor.process(context);
            }
        }
    }

    private static Object getFieldValue(Object obj, String fieldName, boolean isTrimSpace) {
        Object val = null;
        if (obj instanceof Map) {
            val = ((Map) obj).get(fieldName);
        } else {// java bean
            val = getProperty(obj, fieldName);
        }
        // trim
        if (val != null && val instanceof String && isTrimSpace) {
            val = ((String) val).trim();
            if ("".equals(val)) {
                val = null;
            }
        }
        return val;
    }

    private static Object getProperty(Object obj, String fieldName) {
        PropertyDescriptor pd = getPropertyDescriptor(obj.getClass(), fieldName);
        if (pd == null || pd.getReadMethod() == null) {
            throw new IllegalStateException("In class" + obj.getClass() + ", no getter method found for field '"
                    + fieldName + "'");
        }
        try {
            return pd.getReadMethod().invoke(obj, (Object[]) null);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private static PropertyDescriptor getPropertyDescriptor(Class<?> clazz, String propertyName) {
        return BeanUtils.getPropertyDescriptor(clazz, propertyName);
    }
}
