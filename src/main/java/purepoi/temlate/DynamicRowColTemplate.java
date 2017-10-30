package purepoi.temlate;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.Assert;

import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by luzy on 2017/10/26.
 */
public class DynamicRowColTemplate extends ExcelTemplate{

    /* 模板的基本信息 */
    private int titleColIndex;// 大标题列索引
    private int titleRowIndex;// 大标题行索引
    private int idxIndex;// 序号列
    private int rowHeadIndex;// 行标题索引
    private int colHeadIndex;// 列标题索引
    private int[] rowHeadTemplateCellStyleIndex;
    private int[] colHeadTemplateCellStyleIndex;
    private int[] idxTemplateCellStyleIndex;
    private int[] middleTemplateCellStyleIndex;
    private int maxRowIndex; // 行索引最大值
    private int maxColIndex; // 列索引最大值

    public static void main(String[] args) {
        //System.out.println(DynamicRowCol.class.getClassLoader().getResourceAsStream("SettleDetailsSummary.xlsx"));
        System.out.println(DynamicRowColTemplate.class.getClassLoader().getResourceAsStream("excel.template/SettleDetailsSummary.xlsx"));
    }

    public void setTitleIndex(Integer titleRowIndex, Integer titleColIndex) {
        this.titleRowIndex = titleRowIndex;
        this.titleColIndex = titleColIndex;
    }

    public void setRowAndColHeadIndex(Integer rowHeadIndex, Integer colHeadIndex) {
        this.rowHeadIndex = rowHeadIndex;
        this.colHeadIndex = colHeadIndex;
        this.maxRowIndex = rowHeadIndex;
        this.maxColIndex = colHeadIndex;
    }

    public void setIdxIndex(Integer idxIndex) {
        this.idxIndex = idxIndex;
    }

    public void setRowHeadTemplateCellStyleIndex(Integer rowIdx, Integer colIdx) {
        this.rowHeadTemplateCellStyleIndex = new int[]{rowIdx, colIdx};
    }
    public void setColHeadTemplateCellStyleIndex(Integer rowIdx, Integer colIdx) {
        this.colHeadTemplateCellStyleIndex = new int[]{rowIdx, colIdx};
    }
    public void setIdxTemplateCellStyleIndex(Integer rowIdx, Integer colIdx) {
        this.idxTemplateCellStyleIndex = new int[]{rowIdx, colIdx};
    }
    public void setMiddleTemplateCellStyleIndex(Integer rowIdx, Integer colIdx) {
        this.middleTemplateCellStyleIndex = new int[]{rowIdx, colIdx};
    }

    //public void write(Workbook wb, String fileName)

    @Override
    public Workbook build(String templatePath, String title, List<Object[]> data) throws IOException {
        Assert.hasLength(templatePath);
        Assert.notEmpty(data);

        // 读取Excel模板
        //XSSFWorkbook wb = new XSSFWorkbook(this.getClass().getClassLoader().getResourceAsStream("excel/template/SettleDetailsSummary.xlsx"));
        XSSFWorkbook wb = new XSSFWorkbook(this.getClass().getClassLoader().getResourceAsStream(templatePath));
        Sheet sheet0 = wb.getSheetAt(0);

        CellStyle titleTemplateCellStyle = sheet0.getRow(titleRowIndex).getCell(titleColIndex).getCellStyle();
        if (titleTemplateCellStyle == null) {
            // exception: 提供的模板有误
        }
        CellStyle rowHeadTemplateCellStyle = sheet0.getRow(rowHeadTemplateCellStyleIndex[0]).getCell(rowHeadTemplateCellStyleIndex[1]).getCellStyle();
        if (rowHeadTemplateCellStyle == null) {
            // exception: 提供的模板有误
        }
        CellStyle colHeadTemplateCellStyle = sheet0.getRow(colHeadTemplateCellStyleIndex[0]).getCell(colHeadTemplateCellStyleIndex[1]).getCellStyle();
        if (colHeadTemplateCellStyle == null) {
            // exception: 提供的模板有误
        }
        CellStyle idxTemplateCellStyle = sheet0.getRow(idxTemplateCellStyleIndex[0]).getCell(idxTemplateCellStyleIndex[1]).getCellStyle();
        if (idxTemplateCellStyle == null) {
            // exception: 提供的模板有误
        }
        CellStyle middleTemplateCellStyle = sheet0.getRow(middleTemplateCellStyleIndex[0]).getCell(middleTemplateCellStyleIndex[1]).getCellStyle();
        if (middleTemplateCellStyle == null) {
            // exception: 提供的模板有误
        }
        int columnWidth = sheet0.getColumnWidth(rowHeadTemplateCellStyleIndex[1]);// 行高
        int lineHeight = sheet0.getRow(rowHeadTemplateCellStyleIndex[0]).getHeight();// 列宽

        /* 行标题，列标题，数据区的创建 */
        Map<String, Integer> rowHeadMapping = new HashMap<>();// distributor
        Map<String, Integer> colHeadMapping = new HashMap<>();// hospital
        Row hRow = sheet0.getRow(rowHeadIndex); // 标题行
        if (hRow == null) {
            // exception:提供的模板有误
        }
        boolean maxColIndexAdded = false;
        boolean maxRowIndexAdded = false;
        Cell c;
        for (Object[] o : data) {
            int curRowIndex;
            int curColIndex;
            if (!rowHeadMapping.containsKey(o[0])) {
                maxColIndexAdded = true;
                maxColIndex++;
                sheet0.setColumnWidth(maxColIndex, columnWidth);
                c = hRow.getCell(maxColIndex);
                if (c == null) {
                    c = hRow.createCell(maxColIndex);
                }
                c.setCellStyle(rowHeadTemplateCellStyle);
                //c.setCellValue((String) o[0]);
                setCellValue(c, o[0]);
                rowHeadMapping.put((String) o[0], maxColIndex);
                curColIndex = maxColIndex;
            } else {
                curColIndex = rowHeadMapping.get(o[0]);
            }

            if (!colHeadMapping.containsKey(o[1])) {
                maxRowIndexAdded = true;
                maxRowIndex++;
                Row colTmpRow = sheet0.getRow(maxRowIndex);
                if (colTmpRow == null) {
                    colTmpRow = sheet0.createRow(maxRowIndex);
                }
                colTmpRow.setHeight((short) lineHeight);
                c = colTmpRow.getCell(colHeadIndex);
                if (c == null) {
                    c = colTmpRow.createCell(colHeadIndex);
                }
                c.setCellStyle(colHeadTemplateCellStyle);
                //c.setCellValue((String) o[1]);
                setCellValue(c, o[1]);
                colHeadMapping.put((String) o[1], maxRowIndex);
                curRowIndex = maxRowIndex;
            } else {
                curRowIndex = colHeadMapping.get(o[1]);
            }

            Row r = sheet0.getRow(maxRowIndex);

            // 递增序号
            if (maxRowIndexAdded) {
                c = r.createCell(idxIndex);
                c.setCellStyle(idxTemplateCellStyle);
                c.setCellValue(maxRowIndex - rowHeadIndex);
            }

            // 创建单元格，包括暂时没值的
            if (maxColIndexAdded && maxRowIndexAdded) {
                for (int j = colHeadIndex + 1; j <= maxColIndex; j++) {
                    c = r.createCell(j);
                    c.setCellStyle(middleTemplateCellStyle);
                    c.setCellValue(0);
                }
                for (int i = rowHeadIndex + 1; i < maxRowIndex; i++) {
                    r = sheet0.getRow(i);
                    c = r.createCell(maxColIndex);
                    c.setCellStyle(middleTemplateCellStyle);
                    c.setCellValue(0);
                }
            } else if (maxColIndexAdded && !maxRowIndexAdded) {
                for (int i = rowHeadIndex + 1; i <= maxRowIndex; i++) {
                    r = sheet0.getRow(i);
                    c = r.createCell(maxColIndex);
                    c.setCellStyle(middleTemplateCellStyle);
                    c.setCellValue(0);
                }
            } else if (!maxColIndexAdded && maxRowIndexAdded) {
                for (int j = colHeadIndex + 1; j <= maxColIndex; j++) {
                    c = r.createCell(j);
                    c.setCellStyle(middleTemplateCellStyle);
                    c.setCellValue(0);
                }
            }
            maxColIndexAdded = false;
            maxRowIndexAdded = false;
            // 设置当前值
            r = sheet0.getRow(curRowIndex);
            c = r.getCell(curColIndex);
            c.setCellStyle(middleTemplateCellStyle);
            // 针对这种情况可以定义一个枚举类，来盛放不同数据类型的当前工具的定义，方便在某些逻辑里，判断处理
            //c.setCellValue(o[2] == null ? 0 : (Double) o[2]);
            setCellValue(c, o[2]);
        }

        /* sheet完善 */
        String formula = "SUM(%s:%s)";  // excel 求和表达式
        // 列合计
        c = sheet0.getRow(rowHeadIndex).createCell(maxColIndex + 1);
        c.setCellStyle(rowHeadTemplateCellStyle);
        c.setCellValue("合计");
        Row sRow = sheet0.createRow(maxRowIndex + 1);
        sRow.setHeight((short) lineHeight);
        for (int i = colHeadIndex + 1; i <= maxColIndex; i++) {
            c = sRow.createCell(i);
            c.setCellStyle(middleTemplateCellStyle);
            c.setCellFormula(String.format(formula, new CellReference(rowHeadIndex + 1, i).formatAsString(), new CellReference(maxRowIndex, i).formatAsString()));
        }
        // 行合计
        c = sheet0.getRow(maxRowIndex + 1).createCell(colHeadIndex);
        c.setCellStyle(colHeadTemplateCellStyle);
        sheet0.setColumnWidth(maxColIndex+1,columnWidth);
        c.setCellValue("合计");
        for (int i = rowHeadIndex + 1; i <= maxRowIndex + 1; i++) {
            c = sheet0.getRow(i).createCell(maxColIndex + 1);
            c.setCellStyle(middleTemplateCellStyle);
            c.setCellFormula(String.format(formula, new CellReference(i, colHeadIndex + 1).formatAsString(), new CellReference(i, maxColIndex).formatAsString()));
        }
        sheet0.setForceFormulaRecalculation(true);// workbook打开的时候计算设置該值的sheet的公式

        // 序号列最后一行
        sheet0.getRow(maxRowIndex + 1).createCell(idxIndex).setCellStyle(idxTemplateCellStyle);

        // 大标题行，autosize
        Row tRow = sheet0.getRow(titleRowIndex);
        Cell tCell = tRow.getCell(titleColIndex);
        tCell.setCellStyle(titleTemplateCellStyle);
        tCell.setCellValue(title);// 变更标题

        // 大标题行，合并单元格 params: first row  last row  first column  last column
        CellRangeAddress region = new CellRangeAddress(titleRowIndex,titleRowIndex,titleColIndex,maxColIndex+1);
        sheet0.addMergedRegion(region);

        return wb;
    }

    @Override
    void write(OutputStream output, Workbook wb) throws IOException {
        wb.write(output);
    }
}
