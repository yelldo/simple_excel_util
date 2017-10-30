package purepoi.draw;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;

/**
 * draw a line in a cell or a series of cells
 *
 * Created by luzy on 2017/10/24.
 */
public class ExcelWriteTest {

    /*  用到的API：
        row.setHeightInPoints(height * 0.75f);         //设置行高
        sheet.setColumnWidth(colIndex, cellWidth);  //设置列宽

        合并单元格
        CellRangeAddress region = new CellRangeAddress(first row, last row, first column, last column);
        sheet0.addMergedRegion(region);

        CellStyle titleStyle = wb.createCellStyle();
        titleStyle.setAlignment(CellStyle.ALIGN_CENTER);//水平居中
        titleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);//垂直居中
    */

    public static void main(String[] args) throws IOException {
        URL url = ExcelWriteTest.class.getResource("/");
        final String outputFilePath = url.getPath() + "output_file999.xlsx";
        File outputFile = new File(outputFilePath);
        outputFile.createNewFile();
        FileOutputStream output = new FileOutputStream(outputFile);
        //List<Object[]> data = getDataList();
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet0 = wb.createSheet();
        Row tRow = sheet0.createRow(0);
        Cell titleCell = tRow.createCell(0);
        titleCell.setCellValue("标题");

        // 合并居中 first row  last row  first column  last column
        CellRangeAddress region = new CellRangeAddress(0,0,0,4);
        sheet0.addMergedRegion(region);

        tRow.setHeightInPoints(77*0.75f);//行高

        //标题样式
        CellStyle titleStyle = wb.createCellStyle();
        titleStyle.setAlignment(CellStyle.ALIGN_CENTER);//水平居中
        titleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);//垂直居中
        //标题字体
        Font tFont = wb.createFont();
        tFont.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
        tFont.setFontHeightInPoints((short)20);

        titleStyle.setFont(tFont);
        titleCell.setCellStyle(titleStyle);

        Row head = sheet0.createRow(1);

        /*
        需要乘以0.75f,原因为什么去google
        */
        head.setHeightInPoints(38.5f*0.75f);//行高
        /*
        一个单元格cell 默认 基础长宽 1023 255
        */
        sheet0.setColumnWidth(1, 1023*4);  //设置列宽
        head.createCell(0).setCellValue("序号");
        //head.createCell(1).setCellValue("斜线");//画线

        Cell cell11 = head.createCell(1);
        cell11.setCellValue("斜线");//画线
        CellStyle cellStyle11 = wb.createCellStyle();
        cellStyle11.setWrapText(true);//先设置为自动换行
        cell11.setCellStyle(cellStyle11);
        cell11.setCellValue(new XSSFRichTextString("        hello\r\nworld!"));


        //draw a line in a cell
        CreationHelper helper = wb.getCreationHelper();
        ClientAnchor clientAnchor = helper.createClientAnchor();
        clientAnchor.setAnchorType(ClientAnchor.MOVE_AND_RESIZE);
        clientAnchor.setCol1(1);
        clientAnchor.setRow1(1);
        clientAnchor.setCol2(1);
        clientAnchor.setRow2(1);
        clientAnchor.setDx1(0);
        clientAnchor.setDy1(0);
        clientAnchor.setDx2(1023* XSSFShape.EMU_PER_PIXEL);
        clientAnchor.setDy2(255* XSSFShape.EMU_PER_PIXEL);

        //XSSFClientAnchor clientAnchor = new XSSFClientAnchor(0,0,1023,249,1,1,2,1);
        XSSFDrawing drawing = (XSSFDrawing) sheet0.createDrawingPatriarch();
        XSSFSimpleShape simpleShape = drawing.createSimpleShape((XSSFClientAnchor) clientAnchor);
        simpleShape.setShapeType(ShapeTypes.LINE);
        //simpleShape.setLineWidth(1.5);
        simpleShape.setLineStyleColor(0,0,0);
        //simpleShape.setLineStyle(3);

        head.createCell(2).setCellValue("斜线w");
        head.createCell(3).setCellValue("斜线");
        head.createCell(4).setCellValue("斜线");

        wb.write(output);

    }

    /*public static List<Object[]> getDataList(){

        Object[] o1 = new Object[]{"a1","b1",10};
        Object[] o2 = new Object[]{"a2","b2",20};
        Object[] o3 = new Object[]{"a2","b1",30};
        Object[] o4 = new Object[]{"a3","b2",40};
        List<Object[]> arr = new ArrayList<>();
        arr.add(o1);
        arr.add(o2);
        arr.add(o3);
        arr.add(o4);
        return arr;
    }*/
}
