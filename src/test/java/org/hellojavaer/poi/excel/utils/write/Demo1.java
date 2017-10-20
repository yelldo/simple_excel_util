package org.hellojavaer.poi.excel.utils.write;


import org.hellojavaer.poi.excel.utils.TestBean;
import org.hellojavaer.poi.excel.utils.write2.ExcelWriteUtil;
import org.hellojavaer.poi.excel.utils.write2.SheetProcessor;
import org.hellojavaer.poi.excel.utils.write2.WriteFieldMapping;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.*;

/**
 * Created by tianhc on 2017/10/20.
 */
public class Demo1 {

    public static void main(String[] args) throws IOException {

        System.out.println("...start");

        //文件输出路径
        URL url = Demo1.class.getResource("/");
        final String outputFilePath = url.getPath() + "demo_output2.xlsx";
        File outputFile = new File(outputFilePath);
        FileOutputStream output = new FileOutputStream(outputFile);

        WriteFieldMapping fieldMapping = new WriteFieldMapping();
        int index = 0;
        fieldMapping.put(index++, "shortField").setHead("shortField");
        fieldMapping.put(index++, "intField").setHead("intField");
        fieldMapping.put(index++, "longField").setHead("longField");
        fieldMapping.put(index++, "floatField").setHead("floatField");
        fieldMapping.put(index++, "doubleField").setHead("doubleField");
        fieldMapping.put(index++, "boolField").setHead("boolField");
        fieldMapping.put(index++, "stringField").setHead("stringField");
        fieldMapping.put(index++, "dateField").setHead("dateField");

        SheetProcessor sheetProcessor = new SheetProcessor();
        sheetProcessor.setSheetIndex(0);// required. It can be replaced with 'setSheetName(sheetName)';
        sheetProcessor.setStartRowIndex(1);//
        sheetProcessor.setFieldMapping(fieldMapping);// required
        sheetProcessor.setHeadRowIndex(0);
        //sheetProcessor.setTheme(0);
        // sheetProcessor.setTemplateRowIndex(1);
        //sheetProcessor.setDataList(getDateList());
        sheetProcessor.setDataList(getDateList2());

        ExcelWriteUtil.write(output,sheetProcessor);

        System.out.println("end...");
    }

    private static List<TestBean> getDateList() {
        List<TestBean> list = new ArrayList<TestBean>();
        TestBean bean = new TestBean();
        bean.setShortField((short) 2);
        bean.setIntField(3);
        bean.setLongField(4L);
        bean.setFloatField(5.1f);
        bean.setDoubleField(6.23);
        bean.setBoolField(true);
        //bean.setEnumField1("enumField1");
        //bean.setEnumField2("enumField2");
        bean.setStringField("test");
        bean.setDateField(new Date());

        list.add(bean);
        list.add(bean);
        list.add(bean);
        return list;
    }

    private static List<Map<String, Object>> getDateList2() {
        List<Map<String, Object>> list = new ArrayList<>();
        Map<String, Object> data = new HashMap<>();
        data.put("shortField", (short) 2);
        data.put("intField", 3);
        data.put("longField", 4L);
        data.put("floatField", 5.1f);
        data.put("doubleField", 6.23d);
        data.put("boolField", true);
        data.put("stringField", "yelldo");
        data.put("dateField", new Date());

        list.add(data);
        list.add(data);

        return list;
    }
}
