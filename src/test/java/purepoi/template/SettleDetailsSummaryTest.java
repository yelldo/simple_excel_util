package purepoi.template;

import purepoi.temlate.DynamicRowColTemplate;
import purepoi.temlate.ExcelTemplate;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by luzy on 2017/10/25.
 */
public class SettleDetailsSummaryTest<T> {

    public static void main(String[] args) throws IOException {
        DynamicRowColTemplate template = new DynamicRowColTemplate();
        template.setTitleIndex(0, 0);
        template.setIdxIndex(0);
        template.setRowAndColHeadIndex(2, 1);
        template.setRowHeadTemplateCellStyleIndex(2, 2);
        template.setColHeadTemplateCellStyleIndex(3, 1);
        template.setIdxTemplateCellStyleIndex(3, 0);
        template.setMiddleTemplateCellStyleIndex(3, 2);
        Map<String, Object> map = new HashMap<>();
        map.put("title", "title");
        map.put("data", getDataList());

        URL url = SettleDetailsSummaryTest.class.getResource("/");
        final String outputFilePath = url.getPath() + "SettleDetailsSummary_filled_output.xlsx";
        File outputFile = new File(outputFilePath);
        outputFile.createNewFile();
        FileOutputStream output = new FileOutputStream(outputFile);

        template.build(ExcelTemplate.DYNAMIC_ROW_COL_TEMPLATE,"title2",getDataList()).write(output);
    }

    public static List<Object[]> getDataList() {
        Object[] o1 = new Object[]{"a1", "b1", 10d};// distributor,hospital,amount
        Object[] o2 = new Object[]{"a2", "b2", 20d};
        Object[] o3 = new Object[]{"a2", "b1", 30d};
        Object[] o4 = new Object[]{"a3", "b2", 40d};
        Object[] o5 = new Object[]{"a3", "b3", 50d};
        Object[] o6 = new Object[]{"a4", "b4", 60d};
        Object[] o7 = new Object[]{"a5", "b5", 70d};
        List<Object[]> arr = new ArrayList<>();
        arr.add(o1);
        arr.add(o2);
        arr.add(o3);
        arr.add(o4);
        arr.add(o5);
        arr.add(o6);
        arr.add(o7);
        return arr;
    }
}
