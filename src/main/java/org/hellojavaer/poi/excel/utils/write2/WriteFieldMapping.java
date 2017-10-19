package org.hellojavaer.poi.excel.utils.write2;

import lombok.Data;
import org.hellojavaer.poi.excel.utils.common.Assert;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * 配置Excel列与属性之间的映射关系
 *
 * Created by luzy on 2017/10/17.
 */
public class WriteFieldMapping {


    Map<String, Map<Integer, ValueAttribute>> fieldMapping = new LinkedHashMap<>();

    public ValueAttribute put(String fieldName){
        Assert.notNull(fieldName);
        Map<Integer, ValueAttribute> map = fieldMapping.get(fieldName);
        if (map == null) {
            map = new HashMap<>();
            fieldMapping.put(fieldName, map);
        }
        ValueAttribute attribute = new ValueAttribute();
        //map.put()
        return null;
    }

    /**
     * 映射关系的属性
     * 单个值，或多个值（下拉列表）
     */
    @Data
    public class ValueAttribute{

        private CellProcessor cellProcessor;
        private CellValueMapping valueMapping;
        private String head;

        public ValueAttribute setCellProcessor(CellProcessor cellProcessor) {
            this.cellProcessor = cellProcessor;;
            return this;
        }

        public ValueAttribute setValueMapping(CellValueMapping valueMapping) {
            this.valueMapping = valueMapping;
            return this;
        }

        public ValueAttribute setHead(String head) {
            this.head = head;
            return this;
        }
    }

    public Map<String, Map<Integer, ValueAttribute>> export(){
        return fieldMapping;
    }
}
