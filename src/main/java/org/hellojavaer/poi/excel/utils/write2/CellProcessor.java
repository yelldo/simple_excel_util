package org.hellojavaer.poi.excel.utils.write2;

import lombok.Data;
import org.apache.poi.ss.usermodel.Cell;

/**
 * Cell 处理器
 *
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

    @Override
    public void process(WriteProcessor processor) {
        writeContent();
    }

    private void writeContent(){

    }
}
