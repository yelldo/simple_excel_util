package org.hellojavaer.poi.excel.utils.write2;

import lombok.Data;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Created by luzy on 2017/10/17.
 */
@Data
public class WriteContext {

    private Sheet                   curSheet;
    private Integer                curSheetIndex;
    private String                 curSheetName;

    private Row                     curRow;
    private Integer                curRowIndex;

    private Cell                    curCell;
    private Integer                curColIndex;
    private String                 curColStrIndex;

}
