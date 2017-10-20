package org.hellojavaer.poi.excel.utils.write2;

import lombok.Data;

import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;

/**
 * Created by luzy on 2017/10/17.
 */
@Data
abstract class WriteProcessor {

    protected static long TIME_1899_12_31_00_00_00_000;
    protected static long TIME_1900_01_01_00_00_00_000;
    protected static long TIME_1900_01_02_00_00_00_000;

    static {
        DateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss:SSS");
        try {
            TIME_1899_12_31_00_00_00_000 = df.parse("1899-12-31 00:00:00:000").getTime();
            TIME_1900_01_01_00_00_00_000 = df.parse("1900-01-01 00:00:00:000").getTime();
            TIME_1900_01_02_00_00_00_000 = df.parse("1900-01-02 00:00:00:000").getTime();
        } catch (ParseException e) {
            throw new RuntimeException(e);
        }
    }

    protected void beforeProcess() {
    }

    abstract void process(WriteContext context);

    protected void onException() {
    }

    protected void afterProcess() {
    }

}
