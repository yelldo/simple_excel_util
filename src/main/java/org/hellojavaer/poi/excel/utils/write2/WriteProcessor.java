package org.hellojavaer.poi.excel.utils.write2;

import lombok.Data;

/**
 * Created by luzy on 2017/10/17.
 */
@Data
abstract class WriteProcessor {

    protected WriteProcessor writeProcessor;
    protected WriteContext writeContext;

    public WriteProcessor setWriteContext(WriteContext context) {
        this.writeContext = context;
        return this;
    }

    public WriteProcessor setWriteProcessor(WriteProcessor processor) {
        this.writeProcessor = processor;
        return this;
    }

    protected void beforeProcess() {
    }

    abstract void process(WriteProcessor processor);

    protected void onException() {
    }

    protected void afterProcess() {
    }

}
