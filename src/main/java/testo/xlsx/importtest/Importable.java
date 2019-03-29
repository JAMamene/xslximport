package testo.xlsx.importtest;

import java.io.File;
import java.io.IOException;
import java.io.PrintStream;

public interface Importable {

    default ExecutionSummary importFile(File file, PrintStream out) throws IOException {
        File temp = this.beforeExecute(file);
        if (temp != null) {
            file = temp;
        }
        long startTime = System.currentTimeMillis();
        this.execute(file, out);
        long endTime = System.currentTimeMillis();
        return new ExecutionSummary(
                this instanceof ImportType ? ((ImportType) this).name() : this.getClass().getSimpleName(),
                endTime - startTime,
                Runtime.getRuntime().totalMemory(),
                Runtime.getRuntime().freeMemory());
    }

    default File beforeExecute(File file) {
        return null;
    }

    void execute(File file, PrintStream out) throws IOException;

}
