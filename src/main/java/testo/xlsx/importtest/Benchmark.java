package testo.xlsx.importtest;

import org.apache.poi.openxml4j.util.ZipSecureFile;

import java.io.File;
import java.io.OutputStream;
import java.io.PrintStream;
import java.util.List;

public class Benchmark {

    private String fileName;
    private List<ImportType> importTypes;
    private boolean bench;

    public Benchmark(String fileName, List<ImportType> importTypes, boolean bench) {
        this.fileName = fileName;
        this.importTypes = importTypes;
        this.bench = bench;
    }

    public void bench() {
        PrintStream emptyStream = new PrintStream(new OutputStream() {
            public void write(int b) {
                //DO NOTHING
            }
        });
        ZipSecureFile.setMinInflateRatio(0.00001);
        ClassLoader classLoader = Main.class.getClassLoader();

        File file = new File(classLoader.getResource(fileName).getFile());
        for (ImportType importType : importTypes) {
            try {
                if (bench) {
                    System.out.println(importType.importFile(file, emptyStream).toBench());
                }
                else {
                    System.out.println(importType.importFile(file, System.out));
                }
                emptyStream.flush();
                System.gc();
            } catch (Exception | OutOfMemoryError e) {
                System.err.println(importType.name() + " FAILED, reason : " + e.getMessage());
                e.printStackTrace();
            }
        }
    }
}
