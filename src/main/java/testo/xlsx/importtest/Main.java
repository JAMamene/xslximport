package testo.xlsx.importtest;

import org.apache.commons.cli.*;

import java.util.Arrays;
import java.util.Collections;
import java.util.List;

public class Main {

    public static void main(String[] args) {
        Options options = new Options();

        Option method = new Option("m", "method", true, "specific method name");
        method.setRequired(false);
        options.addOption(method);

        Option input = new Option("i", "input", true, "xlsx input file");
        input.setRequired(true);
        options.addOption(input);

        Option bench = new Option("b", "bench", false, "benchmark mode");
        bench.setRequired(false);
        options.addOption(bench);

        CommandLineParser parser = new DefaultParser();
        HelpFormatter formatter = new HelpFormatter();
        CommandLine cmd;

        try {
            cmd = parser.parse(options, args);
            String inputFileName = cmd.getOptionValue("input");
            boolean bencho = cmd.hasOption("bench");
            List<ImportType> importTypes;
            if (cmd.hasOption("method")) {
                ImportType importType = ImportType.getEnumByString(cmd.getOptionValue("method").toUpperCase());
                if (importType == null) {
                    System.err.println("Method " + cmd.getOptionValue("method") + " does not exist");
                    System.exit(1);
                }
                importTypes = Collections.singletonList(importType);
            } else {
                importTypes = Arrays.asList(ImportType.values());
            }
            Benchmark benchmark = new Benchmark(inputFileName, importTypes, bencho);
            benchmark.bench();
        } catch (ParseException e) {
            System.out.println(e.getMessage());
            formatter.printHelp("XLSX import test", options);
            System.exit(1);
        }
    }
}
