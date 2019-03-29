package testo.xlsx.importtest;

public class ExecutionSummary {

    private String method;
    private long execTime;
    private long heapSize;
    private long heapFreeSize;

    public ExecutionSummary(String method, long execTime, long heapSize, long heapFreeSize) {
        this.method = method;
        this.execTime = execTime;
        this.heapSize = heapSize;
        this.heapFreeSize = heapFreeSize;
    }

    @Override
    public String toString() {
        return this.method + " : \n\tExecution Time : " + execTime + " MilliSeconds\n" +
                "\tHeap Size : " + formatSize(heapSize) + "\n" +
                "\tFree Heap Size : " + formatSize(heapFreeSize) + "\n";
    }

    public String toBench() {
        return "{\"method\":\"" + method + "\",\"time\":" + execTime + ",\"heap\":" + heapSize/1000000 + ",\"freeheap\":" + heapFreeSize/1000000 + "}";
    }

    private static String formatSize(long v) {
        if (v < 1024) return v + " B";
        int z = (63 - Long.numberOfLeadingZeros(v)) / 10;
        return String.format("%.1f %sB", (double) v / (1L << (z * 10)), " KMGTPE".charAt(z));
    }
}
