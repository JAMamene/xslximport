import subprocess
import matplotlib.pyplot as plt
import pandas as pd
import json
import sys

dataDict = {}


def runBench(method_name, file_name):
    res = subprocess.check_output(["mvn", "-q", "exec:java", "-Dexec.args=-i " + file_name + " -b -m " + method_name],
                                  shell=True,
                                  stderr=subprocess.PIPE).decode("utf-8")
    print(res, flush=True)
    if res != "":
        jsonRes = json.loads(res)
        dataDict.update({jsonRes["method"]: [jsonRes["time"], jsonRes["heap"]]})


if len(sys.argv) != 2:
    print("Failure, you must pass a file name in the resource folder to start benchmarking")
    sys.exit(1)
file_name = sys.argv[1]
p = subprocess.Popen(["mvn", "-q", "clean", "install"], shell=True)
p.communicate()
p.wait()
methods = ["FASTEXCEL", "EXCEL_STREAMING", "POI_EVENTDRIVEN", "CSV", "POI"]
for m in methods:
    runBench(m, file_name)

barWidth = 0.25

df = pd.DataFrame.from_dict(dataDict, orient="index")
df.columns = ["time", "heap"]

fig = plt.figure()

ax = fig.add_subplot(111)
ax2 = ax.twinx()

print(df)
df.time.plot(kind="bar", color='gray', width=barWidth, edgecolor='white', position=1, ax=ax,
             title="Execution time and memory usage if different xlsx import methods")
df.heap.plot(kind="bar", color='blue', width=barWidth, edgecolor='white', position=0, ax=ax2)

ax.set_ylabel('Time(ms)')
ax2.set_ylabel('Heap(Mb)')

ax.legend(loc=2)
ax2.legend(loc=1)

plt.show()
