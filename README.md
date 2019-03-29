
# XLSX Import test
This program is used to benchmark and test various xlsx import technique in order to see which is the most efficient to load and parse an excel file of varying size.
Files have to be placed in the resources folder for testing

## Java Program
### Installation
The installation requires to have a working maven installation

    mvn clean install
### Execution
Minimal execution

    mvn exec:java -Dexec.args="-i  <xlsx filename in resources folder>"

The available arguments are :

     -b,--bench                  benchmark output mode
     -i,--input <filename>       xlsx input file
     -m,--method <method_name>   specific import method to test

## Python Benchmark Script
A python script is available to more precisely measure the impact of each method as well as produce graph representation of the execution.
This script requires python3 as well as a few libraries
### Execution

    python benchmark.py <xlsx filename in resources folder>

This script will produce a visualisation like this one

![Example on a 15Mb Excel File](https://image.noelshack.com/fichiers/2019/13/5/1553874052-15000.png)