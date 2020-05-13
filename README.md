# fluidix_merge

## How to use the code 

```
usage: fluidix_merge.py [-h] -i INPUT_FILENAME -t TEMPLATE_FILENAME -p
                        {96,48,196} -op OUTPUT_PREFIX -od OUTPUT_DIRECTORY

This is a test of the command line argument parser in Python.

optional arguments:
  -h, --help            show this help message and exit
  -i INPUT_FILENAME, --input_filename INPUT_FILENAME
  -t TEMPLATE_FILENAME, --template_filename TEMPLATE_FILENAME
  -p {96,48,196}, --plate_type {96,48,196}
  -op OUTPUT_PREFIX, --output_prefix OUTPUT_PREFIX
  ```
  
  ## Before you start
  
  You should source the environment by doing 
  
  ```
  source venv/bin/activate
  ```
  
  Example :
  
  ```
  python fluidix_merge.py -i  /Users/raniba/Documents/Dev/fluidix_merge/data/196_input/SA00728662_Box.csv -t templates/template_196.xlsx -p 196 -op SA00728662_Box -od data/196_output
  ```
  
