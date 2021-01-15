# Barcode Label Generator

Barcode Label Generator is a Python program for generating barcode labels in Code 39.
This program is made at the request of a company to facilitate their export and shipping process.
The barcode label contains the product's name (part number) and its quantity.  

The program will read information from an excel file that is saved in the `data` directory.
The output will be saved as one single PDF document that includes all the labels generated, 
in the `data` directory.

For Windows users, a packaged executable file `cli(Windows).exe` is included. It can be run without a Python interpreter or installing any modules. 

## Usage
To use this generator, first prepare an Excel file using the format of `template.xlsx`, and then save the file in the `data` directory. 

For Windows users, simply run the executable file `cli(Windows).exe`, and follow the instructions in the command-line interface. For others, you need to run the `BarcodeGenerator.py` which can be found in the `barcodeLabelGenerator` directory. Note: the Python modules listed in the `requirements.txt` need to be installed first. 

## Development
For future development, this program can be easily modified to generate other types of barcodes, and allow users to select from different formats for the output file. 

## License
[MIT](https://choosealicense.com/licenses/mit/)
