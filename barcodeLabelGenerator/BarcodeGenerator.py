'''Barcode Label Generator

   BarcodeGenerator.py

   Created by Vicky Xu on 2020-12-30
   Copyright © 2020-2021 Vicky Xu. All rights reserved.
'''

from barcode import Code39
from barcode.writer import ImageWriter
from PIL import Image, ImageDraw, ImageFont
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from datetime import datetime
from PyPDF2 import PdfFileMerger, PdfFileReader

from ProductClass import Product
from helper import importData, generateBarcodes

def main():
    done = False
    # main program loop
    while (not done):
        '''USER INPUT'''
        while True:
            try:
                filename = input("Name of the file? Please also include the extension (ie: .xlsx)\n"\
                               "Type \"q\" or \"Q\" to quit. \n")
                path = "../data/" + filename

                if (filename == "q") or (filename == "Q"):
                    print("Thank you for using! Quit successfully.")
                    done = True
                    break
                # open spreadsheet, read-only
                workbook = load_workbook(filename=path, read_only=True)
                sheet = workbook.active
                print("Successfully opened the Excel file!")
                break
            
            except InvalidFileException:
                print("\nInvalid File. \nCheck if the file exists in the current directory, or\n"\
                      "check if you can open it with Excel first. \n"\
                      "Supported formats are: .xlsx,.xlsm,.xltx,.xltm\n")
            except FileNotFoundError:
                print("\nFile not found. \nCheck if the file exists in the current directory, or\n"\
                      "check if you can open it with Excel first. \n"\
                      "Supported formats are: .xlsx,.xlsm,.xltx,.xltm\n")

        if not done:
            '''CREATE NEW DIRECTORIES'''
            # Leaf directories 
            directory1 = "new_codes"
            directory2 = "Barcode-info-images"
                
            # Parent Directory  
            parent_dir = "../data/" + filename.split('.')[0] + datetime.now().strftime("%Y.%m.%d-%H.%M.%S")
                
            # Path  
            path1 = os.path.join(parent_dir, directory1)  
            path2 = os.path.join(parent_dir, directory2)

            # Create the directory    
            os.makedirs(path1)
            os.mkdir(path2)


            '''IMPORT DATA FROM EXCEL SHEET'''
            products = importData(sheet)

            '''GENERATE BARCODE IMAGES'''
            generateBarcodes(products, parent_dir, directory1, directory2)
            

            '''MERGER ALL PDFS IN THE Barcode-info-images FOLDER'''
            loc = path2

            merger = PdfFileMerger()

            for i in products:
                pdf = path2 + "/" + "im%s_" % (i.refnum) + i.info + ".pdf"
                merger.append(PdfFileReader(pdf), 'rb')

            merger.write("../data/" + filename.split('.')[0] + "_" + datetime.now().strftime("%Y.%m.%d_%H.%M.%S")
                      + ".pdf")
            
            '''DELETE FILES IN new_code AND Barcode-info-images FOLDERS'''
            for i in products:
                # delete the new_code png file
                os.remove(os.path.expanduser(parent_dir + "/" + directory1 + "/" + "code%s_" % (i.refnum)
                                             + i.info + ".png"))
                # delete the pdf file
                os.remove(os.path.expanduser(parent_dir + "/" + directory2 + "/" +
                                            "im%s_" % (i.refnum) + i.info + ".pdf"))

            
            '''FINISHING'''
            # delete folders
            os.rmdir(path1)
            os.rmdir(path2)
            os.rmdir(parent_dir)
            
            print("Barcode labels have been successfully generated!\n")
            print("Do you have another Excel file to generate barcodes? (y/n)")
            while True:
                choice = input()
                if (choice == "y") or (choice == "Y"):
                    done = False
                    break
                elif (choice == "n") or (choice == "N"):
                    print("Thank you for using! Quit successfully.")
                    done = True
                    break
                else:
                    print("Invalid input. Please type \"y\" or \"Y\" or \"n\" or \"N\".")


if __name__ == "__main__":
    main()
    
