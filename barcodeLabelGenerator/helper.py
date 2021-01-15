'''Barcode Label Generator

   helper.py

   Created by Vicky Xu on 2020-12-30
   Copyright © 2020-2021 Vicky Xu. All rights reserved.
'''
from barcode import Code39
from barcode.writer import ImageWriter
from PIL import Image, ImageDraw, ImageFont
import os

from ProductClass import Product
from mapping import PRODUCT_REFNUM, PRODUCT_PARTNAME, PRODUCT_QTY

def importData(sheet):
    # empty list
    products = []

    # read data
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if (row[PRODUCT_REFNUM] == None) or\
            (row[PRODUCT_QTY] == None) or\
            (row[PRODUCT_PARTNAME] == None):
            break
        product = Product(refnum=int(row[PRODUCT_REFNUM]),
                          partName=str(row[PRODUCT_PARTNAME]),
                          quantity=str(row[PRODUCT_QTY]),
                          info = str(row[PRODUCT_PARTNAME]) + 'QTY' +
                                 str(row[PRODUCT_QTY])
                          )
        products.append(product)

    return products


def generateBarcodes(products, parent_dir, directory1, directory2):
    for i in products:
        # create an object of Code39 class and  
        # pass the number with the ImageWriter() as the writer 
        my_code = Code39(i.info, writer=ImageWriter(), add_checksum=False)

        # Our barcode is ready. Let's save it. 
        image = my_code.save(parent_dir + "/" + directory1 + "/" + "code%s_" % (i.refnum) + i.info,
                             {"module_width":0.15,
                              "module_height":12,
                              "font_size": 10,
                              "text_distance": 1.5,
                              #"font_path": 'Arial.ttf',
                              "quiet_zone": 3})

        # create a white canvas
        canvasWidth, canvasHeight = 600, 400
        canvas = Image.new('RGB', (canvasWidth, canvasHeight), 'white')
        draw = ImageDraw.Draw(canvas)
        
        # set font
        font1 = ImageFont.truetype("arial.ttf", 28)
        font2 = ImageFont.truetype("arial.ttf", 15)
        
        # draw texts
        draw.text((45,45), "PART NUMBER", font=font1, fill='black')
        draw.text((445,45), "QTY", font=font1, fill='black')
        draw.text((45,100), i.partName, font=font1, fill='black')
        draw.text((445,100), i.quantity, font=font1, fill='black')
        
        # open barcode image
        im1 = Image.open(os.path.expanduser(parent_dir + "/" + directory1
                                            + "/" + "code%s_" % (i.refnum) + i.info + ".png"))
        width, height = im1.size

        pastePos = ((canvasWidth-width)//2, 165)
        
        if width > canvasWidth - 20:
            #resize to fit in centre
            newSize = (canvasWidth - 20, 203)
            im1 = im1.resize(newSize)
            pastePos = (10, 165)

        # copy the barcode image to canvas
        canvas.paste(im1, pastePos)

        # add refnum to left-corner
        draw.text((30,350), str(i.refnum), font=font2, fill='black')
        
        #save images
        canvas.save(parent_dir + "/" + directory2 + "/" + "im%s_" % (i.refnum) + i.info + ".pdf")

