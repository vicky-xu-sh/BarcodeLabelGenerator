'''Barcode Label Generator

   ProductClass.py

   Created by Vicky Xu on 2020-12-30
   Copyright © 2020-2021 Vicky Xu. All rights reserved.
'''

from dataclasses import dataclass

@dataclass
class Product:
    refnum: int
    partName: str
    info: str
    quantity: str
    
