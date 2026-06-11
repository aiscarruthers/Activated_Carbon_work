import pylightxl
import xlsxwriter
import pandas as pd
import json
import datetime as dt
import os
import warnings
import sys

pack_class = os.path.dirname(os.path.realpath(__file__)) + "packing_class.json"

with open(pack_class) as f:
    pack_dict = json.load(f)

class PLN():
    """
    PLN is a pallet lot number this is the primary object to be pulled from the 
    Item ledger(F4111) table in JDE. It should correspond to a batch of material
    packed off from the plant. It may include multiple batches of product packed
    from the plant but should never be a partial emptyong of product from the
    silos. An instance of a silo being filled can only produce one PLN never more.
    however a PLN can be made up of multiple instances of the silo being filled
    and emptied.

    Attributes:
        PLN(int): Product/Pallet Lot Number, This is the numerical identifier for
        material logged in the JDE edwards system. 
        DOM (datetime): Date of Manufacture
        Grade (string): This is the grade of carbon being packed into the PLN
        Recipe (string): This is the plant recipe used to manufacture this 
        product
        Tonnage(float): This the total weight of the carbon packed into the PLN
        FPQ(float): This is the First pass quality of the PLN, how much of this 
        material can be sold
        Packtype(string): This is a string used in JDE to indicate the pack type
        pallets(int): this is the number of pallets of material generated as a 
        part of this PLN
        overs(float): This is the amount of material that did not make up a full
        pallet.
    """
    def __init__(self, pln, dom, tonnage, second_item_no):
        """ PLN constructor to initiallise PLN
        Args:
        PLN (int): this is a 7 digitnumber identifying the Product lot number
        DOM (Datetime): this is the date of manufacture of the lot
        tonnage(float): this is the tonnage of material in the lot
        2nd_item_no(string): this is the string identifying the product grade, 
        and packtype.
        """
        self.pln = pln
        self.dom = dom
        self.tonnage = tonnage
        strings = [i for i in second_item_no.split()]
        if len(strings[0]) > 10:
            strings.insert(1,strings[0][10:])
            strings[0] = strings[0][:10]
    
        if len(strings) == 3:
            self.grade = strings[0]
            self.packtype = strings[1]
        
        else:
            print(second_item_no)
            self.grade = input("What grade should this WO be")
            self.packtype = input("What packtype is this WO")
    




