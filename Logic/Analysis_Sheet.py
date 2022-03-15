import pandas as pd
import numpy as np
import regex as re
from Logic import DataTransforming

data=DataTransforming

class analysis_sheet:
    def __init__(self):
        self.__company_eur=data.df_client_eur_total_company
        self.__company_usd=data.df_client_usd_total_company
        self.__bu_eur=data.df_client_eur_by_BU
        self.__bu_usd=data.df_client_usd_by_BU
        self.__reference_set_eur1=data.referenceset_eur1
        self.__reference_set_eur2=data.referenceset_eur2
        self.__reference_set_eur3=data.referenceset_eur3
        self.__reference_set_eur4=data.referenceset_eur4
        self.__reference_set_usd1=data.referenceset_usd1
        self.__reference_set_usd2=data.referenceset_usd2
        self.__reference_set_usd3=data.referenceset_usd3
        self.__reference_set_usd4=data.referenceset_usd4
        self.__reference_rev_1=data.df_reference_set_revenue1
        self.__reference_sales_1=data.df_reference_set_sales1
        self.__bm_region=data.df_reference_set_region1
        self.__bm_channels=data.df_reference_set_channel1
        self.__company_name=data.company_naming
        self.__bu_naming=data.bu_naming
        self.__reference_set_naming=data.reference_set_naming

    @property
    def company_name(self):
        return self.__company_name

    # a setter function
    @company_name.setter
    def company_name(self, new_val):
        self.__company_name = new_val




    @property
    def bu_naming(self):
        return self.__bu_naming

    # a setter function
    @bu_naming.setter
    def bu_naming(self, new_val):
        self.__bu_naming = new_val


    @property
    def reference_set_naming(self):
        return self.__reference_set_naming

    # a setter function
    @reference_set_naming.setter
    def reference_set_naming(self, new_val):
        self.__reference_set_naming = new_val


    @property
    def reference_rev_1(self):
        return self.__reference_rev_1

    # a setter function
    @reference_rev_1.setter
    def reference_rev_1(self, new_val):
        self.__reference_rev_1 = new_val


    @property
    def reference_sales_1(self):
        return self.__reference_sales_1

    @reference_sales_1.setter
    def reference_sales_1(self, new_val):
        self.__reference_sales_1 = new_val


    @property
    def bm_region(self):
        return self.__bm_region

    @bm_region.setter
    def bm_region(self, new_val):
        self.__bm_region = new_val


    @property
    def bm_channels(self):
        return self.__bm_channels

    @bm_channels.setter
    def bm_channels(self, new_val):
        self.__bm_channels = new_val



    @property
    def company_eur(self):
        return self.__company_eur

    # a setter function
    @company_eur.setter
    def company_eur(self, new_val):
        self.__company_eur = new_val

    @property
    def company_usd(self):
        return self.__company_usd

    # a setter function
    @company_usd.setter
    def company_usd(self, new_val):
        self.__company_usd = new_val

    @property
    def bu_eur(self):
        return self.__bu_eur

    # a setter function
    @bu_eur.setter
    def bu_eur(self, new_val):
        self.__bu_eur = new_val

    @property
    def bu_usd(self):
        return self.__bu_usd

    # a setter function
    @bu_usd.setter
    def bu_usd(self, new_val):
        self.__bu_usd = new_val

    @property
    def reference_set_eur1(self):
        return self.__reference_set_eur1

    # a setter function
    @reference_set_eur1.setter
    def reference_set_eur1(self, new_val):
        self.__reference_set_eur1 = new_val
        
        
    @property
    def reference_set_eur2(self):
        return self.__reference_set_eur2

    # a setter function
    @reference_set_eur2.setter
    def reference_set_eur2(self, new_val):
        self.__reference_set_eur2 = new_val


    @property
    def reference_set_eur3(self):
        return self.__reference_set_eur3

    # a setter function
    @reference_set_eur3.setter
    def reference_set_eur3(self, new_val):
        self.__reference_set_eur3 = new_val
        
    @property
    def reference_set_eur4(self):
        return self.__reference_set_eur4

    # a setter function
    @reference_set_eur4.setter
    def reference_set_eur4(self, new_val):
        self.__reference_set_eur4 = new_val

    @property
    def reference_set_usd1(self):
        return self.__reference_set_usd1

    # a setter function
    @reference_set_usd1.setter
    def reference_set_usd1(self, new_val):
        self.__reference_set_usd1 = new_val

    @property
    def reference_set_usd2(self):
        return self.__reference_set_usd2

    # a setter function
    @reference_set_usd2.setter
    def reference_set_usd2(self, new_val):
        self.__reference_set_usd2 = new_val

    @property
    def reference_set_usd3(self):
        return self.__reference_set_usd3

    # a setter function
    @reference_set_usd3.setter
    def reference_set_usd3(self, new_val):
        self.__reference_set_usd3 = new_val

    @property
    def reference_set_usd4(self):
        return self.__reference_set_usd4

    # a setter function
    @reference_set_usd4.setter
    def reference_set_usd4(self, new_val):
        self.__reference_set_usd4 = new_val
