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
        self.reference_set_eur1=data.referenceset_eur1
        self.reference_set_eur2=data.referenceset_eur2
        self.reference_set_eur3=data.referenceset_eur3
        self.reference_set_eur4=data.referenceset_eur4
        self.reference_set_usd1=data.referenceset_usd1
        self.reference_set_usd2=data.referenceset_usd2
        self.reference_set_usd3=data.referenceset_usd3
        self.reference_set_usd4=data.referenceset_usd4

    # a getter function

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
    def referenceset_eur1(self):
        return self.__referenceset_eur1

    # a setter function
    @referenceset_eur1.setter
    def referenceset_eur1(self, new_val):
        self.__referenceset_eur1 = new_val
        
        
    @property
    def referenceset_eur2(self):
        return self.__referenceset_eur2

    # a setter function
    @referenceset_eur2.setter
    def referenceset_eur2(self, new_val):
        self.__referenceset_eur2 = new_val


    @property
    def referenceset_eur3(self):
        return self.__referenceset_eur3

    # a setter function
    @referenceset_eur3.setter
    def referenceset_eur3(self, new_val):
        self.__referenceset_eur3 = new_val
        
    @property
    def referenceset_eur4(self):
        return self.__referenceset_eur4

    # a setter function
    @referenceset_eur4.setter
    def referenceset_eur4(self, new_val):
        self.__referenceset_eur4 = new_val

    @property
    def referenceset_usd1(self):
        return self.__referenceset_usd1

    # a setter function
    @referenceset_usd1.setter
    def referenceset_usd1(self, new_val):
        self.__referenceset_usd1 = new_val

    @property
    def referenceset_usd2(self):
        return self.__referenceset_usd2

    # a setter function
    @referenceset_usd2.setter
    def referenceset_usd2(self, new_val):
        self.__referenceset_usd2 = new_val

    @property
    def referenceset_usd3(self):
        return self.__referenceset_usd3

    # a setter function
    @referenceset_usd3.setter
    def referenceset_usd3(self, new_val):
        self.__referenceset_usd3 = new_val

    @property
    def referenceset_usd4(self):
        return self.__referenceset_usd4

    # a setter function
    @referenceset_usd4.setter
    def referenceset_usd4(self, new_val):
        self.__referenceset_usd4 = new_val
