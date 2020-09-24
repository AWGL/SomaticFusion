#!/usr/bin/env python
# coding: utf-8

"""
contamination_check.py


Author: Laura McCluskey
Created: 22nd Sept 2020
Version: 0.0.6
"""


import pandas 

#read in star-fusion and arriba contamination files and merge them into one dataframe
arriba = pandas.read_csv("contamination_arriba.csv", sep=",")
star_fusion = pandas.read_csv("contamination_STAR_Fusion.csv", sep=",")
overall = pandas.merge(arriba, star_fusion, on=["Sample"])


#function to combine the star-fusion and arriba contamination columns
def contamination(table):
    if ((table["Contamination_arriba"]=="Yes") or (table["Contamination_STAR_Fusion"]=="Yes")):
        return "Yes"
    else:
        return "No"



#combine the arriba and star-fusion referral contamination files
def contamination_referral(table):
    if ((table["Contamination_referral_arriba"]=="Yes") or (table["Contamination_referral_STAR_Fusion"]=="Yes")):
        return "Yes"
    else:
        return "No"



overall["contamination"]= overall.apply(lambda overall : contamination(overall), axis=1)
overall["contamination_referral"]= overall.apply(lambda overall : contamination_referral(overall), axis=1)
del [overall["Contamination_arriba"],overall["Contamination_referral_arriba"], overall["Contamination_STAR_Fusion"], overall["Contamination_referral_STAR_Fusion"]]

overall.to_csv('contamination.csv', index=False)