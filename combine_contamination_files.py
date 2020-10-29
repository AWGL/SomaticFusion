#!/usr/bin/env python
# coding: utf-8

"""
contamination_check.py


Author: Laura McCluskey
Created: 22nd Sept 2020
"""


import pandas 

#read in star-fusion and arriba contamination files and merge them into one dataframe
arriba = pandas.read_csv("contamination_arriba.csv", sep=",")
star_fusion = pandas.read_csv("contamination_star_fusion.csv", sep=",")
RMATS= pandas.read_csv("contamination_RMATS.csv", sep=",")
starFusion_arriba = pandas.merge(arriba, star_fusion, on=["Sample"])
overall = pandas.merge(starFusion_arriba, RMATS, on=["Sample"])

#function to combine the star-fusion and arriba contamination columns
def contamination(table):
    if ((table["Contamination_arriba"]=="Yes") or (table["Contamination_star_fusion"]=="Yes") or (table["Contamination_RMATS"]=="Yes")):
        return "Yes"
    else:
        return "No"



#combine the arriba and star-fusion and RMATS referral contamination files
def contamination_referral(table):
    if ((table["Contamination_referral_arriba"]=="Yes") or (table["Contamination_referral_star_fusion"]=="Yes") or (table["Contamination_referral_RMATS"]=="Yes")):
        return "Yes"
    else:
        return "No"



overall["Contamination"]= overall.apply(lambda overall : contamination(overall), axis=1)
overall["Contamination_referral"]= overall.apply(lambda overall : contamination_referral(overall), axis=1)
del [overall["Contamination_arriba"],overall["Contamination_referral_arriba"], overall["Contamination_star_fusion"], overall["Contamination_referral_star_fusion"], overall["Contamination_referral_RMATS"], overall["Contamination_RMATS"]]

overall.to_csv('contamination.csv', index=False)