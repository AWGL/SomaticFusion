#!/usr/bin/env python
# coding: utf-8

"""
contamination_check.py


Author: Laura McCluskey
Created: 19th June 2020
"""


import sys
import pandas 
from collections import OrderedDict

version=sys.argv[1]

samples = pandas.read_csv("samples_correct_order.txt", sep="\t")
sampleList=samples["Sample_ID"].tolist()
len_sample_list=len(sampleList)


#contamination dictionary
contamination_star_fusion=["No"]*len_sample_list
contamination_star_fusion_dict=OrderedDict(zip(sampleList, contamination_star_fusion))

#contamination referral dictionary
contamination_referral_star_fusion=["No"]*len_sample_list
contamination_referral_star_fusion_dict=OrderedDict(zip(sampleList, contamination_referral_star_fusion))

sample_number=0


for sample in sampleList:

    NTC_in_sample= ("NTC" in sample)

    if (NTC_in_sample==False):
        #get referral from variables file for sample
        variables = pandas.read_csv(sample+"/"+ sample +".variables", sep="\t")
        referral_table=variables[variables['#Illumina Variables File'].str.match('referral')]
        referral_string=referral_table.iloc[0,0]
        referral_equals, referral= referral_string.split("=")

        #get list of genes in the referral for the sample
        referral_file=pandas.read_csv("/data/diagnostics/pipelines/SomaticFusion/SomaticFusion-"+version+"/Referrals/"+referral+".txt", sep="\t")
        gene_list=list(referral_file['Genes'])
        len_gene_list=len(gene_list)

        #set default contamination to "No"
        contamination="No"
        contamination_next="No"
        contamination_previous="No"
        contamination_referral="No"
        contamination_referral_next="No"
        contamination_referral_previous="No"

        #get the sampleId of the sample before and after in the sampleList
        if (sample_number!=0):
            sample_previous_number=sample_number-1
            sample_previous=sampleList[sample_previous_number]
 
        if (sample_number!=(len_sample_list-1)):
            sample_next_number=sample_number+1
            sample_next=sampleList[sample_next_number]


        #compare the fusions within the sample to the fusions in sample before and after for the PANEL

        report = pandas.read_csv(sample+"/fusionReport/"+sample+"_fusionReport.txt", sep="\t")
        fusion_list=[]
        fusion_list_previous=[]
        fusion_list_next=[]
        if (len(report)>0):
            report=report[report["Fusion_Name"].str.contains("ALK|ROS1|RET|BRAF|NTRK1|NTRK2|NTRK3")]
            if (len(report)>0):
                fusion_list1=report["Fusion_Name"].tolist()

                #get a list of fusions with the genes the alternate way round
                report["Fusion_Name"]=report["Fusion_Name"].astype(str)
                report["gene1"]=report["Fusion_Name"].str.split('--', expand=True)[0]
                report["gene2"]=report["Fusion_Name"].str.split('--', expand=True)[1]
                report["Fusion_Name2"]=report["gene2"].str.cat(report["gene1"], sep="--")
                fusion_list2=report["Fusion_Name2"].tolist()

                #create list containing the fusions with genes in both orders
                fusion_list=fusion_list1+fusion_list2

                 #compare the list of fusions in the current sample with those in the sample before
                if (sample_number!=0):
                    report_previous = pandas.read_csv(sample_previous+"/fusionReport/"+sample_previous+"_fusionReport.txt", sep="\t")
                    if (len(report_previous)>0):
                        report_previous=report_previous[report_previous["Fusion_Name"].str.contains("ALK|ROS1|RET|BRAF|NTRK1|NTRK2|NTRK3")]
                        if (len(report_previous)>0):
                            fusion_list_previous1=report_previous["Fusion_Name"].tolist()

                            #get a list of fusions with the genes the alternate way round for the sample before
                            report_previous["Fusion_Name"]=report_previous["Fusion_Name"].astype(str)
                            report_previous["gene1"]=report_previous["Fusion_Name"].str.split('--', expand=True)[0]
                            report_previous["gene2"]=report_previous["Fusion_Name"].str.split('--', expand=True)[1]
                            report_previous["Fusion_Name2"]=report_previous["gene2"].str.cat(report_previous["gene1"], sep="--")
                            fusion_list_previous2=report_previous["Fusion_Name2"].tolist()

                            #create list containing the fusions with genes in both orders for sample before
                            fusion_list_previous=fusion_list_previous1+fusion_list_previous2
                      

                            #compare the fusion list with the fusion list for the sample before. Label contamination as "yes" if any fusions match.                            
                            for fusion_current in fusion_list:
                                for fusion_previous in fusion_list_previous:
                                    if (fusion_current==fusion_previous):
                                        contamination_previous="Yes"
                                        contamination="Yes"


                #compare the list of fusions in the current sample with those in the sample after
                if (sample_number<(len_sample_list-1)):
                    report_next = pandas.read_csv(sample_next+"/fusionReport/"+sample_next+"_fusionReport.txt", sep="\t")
                    if (len(report_next)>0):
                        report_next=report_next[report_next["Fusion_Name"].str.contains("ALK|ROS1|RET|BRAF|NTRK1|NTRK2|NTRK3")]
                        if (len(report_next)>0):
                            fusion_list_next1=report_next["Fusion_Name"].tolist()

                            #get a list of fusions with the genes the alternate way round for the sample after 
                            report_next["Fusion_Name"]=report_next["Fusion_Name"].astype(str)
                            report_next["gene1"]=report_next["Fusion_Name"].str.split('--', expand=True)[0]
                            report_next["gene2"]=report_next["Fusion_Name"].str.split('--', expand=True)[1]
                            report_next["Fusion_Name2"]=report_next["gene2"].str.cat(report_next["gene1"], sep="--")
                            fusion_list_next2=report_next["Fusion_Name2"].tolist()

                            #create list containing the fusions with genes in both orders for the sample after
                            fusion_list_next=fusion_list_next1+fusion_list_next2

                            #compare the fusion list for the sample with the fusion list for the sample after. Label contamination as "yes" if any fusions match.
                            for fusion_current in fusion_list:
                                for fusion_next in fusion_list_next:
                                    if (fusion_current==fusion_next):
                                        contamination_next="Yes"
                                        contamination="Yes"



        #compare the fusions within the sample to the fusions in sample before and after for the REFERRAL

        #loop through all the genes in the referral
        if (len_gene_list!=0):

            for gene in gene_list:
                if ((gene!="MET_exon14_skipping") and (gene!="EGFRv3")):

                    report_referral = pandas.read_csv(sample+"/Results/STAR_Fusion/"+sample+"_fusion_report_"+gene+".txt", sep="\t")
                    if (len(report_referral)>0):
                        fusion_list_referral1=report_referral["Fusion_Name"].tolist()

                        #get a list of fusions with the genes the alternate way round
                        report_referral["Fusion_Name"]=report_referral["Fusion_Name"].astype(str)
                        report_referral["gene1"]=report_referral["Fusion_Name"].str.split('--', expand=True)[0]
                        report_referral["gene2"]=report_referral["Fusion_Name"].str.split('--', expand=True)[1]
                        report_referral["Fusion_Name2"]=report_referral["gene2"].str.cat(report_referral["gene1"], sep="--")
                        fusion_list_referral2=report_referral["Fusion_Name2"].tolist()

                        #create list containing the fusions with genes in both orders for sample
                        fusion_list_referral=fusion_list_referral1+fusion_list_referral2

        
                        if (sample_number!=0): 
                            report_referral_previous = pandas.read_csv(sample_previous+"/Results/STAR_Fusion/"+sample_previous+"_fusion_report_"+gene+".txt", sep="\t")
                            if (len(report_referral_previous)>0):
                                fusion_list_referral_previous1=report_referral_previous["Fusion_Name"].tolist()


                                #get a list of fusions with the genes the alternate way round for the sample before
                                report_referral_previous["Fusion_Name"]=report_previous["Fusion_Name"].astype(str)
                                report_referral_previous["gene1"]=report_referral_previous["Fusion_Name"].str.split('--', expand=True)[0]
                                report_referral_previous["gene2"]=report_referral_previous["Fusion_Name"].str.split('--', expand=True)[1]
                                report_referral_previous["Fusion_Name2"]=report_referral_previous["gene2"].str.cat(report_referral_previous["gene1"], sep="--")
                                fusion_list_referral_previous2=report_referral_previous["Fusion_Name2"].tolist()

 
                                #create list containing the fusions with genes in both orders for sample before
                                fusion_list_referral_previous=fusion_list_referral_previous1+fusion_list_referral_previous2

        
                                #compare the fusion list for the sample with fusion list for the sample before. Label contamination as "yes" if any fusions match.
                                for fusion_current_referral in fusion_list_referral:
                                    for fusion_previous_referral in fusion_list_referral_previous:
                                        if (fusion_current_referral==fusion_previous_referral):
                                            contamination_referral="Yes"


                        if (sample_number<(len_sample_list-1)):              
                            report_referral_next = pandas.read_csv(sample_next+"/Results/STAR_Fusion/"+sample_next+"_fusion_report_"+gene+".txt", sep="\t")
                            if (len(report_referral_next)>0):
                                fusion_list_referral_next1=report_referral_next["Fusion_Name"].tolist()

                                #get a list of fusions with the genes the alternate way round for the sample after
                                report_referral_next["Fusion_Name"]=report_referral_next["Fusion_Name"].astype(str)
                                report_referral_next["gene1"]=report_referral_next["Fusion_Name"].str.split('--', expand=True)[0]
                                report_referral_next["gene2"]=report_referral_next["Fusion_Name"].str.split('--', expand=True)[1]
                                report_referral_next["Fusion_Name2"]=report_referral_next["gene2"].str.cat(report_referral_next["gene1"], sep="--")
                                fusion_list_referral_next2=report_referral_next["Fusion_Name2"].tolist()

                                #create list containing the fusions with genes in both orders for sample after
                                fusion_list_referral_next=fusion_list_referral_next1+fusion_list_referral_next2


                                #compare the fusion list for the sample with the fusion list for the sample after. Label contamination as "yes" if any fusions match.
                                for fusion_current_referral in fusion_list_referral_next:
                                    for fusion_next_referral in fusion_list_referral_next:
                                        if (fusion_next_referral==fusion_next_referral):
                                            contamination_referral="Yes"



        #Replace the contamination values in the contamination dictionaries

        if (contamination_star_fusion_dict[sample]=="No"):
            contamination_star_fusion_dict[sample]=contamination
        if (sample_number<(len_sample_list-1)):
            if (contamination_star_fusion_dict[sample_next]=="No"):
                contamination_star_fusion_dict[sample_next]=contamination_next
        if (sample_number!=0):
            if (contamination_star_fusion_dict[sample_previous]=="No"):
                contamination_star_fusion_dict[sample_previous]=contamination_previous

        if (contamination_referral_star_fusion_dict[sample]=="No"):
            contamination_referral_star_fusion_dict[sample]=contamination_referral


    #calculate the level of contamination in the NTC. If the number of fusions for the panel genes is less 0, contamination is "No". Otherwise contamination is "Yes".
    elif (NTC_in_sample==True):
        ntc_star_fusion_report=pandas.read_csv(sample+"/fusionReport/"+sample+"_fusionReport.txt")
        if (len(ntc_star_fusion_report)>0):
            ntc_star_fusion_report=ntc_star_fusion_report[ntc_star_fusion_report["Fusion_Name"].str.contains("ALK|ROS1|RET|BRAF|NTRK1|NTRK2|NTRK3")]
            len_ntc_star_fusion_report=len(ntc_star_fusion_report)
            if (len_ntc_star_fusion_report>0):
                contamination_star_fusion_dict[sample]="Yes"
                contamination_referral_star_fusion_dict[sample]="Yes"
 		
    #Increase the sample_number by 1       
    sample_number=sample_number+1	

#Convert the contamination dictionaries to a dataframe and output to a csv file
contamination_panel_dataframe=pandas.DataFrame(list(contamination_star_fusion_dict.items()), columns=["Sample", "Contamination_star_fusion"])
contamination_referral_dataframe=pandas.DataFrame(list(contamination_referral_star_fusion_dict.items()), columns=["Sample", "Contamination_referral_star_fusion"])
contamination_dataframe=pandas.merge(contamination_panel_dataframe, contamination_referral_dataframe, on="Sample")
contamination_dataframe.to_csv('contamination_star_fusion.csv', index=False)









    
        
