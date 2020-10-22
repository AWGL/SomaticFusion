#!/usr/bin/env python
# coding: utf-8

"""
contamination_check_RMATS.py


Author: Laura McCluskey
Created: 19th June 2020
"""

import sys
import pandas 
from collections import OrderedDict

seqId=sys.argv[1]
version=sys.argv[2]

samples = pandas.read_csv("samples_correct_order.txt", sep="\t")
sampleList=samples["Sample_ID"].tolist()
len_sample_list=len(sampleList)
contamination_RMATS=["No"]*len_sample_list
contamination_RMATS_dict=OrderedDict(zip(sampleList, contamination_RMATS))
contamination_referral_RMATS=["No"]*len_sample_list
contamination_referral_RMATS_dict=OrderedDict(zip(sampleList, contamination_referral_RMATS))


sample_number=0


for sample in sampleList:
    NTC_in_sample= ("NTC" in sample)

    if (NTC_in_sample==False):
        #get referral from variables file for sample
        variables = pandas.read_csv(sample+"/"+ sample +".variables", sep="\t")
        referral_table=variables[variables['#Illumina Variables File'].str.match('referral')]
        referral_string=referral_table.iloc[0,0]

       #get list of genes in the referral for the sample
        referral_equals, referral= referral_string.split("=")
        referral_file=pandas.read_csv("/data/diagnostics/pipelines/SomaticFusion/SomaticFusion-"+version+"/Referrals/"+referral+".txt", sep="\t")
        referrral_file=referral_file[referral_file['Genes']==("MET_exon14_skipping", "EGFRv3")]
        gene_list=list(referral_file['Genes'])

        len_gene_list=len(gene_list)

        gene_value=0
        while (gene_value<len_gene_list):
            if (gene_list[gene_value]=="MET_exon14_skipping"):
                gene_list[gene_value]="MET"
            elif (gene_list[gene_value]=="EGFRv3"):
                gene_list[gene_value]=="EGFR"
            gene_value=gene_value+1
   

        #set default contamination to "No"
        contamination="No"
        contamination_next="No"
        contamination_previous="No"
        contamination_referral="No"
        contamination_referral_next="No"
        contamination_referral_previous="No"


        #get the sampleId of the sample before and after in the sample list
        if (sample_number!=0):
            sample_previous_number=sample_number-1
            sample_previous=sampleList[sample_previous_number]

        if (sample_number!=(len_sample_list-1)):
            sample_next_number=sample_number+1
            sample_next=sampleList[sample_next_number]






        #compare the fusions within the sample to the fusions in sample before and after for MET_exon_skipping and EGFRvIII

        report = pandas.read_csv(sample+"/"+seqId+"_"+sample+"_RMATS_Report.tsv", sep="\t")
        if (len(report)>0):
            fusion_list=report["geneSymbol"].tolist()


            if (sample_number!=0):
                        
                report_previous = pandas.read_csv(sample_previous+"/"+seqId+"_"+sample_previous+"_RMATS_Report.tsv", sep="\t")  
                if (len(report_previous)>0):
                    fusion_list_previous=report_previous["geneSymbol"].tolist()

                    #compare the fusion list for the sample with fusion list for the sample before. Label contamination as "yes" if any fusions match.
                    for fusion_current in fusion_list:
                        for fusion_previous in fusion_list_previous:
                            if (fusion_previous==fusion_previous):
                                contamination="Yes"


            #compare the fusion list for the sample with the fusion list for the sample after. Label contamination as "yes" if any fusions match.

            if (sample_number<len_sample_list):

                report_next = pandas.read_csv(sample_next+"/"+seqId+"_"+sample_next+"_RMATS_Report.tsv", sep="\t") 
                if (len(report_next)>0):      
                    fusion_list_next=report_next["geneSymbol"].tolist()


                    for fusion_current in fusion_list:
                        for fusion_next in fusion_list_next:
                            if (fusion_current==fusion_next):
                                contamination="Yes"




        #compare the fusions within the sample to the fusions in sample before and after for the REFERRAL
        if (len_gene_list!=0):
            for gene in gene_list:
                report_MET_EGFR = pandas.read_csv(sample+"/"+seqId+"_"+sample+"_RMATS_Report.tsv", sep="\t")
                if (len(report_MET_EGFR)>0):
                    report_referral=report_MET_EGFR[report_MET_EGFR["geneSymbol"]==gene]
                    if (len(report_referral)>0):

                        fusion_list_referral=report_referral["geneSymbol"].tolist()


                        if (sample_number!=0):
                        
                            report_previous_MET_EGFR = pandas.read_csv(sample_previous+"/"+seqId+"_"+sample_previous+"_RMATS_Report.tsv", sep="\t")  
                            if (len(report_previous_MET_EGFR)>0):               
                                report_referral_previous=report_previous_MET_EGFR[report_previous_MET_EGFR["geneSymbol"]==gene]
                                if (len(report_referral_previous)>0):
                                    fusion_list_referral_previous=report_referral_previous["geneSymbol"].tolist()

                                    #compare the fusion list for the sample with fusion list for the sample before. Label contamination as "yes" if any fusions match.
                                    for fusion_current_referral in fusion_list_referral:
                                        for fusion_previous_referral in fusion_list_referral_previous:
                                            if (fusion_previous_referral==fusion_previous_referral):
                                                contamination_referral="Yes"



                        #compare the fusion list for the sample with the fusion list for the sample after. Label contamination as "yes" if any fusions match.

                        if (sample_number<len_sample_list):

                            report_next_MET_EGFR = pandas.read_csv(sample_next+"/"+seqId+"_"+sample_next+"_RMATS_Report.tsv", sep="\t") 
                            if (len(report_next_MET_EGFR)>0):                
                                report_referral_next=report_next_MET_EGFR[report_next_MET_EGFR["geneSymbol"]==gene]
                                if (len(report_referral_next)>0):      
                                    fusion_list_referral_next=report_referral_next["geneSymbol"].tolist()


                                    for fusion_current_referral in fusion_list_referral:
                                        for fusion_next_referral in fusion_list_referral_next:
                                            if (fusion_current_referral==fusion_next_referral):
                                                contamination_referral="Yes"


        #update values in contamination dictionary


        if contamination_referral_RMATS_dict[sample]=="No":
            contamination_referral_RMATS_dict[sample] =contamination_referral
        if (sample_number<(len_sample_list-1)):
            if contamination_referral_RMATS_dict[sample_next]=="No":
                contamination_referral_RMATS_dict[sample_next]=contamination_referral_next
        if (sample_number!=0):
            if contamination_referral_RMATS_dict[sample_previous]=="No":
                contamination_referral_RMATS_dict[sample_previous]=contamination_referral_previous


        if contamination_RMATS_dict[sample]=="No":
            contamination_RMATS_dict[sample] =contamination

                    

    #calculate the level of contamination in the NTC. If the number of fusions for the panel genes is less 0, contamination is "No". Otherwise contamination is "Yes".
    elif (NTC_in_sample==True):
        ntc_RMATS_report=pandas.read_csv(sample+"/"+seqId+"_"+sample+"_RMATS_Report.tsv", sep="\t")
        len_ntc_arriba_report=len(ntc_RMATS_report)
        if (len(ntc_RMATS_report>0)):
            contamination_referral_RMATS_dict[sample]="Yes"

    #Increase the sample_number by 1   
    sample_number=sample_number+1


#Convert the contamination dictionaries to a dataframe and output to a csv file
contamination_dataframe_panel=pandas.DataFrame(list(contamination_RMATS_dict.items()), columns=["Sample", "Contamination_RMATS"])
contamination_dataframe_referral=pandas.DataFrame(list(contamination_referral_RMATS_dict.items()), columns=["Sample", "Contamination_referral_RMATS"])
contamination_dataframe=pandas.merge(contamination_dataframe_panel, contamination_dataframe_referral, on="Sample")
contamination_dataframe.to_csv('contamination_RMATS.csv', index=False)







    
        
