#!/usr/bin/env python
# coding: utf-8

"""
contamination_check.py


Author: Laura McCluskey
Created: 19th June 2020
Version: 0.0.4
"""


import pandas 

samples = pandas.read_csv("samples_list.txt", sep="\t")
sampleList=samples["sampleId"].tolist()
referralList=samples["referral"].tolist()



sample_number=0

contamination_dataframe=pandas.DataFrame(columns=["Sample", "Contamination", "Contamination_referral"])


previous_sample=""

for sample in sampleList:
    contamination="No"
    contamination_previous="No"
    contamination_next="No"
    contamination_referral="No"



    #compare the fusions within the sample to the fusions in the previous sample, label contamination as "yes" if any are the same

    
    report = pandas.read_csv(sample+"/"+sample+"_fusions_adapted.tsv", sep="\t")

    report["fusion"]=report["#gene1"].str.cat(report["gene2"], sep="-")

    fusion_list=report["fusion"].tolist()

    if (sample_number==0):
        
        previous_list=fusion_list
        contamination="NA"
        
    else:
        current_list=fusion_list
        for fusion_current in current_list:
            for fusion_previous in previous_list:
                if (fusion_current==fusion_previous):
                    contamination_previous="Yes"
                    contamination="Yes"
        previous_list=fusion_list





    #if contamination above is "yes", check to see if the fusion is in the referral type files 

    referral=referralList[sample_number]
    referral_file=pandas.read_csv(/data/diagnostics/pipelines/SomaticFusion/SomaticFusion-0.0.4/Referrals/referral+".txt", sep="\t")
    gene_list=list(referral_file['Genes'])



    if (sample_number==0):
        
        contamination_referral="NA"


    if (contamination=="Yes"):

        for gene in gene_list:

            report_referral = pandas.read_csv(sample+"/Results/arriba/"+sample+"_fusion_report_"+gene+"_arriba.txt", sep="\t")
            report_referral["fusion"]=report_referral["gene1"].str.cat(report_referral["gene2"], sep="-")

            fusion_list_referral=report_referral["fusion"].tolist()
        
            report_referral_previous = pandas.read_csv(previous_sample+"/Results/arriba/"+previous_sample+"_fusion_report_"+gene+"_arriba.txt", sep="\t")
   

            report_referral_previous["fusion"]=report_referral_previous["gene1"].str.cat(report_referral["gene2"], sep="-")
            fusion_list_referral_previous=report_referral_previous["fusion"].tolist()
            current_list_referral=fusion_list_referral
            for fusion_current_referral in current_list_referral:
                for fusion_previous_referral in fusion_list_referral_previous:
                    if (fusion_current_referral==fusion_previous_referral):
                        contamination_referral="Yes"
    




    #compare the fusions within the sample to the fusions in the next sample, label contamination as "yes" if any are the same

    fusion_list=report["fusion"].tolist()

 

    if (sample_number==(len(sampleList)-1)):
     
       print("ignore")
        
    else:
        sample_next_number=sample_number+1
        sample_next=sampleList[sample_next_number]
        report_next = pandas.read_csv(sample_next+"/"+sample_next+"_fusions_adapted.tsv", sep="\t")

        report_next["fusion"]=report_next["#gene1"].str.cat(report_next["gene2"], sep="-")

        fusion_list_next=report_next["fusion"].tolist()

        current_list=fusion_list
        for fusion_current in current_list:
            for fusion_next in fusion_list_next:
                if (fusion_current==fusion_next):
                    contamination_next="Yes"
                    contamination="Yes"
        previous_list=fusion_list






    #if contamination with the next sample is "yes", check to see if the fusion is in the referral type files 


    if (sample_number==0):
        
        print("ignore")
    if (contamination_next=="Yes"):
        for gene in gene_list:
            
            report_referral = pandas.read_csv(sample+"/Results/arriba/"+sample+"_fusion_report_"+gene+"_arriba.txt", sep="\t")
            report_referral["fusion"]=report_referral["gene1"].str.cat(report_referral["gene2"], sep="-")

            fusion_list_referral=report_referral["fusion"].tolist()
            sample_next_number=sample_number+1
            report_referral_next = pandas.read_csv(sample_next+"/Results/arriba/"+sample_next+"_fusion_report_"+gene+"_arriba.txt", sep="\t")
   

            report_referral_next["fusion"]=report_referral_next["gene1"].str.cat(report_referral_next["gene2"], sep="-")
            fusion_list_referral_next=report_referral_next["fusion"].tolist()
            current_list_referral=fusion_list_referral
            for fusion_current_referral in current_list_referral:
                for fusion_next_referral in fusion_list_referral_next:
                    if (fusion_current_referral==fusion_next_referral):
                        contamination_referral="Yes"
  


    sample_contamination=pandas.DataFrame([[sample, contamination, contamination_referral]], columns=["Sample", "Contamination", "Contamination_referral"])
    contamination_dataframe=contamination_dataframe.append(sample_contamination)
                    
    
    sample_number=sample_number+1
    previous_sample=sample


contamination_dataframe.to_csv('contamination.csv', index=False)







    
        
