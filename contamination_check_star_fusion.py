#!/usr/bin/env python
# coding: utf-8

"""
contamination_check.py


Author: Laura McCluskey
Created: 19th June 2020
Version: 0.0.6
"""


import pandas 

samples = pandas.read_csv("samples_list.txt", sep="\t")
sampleList=samples["sampleId"].tolist()




sample_number=0

contamination_dataframe=pandas.DataFrame(columns=["Sample", "Contamination_STAR_Fusion", "Contamination_referral_STAR_Fusion"])


previous_sample=""

for sample in sampleList:
    NTC_in_sample= ("NTC" in sample)

    if (NTC_in_sample==False):

        variables = pandas.read_csv(sample+"/"+ sample +".variables", sep="\t")
        referral_table=variables[variables['#Illumina Variables File'].str.match('referral')]
        referral_string=referral_table.iloc[0,0]
        referral_equals, referral= referral_string.split("=")

        contamination="No"
        contamination_previous="No"
        contamination_next="No"
        contamination_referral="No"



        #compare the fusions within the sample to the fusions in the previous sample, label contamination as "yes" if any are the same

    
        report = pandas.read_csv(sample+"/fusionReport/"+sample+"_fusionReport.txt", sep="\t")


        fusion_list=report["Fusion_Name"].tolist()
 
        if (sample_number==0):
        
            previous_list=fusion_list
            contamination="No"
        
        else:
            current_list=fusion_list
            for fusion_current in current_list:
                for fusion_previous in previous_list:
                    if (fusion_current==fusion_previous):
                        contamination_previous="Yes"
                        contamination="Yes"
            previous_list=fusion_list



        #if contamination above is "yes", check to see if the fusion is in the referral type files 

        referral_file=pandas.read_csv("/data/diagnostics/pipelines/SomaticFusion/SomaticFusion-0.0.4/Referrals/"+referral+".txt", sep="\t")
        gene_list=list(referral_file['Genes'])



        if (sample_number==0):
        
            contamination_referral="No"


        if (contamination=="Yes"):

            for gene in gene_list:

                report_referral = pandas.read_csv(sample+"/Results/STAR_Fusion/"+sample+"_fusion_report_"+gene+".txt", sep="\t")

                if (len(report_referral)!=0):
                    print(len(report_referral))

                    fusion_list_referral=report_referral["Fusion_Name"].tolist()
        
                    report_referral_previous = pandas.read_csv(previous_sample+"/Results/STAR_Fusion/"+previous_sample+"_fusion_report_"+gene+".txt", sep="\t")
   

                    fusion_list_referral_previous=report_referral_previous["Fusion_Name"].tolist()
                    current_list_referral=fusion_list_referral
                    for fusion_current_referral in current_list_referral:
                        for fusion_previous_referral in fusion_list_referral_previous:
                            if (fusion_current_referral==fusion_previous_referral):
                                contamination_referral="Yes"
    




        #compare the fusions within the sample to the fusions in the next sample, label contamination as "yes" if any are the same

        fusion_list=report["Fusion_Name"].tolist()

 

        if (sample_number==(len(sampleList)-1)):
     
            print("ignore")
        
        else:
            sample_next_number=sample_number+1
            sample_next=sampleList[sample_next_number]
            report_next = pandas.read_csv(sample_next+"/fusionReport/"+sample_next+"_fusionReport.txt", sep="\t")


            fusion_list_next=report_next["Fusion_Name"].tolist()

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
            
                report_referral = pandas.read_csv(sample+"/Results/STAR_Fusion/"+sample+"_fusion_report_"+gene+".txt", sep="\t")
                
                if (len(report_referral)!=0):
                    fusion_list_referral=report_referral["Fusion_Name"].tolist()
                    sample_next_number=sample_number+1
                    report_referral_next = pandas.read_csv(sample_next+"/Results/STAR_Fusion/"+sample_next+"_fusion_report_"+gene+".txt", sep="\t")
   
                    if (len(report_referral_next)!=1):
 
                        fusion_list_referral_next=report_referral_next["Fusion_Name"].tolist()
                        current_list_referral=fusion_list_referral
                        for fusion_current_referral in current_list_referral:
                            for fusion_next_referral in fusion_list_referral_next:
                                if (fusion_current_referral==fusion_next_referral):
                                    contamination_referral="Yes"
  


        sample_contamination=pandas.DataFrame([[sample, contamination, contamination_referral]], columns=["Sample", "Contamination_STAR_Fusion", "Contamination_referral_STAR_Fusion"])
        contamination_dataframe=contamination_dataframe.append(sample_contamination)
                    
    
        sample_number=sample_number+1
        previous_sample=sample



    elif (NTC_in_sample==True):
	ntc_star_fusion_report=pandas.read_csv(sample+"/fusionReport/"+sample+"_fusionReport.txt", sep="\t")
	len_ntc_star_fusion_report=len(ntc_star_fusion_report)
	contamination_NTC="No"
	if (len_ntc_star_fusion_report>0):
		contamination_NTC="Yes"
 		
	contamination_dataframe = contamination_dataframe.append({"Sample": sample, "Contamination_STAR_Fusion": contamination_NTC, "Contamination_referral_STAR_Fusion":contamination_NTC}, ignore_index=True)

	

    contamination_dataframe.to_csv('contamination_STAR_Fusion.csv', index=False)







    
        
