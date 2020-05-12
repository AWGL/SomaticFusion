
#!/usr/bin/env python

"""
fusion_report_referrals.py

Author: Laura McCluskey
Created: 5th May 2020
Version: 0.0.2
"""

import numpy
import pandas
import sys

path=sys.argv[1]
sampleId=sys.argv[2]


referrals=['ALK', 'ROS1', 'RET', 'BRAF', 'EGFR', 'MET', 'NTRK1', 'NTRK2', 'NTRK3']


#fusion report from STAR-Fusion

fusion_report=pandas.read_csv(path+"fusionReport/"+sampleId+"_fusionReport.txt", sep="\t")
genes=fusion_report["Fusion_Name"].str.split("--", expand=True)
fusion_report["gene1"]=genes[0]
fusion_report["gene2"]=genes[1]
fusion_report=pandas.DataFrame(fusion_report)


#fusion reports from arriba

fusion_report_arriba=pandas.read_csv(path+sampleId+"_fusions_adapted.tsv", sep="\t")
fusion_report_arriba=fusion_report_arriba.rename(columns={'#gene1': 'gene1'})

fusion_report_arriba_discarded=pandas.read_csv(path+sampleId+"_fusions_discarded_adapted.tsv", sep="\t")
fusion_report_arriba_discarded=fusion_report_arriba_discarded.rename(columns={'#gene1': 'gene1'})


for referral in referrals:

    fusion_report_filtered= fusion_report[(fusion_report.gene1==referral) | (fusion_report.gene2==referral)]
    fusion_report_filtered=fusion_report_filtered.iloc[:,1:17]
    fusion_report_filtered.to_csv("./Results/STAR_Fusion/"+sampleId+"_fusion_report_"+referral+".txt", index="None", sep ='\t')
 
    fusion_report_filtered_arriba= fusion_report_arriba[(fusion_report_arriba.gene1==referral) | (fusion_report_arriba.gene2==referral)]
    fusion_report_filtered_arriba.to_csv("./Results/arriba/"+sampleId+"_fusion_report_"+referral+"_arriba.txt", index="None", sep ='\t')

    fusion_report_filtered_arriba_discarded= fusion_report_arriba_discarded[(fusion_report_arriba_discarded.gene1==referral) | (fusion_report_arriba_discarded.gene2==referral)]
    fusion_report_filtered_arriba_discarded.to_csv("./Results/arriba_discarded/"+sampleId+"_fusion_report_"+referral+"_arriba_discarded.txt", index="None", sep ='\t')



