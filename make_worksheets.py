import pandas
from openpyxl import Workbook
import os
import numpy
from openpyxl.utils.dataframe import dataframe_to_rows
import sys
from openpyxl.styles.borders import Border, Side, BORDER_MEDIUM, BORDER_THIN, BORDER_THICK
from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

wb=Workbook()
ws7=wb.create_sheet("Patient_demographics")
ws6= wb.create_sheet("NTC fusion report")
ws2=wb.create_sheet("total_coverage")
ws1= wb.create_sheet("Gene_fusion_report")
ws5=wb.create_sheet("Summary")
ws3=wb.create_sheet("coverage_with_duplicates")
ws4=wb.create_sheet("coverage_without_duplicates")

ws1["A4"]="Lab number"
ws1["B4"]="Patient name"
ws1["C4"]="Reason for referral"
ws1["D4"]="NGS wks"
ws1["E4"]="NextSeq runId"
ws1["F4"]="Checker 1 initials and date"
ws1["G4"]="Checker 2 initials and date"


ws7['A1']='Date Received'
ws7['B1']='LABNO'
ws7['C1']='Patient name'
ws7['D1']='DOB'
ws7['E1']='Reason for referral'
ws7['F1']='Referring Clinician'
ws7['G1']='Indication'
ws7['H1']='Due Date'
ws7['I1']='% Tumour'
ws7['J1']='DV200'
ws7['K1']='Pre-processing (dil/zymo)'
ws7['L1']='Qubit RNA conc after pre-processing'
ws7['M1']='NGS Worksheet'
ws7['N1']='NextSeq run ID'
ws7['O1']='PostPCR1 Qubit'
ws7['P1']='Result'
ws7['Q1']='Date reported'
ws7['R1']='Comments'

ws7['A6']="Checker 1 initials and date"
ws7['A7']="Checker 2 initals and date"

ws5['H5'].number_format='mm-dd-yy'




def get_referral_list(referral):
	
	#get the genes in the referral file for the sample
	referral_file=pandas.read_csv("/data/diagnostics/pipelines/SomaticFusion/SomaticFusion-"+version+"/Referrals/"+referral+".txt", sep="\t")
	referral_list=list(referral_file['Genes'])
	return(referral_list)





def get_NTC_fusion_report(ntc):
	
	#get the NTC star fusion results for all the genes in the samples referral type
	ws6["A4"]="NTC STAR-Fusion results"
	for referral in referral_list:
                
		if ((referral!= "MET_exon14_skipping") and (referral!="EGFRv3")):
			if (os.stat("../"+ntc+"/Results/STAR_Fusion/"+ntc+"_fusion_report_"+referral+".txt").st_size!=0):
      	
				NTC_star_fusion_report=pandas.read_csv("../"+ntc+"/Results/STAR_Fusion/"+ntc+"_fusion_report_"+referral+".txt", sep="\t")
	
				NTC_star_fusion_report=pandas.DataFrame(NTC_star_fusion_report)
	
				for row in dataframe_to_rows(NTC_star_fusion_report, header=True, index=False):
					ws6.append(row)
			else:
				print("NTC STAR-Fusion results cannot be found")


	#get the NTC arriba results for all the genes in the samples referral type
	ws6["A19"]="Arriba results"
	for referral in referral_list:

		if ((referral!= "MET_exon14_skipping") and (referral!="EGFRv3")):
	
			if (os.stat("../"+ntc+"/Results/arriba/"+ntc+"_fusion_report_"+referral+"_arriba.txt").st_size!=0):
      	
				NTC_arriba_report=pandas.read_csv("../"+ntc+"/Results/arriba/"+ntc+"_fusion_report_"+referral+"_arriba.txt", sep="\t")
	
				for row in dataframe_to_rows(NTC_arriba_report, header=True, index=False):
					ws6.append(row)
			else:
				print("NTC arriba file does not exist")

	return(NTC_star_fusion_report, NTC_arriba_report)

	

def get_star_fusion_report(referral_list, ntc_star_fusion_report):
	ws1["A9"]="STAR-Fusion results"

	#create empty pandas dataframe and append the STAR-Fusion results for all the genes in the sample's referral type

	star_fusion_report_final= pandas.DataFrame(columns=["Fusion_Name", "Split_Read_Count", "Spanning_Read_Count", "Left_Breakpoint", "Right_Breakpoint", "SpliceType", "LargeAnchorSupport", "FFPM", "LeftBreakEntropy", "RightBreakEntropy","CDS_Left_ID", "CDS_Left_Range", "CDS_Right_ID", "CDS_Right_Range", "Prot_Fusion_Type","Num_WT_Fragments_Left", "Num_WT_Fragments_Right", "Fusion_Allelic_Fraction"]) 

	for referral in referral_list:

		if ((referral!= "MET_exon14_skipping") and (referral!="EGFRv3")):


			if (os.stat("./Results/STAR_Fusion/"+sampleId+"_fusion_report_"+referral+".txt").st_size!=0):
      
				star_fusion_report=pandas.read_csv("Results/STAR_Fusion/"+sampleId+"_fusion_report_"+referral+".txt", sep="\t")	
				if len(star_fusion_report!=0):
					star_fusion_report=pandas.DataFrame(star_fusion_report)
					del star_fusion_report['Unnamed: 0']
				else:
					#if there are no fusions for the gene in the report append "<gene name>- no fusions found"
					dict_star_fusion=({"Fusion_Name": referral+"-no fusions found", "Split_Read_Count": "", "Spanning_Read_Count": "", "Left_Breakpoint": "", "Right_Breakpoint": "", "SpliceType": "", "LargeAnchorSupport": "", "FFPM": "", "LeftBreakEntropy": "", "RightBreakEntropy": "","CDS_Left_ID": "", "CDS_Left_Range": "", "CDS_Right_ID": "", "CDS_Right_Range": "", "Prot_Fusion_Type": "","Num_WT_Fragments_Left": "", "Num_WT_Fragments_Right": "", "Fusion_Allelic_Fraction": ""}) 
					star_fusion_report= pandas.DataFrame(dict_star_fusion, columns=["Fusion_Name", "Split_Read_Count", "Spanning_Read_Count", "Left_Breakpoint", "Right_Breakpoint", "SpliceType", "LargeAnchorSupport", "FFPM", "LeftBreakEntropy", "RightBreakEntropy","CDS_Left_ID", "CDS_Left_Range", "CDS_Right_ID", "CDS_Right_Range", "Prot_Fusion_Type","Num_WT_Fragments_Left", "Num_WT_Fragments_Right", "Fusion_Allelic_Fraction"], index=[0]) 	

			star_fusion_report_final=star_fusion_report_final.append(star_fusion_report, sort=False)

		else:
			print("STAR-Fusion report does not exist")
	

	for row in dataframe_to_rows(star_fusion_report_final, header=True, index=False):
		ws1.append(row)	
		


def get_arriba_fusion_report(referral_list):

	ws1["A23"]="Arriba results"


	#create empty pandas dataframe and append the arriba results for all the genes in the sample's referral type

	arriba_report_final= pandas.DataFrame(columns=["gene1", "gene2", "strand1(gene/fusion)", "strand2(gene/fusion)", "breakpoint1", "breakpoint2", "site1", "site2", "type", "direction1", "direction2", "split_reads1", "split_reads2", "discordant_mates", "coverage1", "coverage2", "confidence", "closest_genomic_breakpoint1", "closest_genomic_breakpoint2", "filters", "fusion_transcript", "reading_frame", "peptide_sequence", "read_identifiers"])

	for referral in referral_list:

		if ((referral!= "MET_exon14_skipping") and (referral!="EGFRv3")):

			if (os.stat("./Results/arriba/"+sampleId+"_fusion_report_"+referral+"_arriba.txt").st_size!=0):     	
				arriba_report=pandas.read_csv("Results/arriba/"+sampleId+"_fusion_report_"+referral+"_arriba.txt", sep="\t")

				if len(arriba_report!=0):
					del arriba_report['Unnamed: 0']					

				else:
					#if there are no fusions for the gene in the report append "<gene name>- no fusions found"
					dict_arriba=({"gene1": referral + "-no fusions found", "gene2": "", "strand1(gene/fusion)": "", "strand2(gene/fusion)": "", "breakpoint1": "", "breakpoint2": "", "site1": "", "site2": "", "type": "", "direction1": "", "direction2": "", "split_reads1": "", "split_reads2": "", "discordant_mates": "", "coverage1": "", "coverage2": "", "confidence": "", "closest_genomic_breakpoint1": "", "closest_genomic_breakpoint2": "", "filters": "", "fusion_transcript": "", "reading_frame" : "", "peptide_sequence": "", "read_identifiers": ""}) 
					arriba_report= pandas.DataFrame(dict_arriba, columns=["gene1", "gene2", "strand1(gene/fusion)", "strand2(gene/fusion)", "breakpoint1", "breakpoint2", "site1", "site2", "type", "direction1", "direction2", "split_reads1", "split_reads2", "discordant_mates", "coverage1", "coverage2", "confidence", "closest_genomic_breakpoint1", "closest_genomic_breakpoint2", "filters", "fusion_transcript", "reading_frame", "peptide_sequence", "read_identifiers"], index=[0]) 	

				arriba_report_final=arriba_report_final.append(arriba_report, sort=False)

			else:
				print("arriba report does not exist")

	for row in dataframe_to_rows(arriba_report_final, header=True, index=False):
		ws1.append(row)
				


def get_ntc_total_coverage():

	#get the ntc total coverage with and without duplicates 

	if (os.stat("../"+ntc+"/"+ntc+"_coverage.totalCoverage").st_size!=0):
		ntc_total_coverage=pandas.read_csv("../"+ntc+"/"+ntc+"_coverage.totalCoverage", sep="\t")
		ntc_total_average_depth=ntc_total_coverage["AVG_DEPTH"]
	else:
		print("NTC total coverage file does not exist")

	if (os.stat("../"+ntc+"/"+ntc+"_rmdup_coverage.totalCoverage").st_size!=0):
		ntc_total_coverage_rmdup=pandas.read_csv("../"+ntc+"/"+ntc+"_rmdup_coverage.totalCoverage", sep="\t")
		ntc_total_average_depth_rmdup=ntc_total_coverage_rmdup["AVG_DEPTH"]
	else:
		print("NTC total coverage rmdup file does not exist")

	return(ntc_total_average_depth, ntc_total_average_depth_rmdup)


def get_total_coverage(ntc_total_average_depth, ntc_total_average_depth_rmdup):

	#get the total coverage of the genes of interest with and without duplicates 
	ws2["A2"]="with duplicates"
	if (os.stat(sampleId+"_coverage.totalCoverage").st_size!=0):
		total_coverage=pandas.read_csv(sampleId+"_coverage.totalCoverage", sep="\t")
		total_coverage["NTC_AVG_DEPTH"]=ntc_total_average_depth
		total_coverage["%NTC contamination"]=(total_coverage["NTC_AVG_DEPTH"]/total_coverage["AVG_DEPTH"])*100
		del total_coverage["PERC_COVERAGE@100"]
		for row in dataframe_to_rows(total_coverage, header=True, index=False):
			ws2.append(row)
	else:
		print(" Total coverage file does not exist")

	if (os.stat(sampleId+"_rmdup_coverage.totalCoverage").st_size!=0):
		ws2["A32"]="without duplicates"
		total_coverage_rmdup=pandas.read_csv(sampleId+"_rmdup_coverage.totalCoverage", sep="\t")
		total_coverage_rmdup["NTC_AVG_DEPTH"]=ntc_total_average_depth_rmdup
		total_coverage_rmdup["%NTC contamination"]=(total_coverage_rmdup["NTC_AVG_DEPTH"].div(total_coverage_rmdup["AVG_DEPTH"]).replace (numpy.inf,0))*100
		del total_coverage_rmdup["PERC_COVERAGE@100"]
		for row in dataframe_to_rows(total_coverage_rmdup, header=True, index=False):
			ws2.append(row)
	else:
		print(" Total coverage file rmdup does not exist")
	ws2.column_dimensions['A'].width=50


def ntc_get_coverage():

	#get the coverage of the regions of interest with and without duplicates in the ntc
	if (os.stat("../"+ntc+"/"+ntc+"_coverage.coverage").st_size!=0):
		ntc_coverage=pandas.read_csv("../"+ntc+"/"+ntc+"_coverage.coverage", sep="\t")
		ntc_average_depth=ntc_coverage["AVG_DEPTH"]
		del ntc_coverage["PERC_COVERAGE@100"]
	else:
		print(" NTC coverage file does not exist")

	if (os.stat("../"+ntc+"/"+ntc+"_rmdup_coverage.coverage").st_size!=0):
		ntc_coverage_rmdup=pandas.read_csv("../"+ntc+"/"+ntc+"_rmdup_coverage.coverage", sep="\t")
		ntc_average_depth_rmdup=ntc_coverage_rmdup["AVG_DEPTH"]
		del ntc_coverage_rmdup["PERC_COVERAGE@100"]
	else:
		print("NTC coverage rmdup file does not exist")

	return(ntc_average_depth, ntc_average_depth_rmdup)


def get_coverage(referral, ntc_average_depth, ntc_average_depth_rmdup):

	#get the coverage of the regions of interest with and without duplicates in the sample
	if (os.stat(sampleId+"_coverage.coverage").st_size!=0):
		coverage=pandas.read_csv(sampleId+"_coverage.coverage", sep="\t")
		coverage["NTC_AVG_DEPTH"]=ntc_average_depth
		coverage["%NTC contamination"]=(coverage["NTC_AVG_DEPTH"]/coverage["AVG_DEPTH"])*100
		del coverage["PERC_COVERAGE@100"]
		for row in dataframe_to_rows(coverage, header=True, index=False):
			ws3.append(row)

		#highlight the cells in the percentage NTC contamination column red if the value is greater than 10%
		colour= PatternFill("solid", start_color="00FF0000", end_color="00FF0000")
		ws3.conditional_formatting.add('G2:G300', CellIsRule( operator='greaterThan', formula=['10'], stopIfTrue=True, fill=colour))
	else:
		print("Coverage file does not exist")


	if (os.stat(sampleId+"_rmdup_coverage.coverage").st_size!=0):
		coverage_rmdup=pandas.read_csv(sampleId+"_rmdup_coverage.coverage", sep="\t")
		coverage_rmdup["NTC_AVG_DEPTH"]=ntc_average_depth_rmdup
		coverage_rmdup["%NTC contamination"]=(coverage_rmdup["NTC_AVG_DEPTH"]/coverage_rmdup["AVG_DEPTH"])*100
		del coverage_rmdup["PERC_COVERAGE@100"]
		for row in dataframe_to_rows(coverage_rmdup, header=True, index=False):
			ws4.append(row)

		#highlight the cells in the percentage NTC contamination column red if the value is greater than 10%
		colour= PatternFill("solid", start_color="00FF0000", end_color="00FF0000")
		ws4.conditional_formatting.add('G2:G300', CellIsRule( operator='greaterThan', formula=['10'], stopIfTrue=True, fill=colour))
	else:
		print("Coverage rmdup file does not exist")


	ws3.column_dimensions['B'].width=15
	ws4.column_dimensions['B'].width=15
	ws3.column_dimensions['C'].width=15
	ws4.column_dimensions['C'].width=15
	ws3.column_dimensions['D'].width=110
	ws4.column_dimensions['D'].width=110

	return(coverage_rmdup)



def format_analysis_sheet(referral, coverage_rmdup):


	ws1['A5']=sampleId
	ws1['B5']='=Patient_demographics!C2'
	ws1['C5']=referral
	ws1['D5']=worksheetId
	ws1['E5']=seqId
	ws1['C1']="Analysis sheet- Pan RNA Cancer Fusion Panel"


	ws1['S10']="Conclusion checker 1"
	ws1['T10']="Conclusion checker 2"
	ws1['U10']="Comments"
	ws1['Y24']="Conclusion checker 1"
	ws1['Z24']="Conclusion checker 2"
	ws1['AA24']="Comments"

	dv = DataValidation(type="list", formula1='"Genuine,Artefact,Fail_QC,no fusions"', allow_blank=True)
	ws1.add_data_validation(dv)

	dv.add('S11:S21')
	dv.add('T11:T21')
	dv.add('Y25:Y35')
	dv.add('Z21:Z35')

	font_bold=Font(bold=True)
	position= ['A4', 'B4', 'C4', 'D4', 'E4', 'F4', 'G4', 'A39',
	'A10','B10','C10','D10','E10','F10','G10','H10','I10','J10','K10','L10','M10','N10','O10','P10','Q10','R10','S10','T10', 'U10',
	'A24','B24','C24','D24','E24','F24','G24','H24','I24','J24','K24','L24','M24','N24','O24','P24','Q24','R24','S24','T24','U24','V24','W24','X24','Y24', 'Z24', 'AA24']
	for cell in position:
		ws1[cell].font=font_bold


	for column in ['A','B', 'C','D','E', 'F','G', 'H', 'I', 'J','K', 'L']:
		ws5.column_dimensions[column].width=25

	for column in ['A','B', 'C','D','E', 'F','G', 'H', 'I', 'J','K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y']:
		ws1.column_dimensions[column].width=20	

	border_a=Border(left=Side(border_style=BORDER_MEDIUM), right=Side(border_style=BORDER_MEDIUM), top=Side(border_style=BORDER_MEDIUM), bottom=Side(border_style=BORDER_MEDIUM))
	position=['A4','B4','C4','D4','E4','F4','G4','A5','B5','C5','D5','E5','F5','G5','A10','B10','C10','D10','E10','F10','G10','H10','I10','J10','K10','L10','M10','N10','O10','P10','Q10','R10','S10','T10','U10',
	'A24','B24','C24','D24','E24','F24','G24','H24','I24','J24','K24','L24','M24','N24','O24','P24','Q24','R24','S24','T24','U24','V24','W24','X24','Y24', 'Z24', 'AA24']
	for cell in position:
		ws1[cell].border=border_a

	colour= PatternFill("solid", fgColor="DCDCDC")
	position= ['A10','B10','C10','D10','E10','F10','G10','H10','I10','J10','K10','L10','M10','N10','O10','P10','Q10','R10',
	'A24','B24','C24','D24','E24','F24','G24','H24','I24','J24','K24','L24','M24','N24','O24','P24','Q24','R24','S24','T24','U24','V24','W24','X24']
	for cell in position:
		ws1[cell].fill=colour


	colour= PatternFill("solid", fgColor="00FFCC00")
	position= ['F4', 'G4', 'S10', 'T10', 'U10', 'Y24', 'Z24', 'AA24']
	for cell in position:
		ws1[cell].fill=colour

	colour= PatternFill("solid", fgColor="00CCFFFF")
	position= ['A4','B4','C4','D4','E4']
	for cell in position:
		ws1[cell].fill=colour


	colour= PatternFill("solid", fgColor="DCDCDC")
	position= ['I7', 'I8']
	for cell in position:
		ws2[cell].fill=colour


	colour= PatternFill("solid", fgColor="00CCFFFF")
	position= ['J6', 'K6']
	for cell in position:
		ws2[cell].fill=colour


	ws2['G3']="Total mapped reads (including duplicates)"
	ws2['I7']="NTC check 1"
	ws2['I8']="NTC check 2"
	ws2['J6']="Result"
	ws2['K6']= "Comments"


	dv = DataValidation(type="list", formula1='"PASS,FAIL,PARTIAL_FAIL"', allow_blank=True)
	ws2.add_data_validation(dv)
	dv.add(ws2['J7'])
	dv.add(ws2['J8'])


	border_a=Border(left=Side(border_style=BORDER_MEDIUM), right=Side(border_style=BORDER_MEDIUM), top=Side(border_style=BORDER_MEDIUM), bottom=Side(border_style=BORDER_MEDIUM))
	position=['I7', 'I8', 'J6', 'J7', 'J8', 'K6', 'K7', 'K8']
	for cell in position:
		ws2[cell].border=border_a


	ws5['C1']='Analysis Sheet-Pan RNA Cancer Fusion Panel Validation'
	ws5['C1'].font=Font(size=16)
	ws5["A4"]="Lab number" 
	ws5["B4"]="Patient name"
	ws5["C4"]="Tumour %"
	ws5["D4"]="Reason for Referral"
	ws5["E4"]="Qubit RNA conc after pre-processing (ng/ul)"
	ws5["F4"]="DV200"
	ws5["G4"]="% mapped reads-duplicates removed"
	ws5["H4"]="Due date"

	ws5["A7"]="NGS wks"
	ws5["B7"]="NextSeq run ID"
	ws5["C7"]="NTC check 1 result"
	ws5["D7"]="NTC check 2 result"
	ws5["E7"]="Analysed by:"
	ws5["F7"]="Checked by:"


	ws5["A12"]= "Gene fusion genuine calls"
	ws5["A13"]= "STAR-Fusion results"
	
	ws5["A14"]="Fusion_Name"
	ws5["B14"]="Split_Read_Count"
	ws5["C14"]="Spanning_Read_Count"
	ws5["D14"]="Left_Breakpoint"
	ws5["E14"]="Right_Breakpoint:"
	ws5["F14"]="SpliceType"
	ws5["G14"]="LargeAnchorSupport"
	ws5["H14"]="FFPM"
	ws5["I14"]="Prot_Fusion_Type"
	ws5["J14"]="Fusion_Allelic_Fraction"
	ws5["K14"]="Conclusion checker 1"
	ws5["L14"]="Conclusion checker 2"
	ws5["M14"]="Comments"
	
	ws5["A27"]="Arriba results"
	ws5["A28"]="#gene1"
	ws5["B28"]="gene2"
	ws5["C28"]="strand1(gene/fusion)"
	ws5["D28"]="strand2(gene/fusion)"
	ws5["E28"]="breakpoint1"
	ws5["F28"]="breakpoint2"
	ws5["G28"]="site1"
	ws5["H28"]="site2"
	ws5["I28"]="type"
	ws5["J28"]="split_reads1"
	ws5["K28"]="split_reads2"
	ws5["L28"]="discordant_mates"
	ws5["M28"]="confidence"
	ws5["N28"]="filters"
	ws5["O28"]="Conclusion checker 1"
	ws5["P28"]="Conclusion checker 2"
	ws5["Q28"]="Comments"

	ws5["A42"]="Gene fusions quality metrics"
	ws5["A44"]="Post PCR1 Qubit"
	ws5["A45"]="Total mapped reads-duplicates removed"
	ws5["B43"]="Value"
	ws5["C43"]="Conclusion checker 1"
	ws5["D43"]="Conclusion checker 2"



	
	border_a=Border(left=Side(border_style=BORDER_MEDIUM), right=Side(border_style=BORDER_MEDIUM), top=Side(border_style=BORDER_MEDIUM), bottom=Side(border_style=BORDER_MEDIUM))
	position=['A4','B4','C4','D4','E4','F4','G4', 'H4',
	'A5','B5','C5','D5','E5','F5', 'G5','H5',
        'A7','B7','C7','D7','E7','F7',
        'A8','B8','C8','D8','E8','F8',
        'A14','B14','C14','D14','E14','F14','G14', 'H14', 'I14', 'J14', 'K14', 'L14', 'M14',
        'A28','B28','C28','D28','E28','F28','G28', 'H28', 'I28', 'J28', 'K28', 'L28', 'M28', 'N28', 'O28', 'P28', 'Q28',
        'B43', 'C43', 'D43',
        'A44', 'B44', 'C44', 'D44',
        'A45', 'B45', 'C45', 'D45']
	for cell in position:
		ws5[cell].border=border_a


	colour= PatternFill("solid", fgColor="DCDCDC")
	position= ['A14', 'B14', 'C14', 'D14', 'E14', 'F14', 'G14', 'H14', 'I14', 'J14', 
                   'A28', 'B28', 'C28', 'D28', 'E28', 'F28', 'G28', 'H28', 'I28', 'J28', 'K28', 'L28' , 'M28', 'N28',
                   'B43', 'A44', 'A45']
	for cell in position:
		ws5[cell].fill=colour


	colour= PatternFill("solid", fgColor="00CCFFFF")
	position= ['A4','B4','C4','D4','E4','F4', 'G4', 'H4', 'A7','B7']
	for cell in position:
		ws5[cell].fill=colour



	colour= PatternFill("solid", fgColor="00FFCC00")
	position= ['C7', 'D7', 'E7', 'F7', 'K14', 'L14', 'M14', 'O28', 'P28', 'Q28','C43', 'D43']
	for cell in position:
		ws5[cell].fill=colour


	font_bold=Font(bold=True)
	position= ['A4','B4','C4','D4','E4','F4','G4','H4','A7','B7','C7','D7','E7','F7','A14','B14','C14','D14','E14','F14','G14','H14','I14', 'J14', 'K14','L14', 'M14',
		'A28','B28','C28','D28','E28','F28','G28','H28','I28', 'J28', 'K28', 'L28', 'M28', 'N28', 'O28', 'P28', 'Q28',
		'B43', 'C43', 'D43', 'A44', 'A45',
		'A53','B53','C53','D53','E53','F53','G53','H53','I53', 'J53', 'K53',
		'A59','B59','C59','D59','E59','F59','G59',
		'A12', 'A13', 'A27', 'A42', 'A52', 'A58']
	for cell in position:
		ws5[cell].font=font_bold


	ws5['A5']=sampleId
	ws5['B5']= '=Patient_demographics!C2'
	ws5['C5']= '=Patient_demographics!I2'
	ws5['D5']=referral
	ws5['E5']= '=Patient_demographics!L2'
	ws5['F5']= '=Patient_demographics!J2'
	ws5['H5']= '=Patient_demographics!H2'
	ws5['H5'].number_format='mm-dd-yy'


	ws5['A8']=worksheetId
	ws5['B8']=seqId
	ws5['C8']='=total_coverage!J7'
	ws5['D8']='=total_coverage!J8'
	ws5['B44']='=Patient_demographics!O2'


	#add pass, fail dropdown boxes to the summary sheet
	dv = DataValidation(type="list", formula1='"PASS,FAIL"', allow_blank=True)
	ws5.add_data_validation(dv)
	dv.add('C44:C45')
	dv.add('D44:D45')
	dv.add('F60:F63')
	dv.add('G60:G63')


	#transfer values in star-fusion table of gene fusion report to the summary tab

	summary_column_list=['A', 'B','C','D','E','F','G','H','I','J', 'K', 'L', 'M']
	fusion_report_column_list=[ 'A','B','C','D','E','F','G','H','O', 'R', 'S', 'T', 'U']

	value=0
	while (value<(len(fusion_report_column_list))):
		fusion_report_column=fusion_report_column_list[value]
		summary_column=summary_column_list[value]

		ws5[summary_column+'15']= '=IF(OR(Gene_fusion_report!S11<>"Artefact",Gene_fusion_report!T11<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'11, "0")'
		ws5[summary_column+'16']= '=IF(OR(Gene_fusion_report!S12<>"Artefact",Gene_fusion_report!T12<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'12, "0")'
		ws5[summary_column+'17']= '=IF(OR(Gene_fusion_report!S13<>"Artefact",Gene_fusion_report!T13<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'13, "0")'
		ws5[summary_column+'18']= '=IF(OR(Gene_fusion_report!S14<>"Artefact",Gene_fusion_report!T14<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'14, "0")'
		ws5[summary_column+'19']= '=IF(OR(Gene_fusion_report!S15<>"Artefact",Gene_fusion_report!T15<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'15, "0")'
		ws5[summary_column+'20']= '=IF(OR(Gene_fusion_report!S16<>"Artefact",Gene_fusion_report!T16<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'16, "0")'
		ws5[summary_column+'21']= '=IF(OR(Gene_fusion_report!S17<>"Artefact",Gene_fusion_report!T17<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'17, "0")'
		ws5[summary_column+'22']= '=IF(OR(Gene_fusion_report!S18<>"Artefact",Gene_fusion_report!T18<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'18, "0")'
		ws5[summary_column+'23']= '=IF(OR(Gene_fusion_report!S19<>"Artefact",Gene_fusion_report!T19<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'19, "0")'
		ws5[summary_column+'24']= '=IF(OR(Gene_fusion_report!S20<>"Artefact",Gene_fusion_report!T20<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'20, "0")'
		ws5[summary_column+'25']= '=IF(OR(Gene_fusion_report!S21<>"Artefact",Gene_fusion_report!T21<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'21, "0")'
		ws5[summary_column+'26']= '=IF(OR(Gene_fusion_report!S22<>"Artefact",Gene_fusion_report!T22<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'22, "0")'

		value=value+1


	#transfer values in arriba table of gene fusion report to the summary tab

	summary_column_list=['A', 'B','C','D','E','F','G','H','I','J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']
	fusion_report_column_list=[ 'A','B','C','D','E','F','G','H','I', 'L', 'M', 'N', 'Q', 'T','Y','Z','AA']

	value=0
	while (value<(len(fusion_report_column_list))):
		fusion_report_column=fusion_report_column_list[value]
		summary_column=summary_column_list[value]



		ws5[summary_column+'29']= '=IF(OR(Gene_fusion_report!Y25<>"Artefact",Gene_fusion_report!Z25<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'25, "")'
		ws5[summary_column+'30']= '=IF(OR(Gene_fusion_report!Y26<>"Artefact",Gene_fusion_report!Z26<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'26, "")'
		ws5[summary_column+'31']= '=IF(OR(Gene_fusion_report!Y27<>"Artefact",Gene_fusion_report!Z27<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'27, "")'
		ws5[summary_column+'32']= '=IF(OR(Gene_fusion_report!Y28<>"Artefact",Gene_fusion_report!Z28<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'28, "")'
		ws5[summary_column+'33']= '=IF(OR(Gene_fusion_report!Y29<>"Artefact",Gene_fusion_report!Z29<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'29, "")'
		ws5[summary_column+'34']= '=IF(OR(Gene_fusion_report!Y30<>"Artefact",Gene_fusion_report!Z30<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'30, "")'
		ws5[summary_column+'35']= '=IF(OR(Gene_fusion_report!Y31<>"Artefact",Gene_fusion_report!Z31<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'31, "")'
		ws5[summary_column+'36']= '=IF(OR(Gene_fusion_report!Y32<>"Artefact",Gene_fusion_report!Z32<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'32, "")'
		ws5[summary_column+'37']= '=IF(OR(Gene_fusion_report!Y33<>"Artefact",Gene_fusion_report!Z33<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'33, "")'
		ws5[summary_column+'38']= '=IF(OR(Gene_fusion_report!Y34<>"Artefact",Gene_fusion_report!Z34<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'34, "")'
		ws5[summary_column+'39']= '=IF(OR(Gene_fusion_report!Y35<>"Artefact",Gene_fusion_report!Z35<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'35, "")'
		ws5[summary_column+'40']= '=IF(OR(Gene_fusion_report!Y36<>"Artefact",Gene_fusion_report!Z36<>"Artefact"),Gene_fusion_report!'+fusion_report_column+'36, "")'

		value=value+1


	font_bold=Font(bold=True)
	position= ['A1','B1','C1','D1','E1','F1','G1','H1','I1','J1','K1','L1','M1','N1','O1','P1','Q1','R1','S1','T1','U1','V1']
	for cell in position:
		ws7[cell].font=font_bold


	colour= PatternFill("solid", fgColor="DCDCDC")
	position= ['A1','D1','F1','G1','H1','I1','J1','K1','L1','M1','N1','O1','P1','Q1','R1','S1','T1','U1','V1']
	for cell in position:
		ws7[cell].fill=colour

	colour= PatternFill("solid", fgColor="00FFFF00")
	position= ['B1', 'C1', 'E1', 'B2', 'C2', 'E2', 'O1', 'O2']
	for cell in position:
		ws7[cell].fill=colour


	colour= PatternFill("solid", fgColor="00FFCC00")
	position= ['A6', 'A7']
	for cell in position:
		ws7[cell].fill=colour


	border_a=Border(left=Side(border_style=BORDER_MEDIUM), right=Side(border_style=BORDER_MEDIUM), top=Side(border_style=BORDER_MEDIUM), bottom=Side(border_style=BORDER_MEDIUM))
	position=['A1','B1','C1','D1','E1','F1','G1','H1','I1','J1','K1','L1','M1','N1','O1','P1','Q1','R1','S1','T1','U1','V1', 'A6', 'A7', 'B6', 'B7']
	for cell in position:
		ws7[cell].border=border_a


	ws7['B2']=sampleId
	ws7['E2']=referral
	ws7['M2']=worksheetId
	ws7['N2']=seqId


	ws7.column_dimensions['A'].width=30
	ws7.column_dimensions['B'].width=20
	ws7.column_dimensions['C'].width=30
	ws7.column_dimensions['N'].width=40



	


def get_alignment_metrics_rmdup():

	#get the percentage aligned reads rmdup and number of aligned reads rmdup
	with open ("../"+sampleId+"_AlignmentSummaryMetrics_rmdup.txt") as file_rmdup:
		for line in file_rmdup:
			if line.startswith("CATEGORY"):
				headers_rmdup=line.split('\t')
			if line.startswith("PAIR"):
				pair_list_rmdup=line.split('\t')


	alignment_metrics_rmdup=pandas.DataFrame([pair_list_rmdup], columns=headers_rmdup)
	total_reads_rmdup=alignment_metrics_rmdup[['PF_READS']]
	total_reads_aligned_rmdup=alignment_metrics_rmdup[['PF_HQ_ALIGNED_READS']]
	total_reads_value_rmdup=total_reads_rmdup.iloc[0,0]
	aligned_reads_value_rmdup=total_reads_aligned_rmdup.iloc[0,0]
	percentage_aligned_reads_rmdup=(int(aligned_reads_value_rmdup)/int(total_reads_value_rmdup))*100
	ws5["G5"]=percentage_aligned_reads_rmdup
	aligned_reads_value_rmdup=int(aligned_reads_value_rmdup)
	ws5["B45"]=aligned_reads_value_rmdup

	#colour the cell red if number of aligned reads rmdup is less than the threshold of 38000
	if (aligned_reads_value_rmdup<38000):
		colour= PatternFill("solid", fgColor="00FF0000")
		position= ['B45']
		for cell in position:
			ws5[cell].fill=colour
		 
	



def get_met_exon_skipping(referral_list):

	ws8=wb.create_sheet("RMATS")

	ws8['A4']='Lab number'
	ws8['B4']='Patient Name'
	ws8['C4']='Reason for referral'
	ws8['D4']='NGS Wks'
	ws8['E4']='Nextseq run ID'
	ws8['F4']='Checker 1 initials and date'
	ws8['G4']='Checker 2 initials and date'

	ws8['A8']='RMATS output'

	ws8['C1']='Analysis Sheet- Pan RNA Cancer Fusion Panel Validation'

	ws8['A5']=sampleId
	ws8['B5']='=Patient_demographics!C2'
	ws8['C5']=referral
	ws8['D5']=worksheetId
	ws8['E5']=seqId

	ws8["A9"]="GeneID"
	ws8["B9"]="geneSymbol"
	ws8["C9"]="chr"
	ws8["D9"]="exonStart_0base"
	ws8["E9"]="exonEnd"
	ws8["F9"]="IJC_SAMPLE_1"
	ws8["G9"]="SJC_SAMPLE_1"
	ws8["H9"]="IJC_SAMPLE_2"
	ws8["I9"]="SJC_SAMPLE_2"
	ws8["J9"]="FDR"
	ws8["K9"]="IncLevel1"
	ws8["L9"]="IncLevel2"
	ws8["M9"]="IncLevelDifference"
	ws8["N9"]="Conclusion checker 1"
	ws8["O9"]="Conclusion checker 2"
	ws8["P9"]="Comments"

	ws5["A52"]="Intragenic fusions(RMATS) genuine calls"
	ws5['A58']="Intragenic fusions quality metrics"
	ws5["A53"]="geneSymbol"
	ws5["B53"]="chr"
	ws5["C53"]="exonStart_0base"
	ws5["D53"]="exonEnd"
	ws5["E53"]="IJC_SAMPLE_1"
	ws5["F53"]="SJC_SAMPLE_1"
	ws5["G53"]="FDR"
	ws5["H53"]="IncLevelDifference"
	ws5["I53"]="Conclusion Checker 1"
	ws5["J53"]="Conclusion Checker 2"
	ws5["K53"]="Comments"


	ws5["A59"]="CHR"
	ws5["B59"]="START"
	ws5["C59"]="END"
	ws5["D59"]="AVG_DEPTH"
	ws5["E59"]="Gene_exon"
	ws5["F59"]="Conclusion 1"
	ws5["G59"]="Conclusion 2"



	colour= PatternFill("solid", fgColor="DCDCDC")
	position= ['A53', 'B53', 'C53', 'D53', 'E53', 'F53', 'G53', 'H53', 
                   'A59', 'B59', 'C59', 'D59', 'E59' ]
	for cell in position:
		ws5[cell].fill=colour

	colour= PatternFill("solid", fgColor="00FFCC00")
	position= ['I53', 'J53', 'K53', 'F59', 'G59']
	for cell in position:
		ws5[cell].fill=colour

	
	border_a=Border(left=Side(border_style=BORDER_MEDIUM), right=Side(border_style=BORDER_MEDIUM), top=Side(border_style=BORDER_MEDIUM), bottom=Side(border_style=BORDER_MEDIUM))
	position=['A53','B53','C53','D53','E53','F53','G53', 'H53', 'I53', 'J53', 'K53',
        'A59','B59','C59','D59','E59','F59','G59']
	for cell in position:
		ws5[cell].border=border_a



	#add the rmats report to the RMATS report tab. Only add exon skipping events for genes that are in the referral type.
	rmats=[]
	headers=[]
	line_number=0

	with open (seqId+"_"+sampleId+"_RMATS_Report.tsv") as file:
		for line in file:
			if line.startswith("#"):
				continue
			else:
				if (line_number==0):
					headers=line.split('\t')
					line_number=1
				else:	
					rmats_line=line.split('\t')
					rmats.append(rmats_line)
	rmats_dataframe = pandas.DataFrame(rmats, columns=headers)
	rmats_dataframe=rmats_dataframe[['GeneID', 'geneSymbol', 'chr', 'exonStart_0base', 'exonEnd', 'IJC_SAMPLE_1', 'SJC_SAMPLE_1', 'IJC_SAMPLE_2', 'SJC_SAMPLE_2', 'FDR', 'IncLevel1', 'IncLevel2', 'IncLevelDifference']]




	for gene in referral_list:
		if (gene=="MET_exon14_skipping"):
			rmats_MET=rmats_dataframe[rmats_dataframe["geneSymbol"]=="MET"]
			if (len(rmats_MET)>0):
				for row in dataframe_to_rows(rmats_MET, header=False, index=False):
					ws8.append(row)
			else:
				ws8['B10']="MET ex14 skipping-No fusions called"
		if (gene=="EGFRv3"):
			rmats_EGFR=rmats_dataframe[rmats_dataframe["geneSymbol"]=="EGFR"]
			if (len(rmats_EGFR)>0):
				for row in dataframe_to_rows(rmats_EGFR, header=False, index=False):
					ws8.append(row)
			else:
				ws8['B11']="EGFR v3-No fusions called"




	#add dropdown box in the conclusion columns of the RMATS report
	dv = DataValidation(type="list", formula1='"Genuine,Artefact,Fail_QC,no_fusions"', allow_blank=True)
	ws8.add_data_validation(dv)
	dv.add('N10:N12')
	dv.add('O10:O12')


	#transfer values in RMATS table in RMATS tab to the summary tab
	summary_column_list=['A', 'B','C','D','E','F','G','H','I','J', 'K']
	RMATS_column_list=[ 'B','C','D','E','F','G','J','M','N', 'O', 'P']
	value=0
	while (value<(len(RMATS_column_list))):
		RMATS_column=RMATS_column_list[value]
		summary_column=summary_column_list[value]

		ws5[summary_column+'54']= '=IF(OR(RMATS!N10<>"Artefact",RMATS!O10<>"Artefact"),RMATS!'+RMATS_column+'10, "0")'
		ws5[summary_column+'55']= '=IF(OR(RMATS!N11<>"Artefact",RMATS!O11<>"Artefact"),RMATS!'+RMATS_column+'11, "0")'
		ws5[summary_column+'56']= '=IF(OR(RMATS!N12<>"Artefact",RMATS!O12<>"Artefact"),RMATS!'+RMATS_column+'12, "0")'
		value=value+1

	

	ws8.column_dimensions['A'].width=30
	ws8.column_dimensions['B'].width=15
	ws8.column_dimensions['C'].width=15
	ws8.column_dimensions['H'].width=70
	ws8.column_dimensions['I'].width=40
	ws8.column_dimensions['L'].width=40

	for column in ['D','E', 'F','G','J', 'k','M', 'N', 'O', 'P','Q']:
		ws8.column_dimensions[column].width=20


	font_bold=Font(bold=True)
	position= ['A4','B4','C4','D4','E4','F4','G4','A8','B8','C8','D8','E8','F8','G8','H8','I8','J8','K8','L8','M8','N8','O8','P8','Q8']
	for cell in position:
		ws8[cell].font=font_bold


	border_a=Border(left=Side(border_style=BORDER_MEDIUM), right=Side(border_style=BORDER_MEDIUM), top=Side(border_style=BORDER_MEDIUM), bottom=Side(border_style=BORDER_MEDIUM))
	position=['A4','B4','C4','D4','E4','F4','G4','A5','B5','C5','D5','E5','F5','G5','A9','B9','C9','D9','E9','F9','G9','H9','I9','J9','K9','L9','M9','N9','O9','P9']
	for cell in position:
		ws8[cell].border=border_a

	colour= PatternFill("solid", fgColor="DCDCDC")
	position= ['A9','B9','C9','D9','E9','F9','G9','H9','I9','J9','K9','L9','M9','N9','O9','P9']
	for cell in position:
		ws8[cell].fill=colour


	colour= PatternFill("solid", fgColor="00CCFFFF")
	position= ['A4','B4','C4','D4','E4','F4', 'G4']
	for cell in position:
		ws8[cell].fill=colour

	colour= PatternFill("solid", fgColor="00FFCC00")
	position= ['N9', 'O9', 'P9']
	for cell in position:
		ws8[cell].fill=colour

	
	ws6['A33']= "RMATS results"
	ws6["A34"]="GeneID"
	ws6["B34"]="geneSymbol"
	ws6["C34"]="chr"
	ws6["D34"]="exonStart_0base"
	ws6["E34"]="exonEnd"
	ws6["F34"]="IJC_SAMPLE_1"
	ws6["G34"]="SJC_SAMPLE_1"
	ws6["H34"]="IJC_SAMPLE_2"
	ws6["I34"]="SJC_SAMPLE_2"
	ws6["J34"]="FDR"
	ws6["K34"]="IncLevel1"
	ws6["L34"]="IncLevel2"
	ws6["M34"]="IncLevelDifference"
	ws6["N34"]="Conclusion checker 1"
	ws6["O34"]="Conclusion checker 2"
	ws6["P34"]="Comments"



	#add the rmats ntc report to the ntc report tab. Only add exon skipping events for genes that are in the referral type.
	rmats_ntc=[]
	headers=[]
	line_number=0

	with open ("../"+ntc+"/"+seqId+"_"+ntc+"_RMATS_Report.tsv") as file:
		for line in file:
			if line.startswith("#"):
				continue
			else:
				if (line_number==0):
					headers=line.split('\t')
					line_number=1
				else:	
					rmats_ntc_line=line.split('\t')
					rmats_ntc.append(rmats_ntc_line)
	rmats_ntc_dataframe = pandas.DataFrame(rmats_ntc, columns=headers)
	rmats_ntc_dataframe=rmats_ntc_dataframe[['GeneID', 'geneSymbol', 'chr', 'exonStart_0base', 'exonEnd', 'IJC_SAMPLE_1', 'SJC_SAMPLE_1', 'IJC_SAMPLE_2', 'SJC_SAMPLE_2', 'FDR', 'IncLevel1', 'IncLevel2', 'IncLevelDifference']]
	for gene in referral_list:
		if (gene=="MET_exon14_skipping"):
			rmats_ntc_MET=rmats_ntc_dataframe[rmats_ntc_dataframe["geneSymbol"]=="MET"]
			if (len(rmats_ntc_MET)>0):
				for row in dataframe_to_rows(rmats_ntc_MET, header=False, index=False):
					ws8.append(row)
			else:
				ws6['B35']="MET ex14 skipping-No fusions called"
		if (gene=="EGFRv3"):
			rmats_ntc_EGFR=rmats_ntc_dataframe[rmats_ntc_dataframe["geneSymbol"]=="EGFR"]
			if (len(rmats_ntc_EGFR)>0):
				for row in dataframe_to_rows(rmats_ntc_EGFR, header=False, index=False):
					ws8.append(row)
			else:
				ws6['B36']="EGFR v3-No fusions called"


#transfer the coverage values for EGFRv3 and MET_exon14_skipping events breakpoints from the coverage report into the summary tab
      
	MET_exon_13=coverage_rmdup[coverage_rmdup["START"]==116411698]
	del [MET_exon_13["META"], MET_exon_13["NTC_AVG_DEPTH"],MET_exon_13["%NTC contamination"]]
		 
	MET_exon_15=coverage_rmdup[coverage_rmdup["END"]==116414944]
	del [MET_exon_15["META"],MET_exon_15["NTC_AVG_DEPTH"],MET_exon_15["%NTC contamination"]]

	EGFR_exon_1=coverage_rmdup[coverage_rmdup["START"]==55087048]
	del [EGFR_exon_1["META"], EGFR_exon_1["NTC_AVG_DEPTH"],EGFR_exon_1["%NTC contamination"]]

	EGFR_exon_8=coverage_rmdup[coverage_rmdup["END"]==55223532]
	del [EGFR_exon_8["META"], EGFR_exon_8["NTC_AVG_DEPTH"], EGFR_exon_8["%NTC contamination"]]


	#If the coverage over the MET_exon14_skipping and EGFRv3 breakpoints are greater than the thresholds, highlight them in red.

	for row in dataframe_to_rows(MET_exon_13, header=False, index=False):
		ws5.append(row)
	MET_exon_13_AVGdepth=MET_exon_13[["AVG_DEPTH"]]
	if (MET_exon_13_AVGdepth.iloc[0,0]<7):
		colour= PatternFill("solid", fgColor="00FF0000")
		position= ['D60']
		for cell in position:
			ws5[cell].fill=colour


	for row in dataframe_to_rows(MET_exon_15, header=False, index=False):
		ws5.append(row)
	MET_exon_15_AVGdepth=MET_exon_15[["AVG_DEPTH"]]
	print(MET_exon_15_AVGdepth)
	if (MET_exon_15_AVGdepth.iloc[0,0]<6):
		colour= PatternFill("solid", fgColor="00FF0000")
		position= ['D61']
		for cell in position:
			ws5[cell].fill=colour


	for row in dataframe_to_rows(EGFR_exon_1, header=False, index=False):
		ws5.append(row)
	EGFR_exon_1_AVGdepth=EGFR_exon_1[["AVG_DEPTH"]]
	if (EGFR_exon_1_AVGdepth.iloc[0,0]<6):
		colour= PatternFill("solid", fgColor="00FF0000")
		position= ['D62']
		for cell in position:
			ws5[cell].fill=colour


	for row in dataframe_to_rows(EGFR_exon_8, header=False, index=False):
		ws5.append(row)
	EGFR_exon_8_AVGdepth=EGFR_exon_8[["AVG_DEPTH"]]
	print(EGFR_exon_8_AVGdepth)
	if (EGFR_exon_8_AVGdepth.iloc[0,0]<9):
		colour= PatternFill("solid", fgColor="00FF0000")
		position= ['D63']
		for cell in position:
			ws5[cell].fill=colour


	ws5["E60"]="MET_Exon13_last10bases"
	ws5["E61"]="MET_Exon15_first10bases"
	ws5["E62"]="EGFR_Exon1_last10bases"
	ws5["E63"]="EGFR_Exon8_first10bases"



		


if __name__ == "__main__":

	seqId=sys.argv[1]
	sampleId=sys.argv[2]
	referral=sys.argv[3]
	ntc=sys.argv[4]
	worksheetId=sys.argv[5]
	version=sys.argv[6]


	referral_list=get_referral_list(referral)
	ntc_star_fusion_report, ntc_arriba_report=get_NTC_fusion_report(ntc)
	get_star_fusion_report(referral_list, ntc_star_fusion_report)
	get_arriba_fusion_report(referral_list)
	ntc_total_average_depth, ntc_total_average_depth_rmdup=get_ntc_total_coverage()
	get_total_coverage(ntc_total_average_depth, ntc_total_average_depth_rmdup)
	ntc_average_depth, ntc_average_depth_rmdup=ntc_get_coverage()
	coverage_rmdup=get_coverage(referral, ntc_average_depth, ntc_average_depth_rmdup)
	format_analysis_sheet(referral, coverage_rmdup)
	get_alignment_metrics_rmdup()

	if (("MET_exon14_skipping" in referral_list) or ("EGFRv3" in referral_list)):
		get_met_exon_skipping(referral_list)
	wb.save("../"+sampleId+"_"+referral +"-TESTER.xlsx")
