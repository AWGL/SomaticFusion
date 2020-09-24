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

wb=Workbook()
ws7=wb.create_sheet("Patient_demographics")
ws6= wb.create_sheet("NTC fusion report")
ws2=wb.create_sheet("total_coverage")
ws1= wb.create_sheet("Gene_fusion_report")
ws8=wb.create_sheet("RMATS")
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



def get_referral_list(referral):
	
	referral_file=pandas.read_csv("/data/diagnostics/pipelines/SomaticFusion/SomaticFusion-0.0.4/Referrals/"+referral+".txt", sep="\t")
	referral_list=list(referral_file['Genes'])
	return(referral_list)


def get_NTC_fusion_report(ntc):
	
	ws6["A4"]="NTC STAR-Fusion results"
	for referral in referral_list:

		if (os.stat("../"+ntc+"/Results/STAR_Fusion/"+ntc+"_fusion_report_"+referral+".txt").st_size!=0):
      
			NTC_star_fusion_report=pandas.read_csv("../"+ntc+"/Results/STAR_Fusion/"+ntc+"_fusion_report_"+referral+".txt", sep="\t")

			NTC_star_fusion_report=pandas.DataFrame(NTC_star_fusion_report)
	

			for row in dataframe_to_rows(NTC_star_fusion_report, header=True, index=False):
				ws6.append(row)


	ws6["A19"]="Arriba results"
	for referral in referral_list:
	
		if (os.stat("../"+ntc+"/Results/arriba/"+ntc+"_fusion_report_"+referral+"_arriba.txt").st_size!=0):
      	
			NTC_arriba_report=pandas.read_csv("../"+ntc+"/Results/arriba/"+ntc+"_fusion_report_"+referral+"_arriba.txt", sep="\t")

	
			for row in dataframe_to_rows(NTC_arriba_report, header=True, index=False):
				ws6.append(row)


	return(NTC_star_fusion_report, NTC_arriba_report)

	
		




def get_star_fusion_report(referral_list, ntc_star_fusion_report):
	ws1["A9"]="STAR-Fusion results"



	#create empty pandas dataframe and append

	star_fusion_report_final= pandas.DataFrame(columns=["Fusion_Name", "Split_Read_Count", "Spanning_Read_Count", "Left_Breakpoint", "Right_Breakpoint", "SpliceType", "LargeAnchorSupport", "FFPM", "LeftBreakEntropy", "RightBreakEntropy","CDS_Left_ID", "CDS_Left_Range", "CDS_Right_ID", "CDS_Right_Range", "Prot_Fusion_Type","Num_WT_Fragments_Left", "Num_WT_Fragments_Right", "Fusion_Allelic_Fraction"]) 

	for referral in referral_list:

		if (referral!="MET"):


			if (os.stat("./Results/STAR_Fusion/"+sampleId+"_fusion_report_"+referral+".txt").st_size!=0):
      
				star_fusion_report=pandas.read_csv("Results/STAR_Fusion/"+sampleId+"_fusion_report_"+referral+".txt", sep="\t")	
				if len(star_fusion_report!=0):
					star_fusion_report=pandas.DataFrame(star_fusion_report)
					del star_fusion_report['Unnamed: 0']
				else:
					dict_star_fusion=({"Fusion_Name": referral+"-no fusions found", "Split_Read_Count": "", "Spanning_Read_Count": "", "Left_Breakpoint": "", "Right_Breakpoint": "", "SpliceType": "", "LargeAnchorSupport": "", "FFPM": "", "LeftBreakEntropy": "", "RightBreakEntropy": "","CDS_Left_ID": "", "CDS_Left_Range": "", "CDS_Right_ID": "", "CDS_Right_Range": "", "Prot_Fusion_Type": "","Num_WT_Fragments_Left": "", "Num_WT_Fragments_Right": "", "Fusion_Allelic_Fraction": ""}) 
					star_fusion_report= pandas.DataFrame(dict_star_fusion, columns=["Fusion_Name", "Split_Read_Count", "Spanning_Read_Count", "Left_Breakpoint", "Right_Breakpoint", "SpliceType", "LargeAnchorSupport", "FFPM", "LeftBreakEntropy", "RightBreakEntropy","CDS_Left_ID", "CDS_Left_Range", "CDS_Right_ID", "CDS_Right_Range", "Prot_Fusion_Type","Num_WT_Fragments_Left", "Num_WT_Fragments_Right", "Fusion_Allelic_Fraction"], index=[0]) 	


			star_fusion_report_final=star_fusion_report_final.append(star_fusion_report, sort=False)
	

	for row in dataframe_to_rows(star_fusion_report_final, header=True, index=False):
		ws1.append(row)


	
		




def get_arriba_fusion_report(referral_list):

	ws1["A23"]="Arriba results"

	arriba_report_final= pandas.DataFrame(columns=["gene1", "gene2", "strand1(gene/fusion)", "strand2(gene/fusion)", "breakpoint1", "breakpoint2", "site1", "site2", "type", "direction1", "direction2", "split_reads1", "split_reads2", "discordant_mates", "coverage1", "coverage2", "confidence", "closest_genomic_breakpoint1", "closest_genomic_breakpoint2", "filters", "fusion_transcript", "reading_frame", "peptide_sequence", "read_identifiers"])

	for referral in referral_list:

		if (referral!="MET"):

			if (os.stat("./Results/arriba/"+sampleId+"_fusion_report_"+referral+"_arriba.txt").st_size!=0):
      	
				arriba_report=pandas.read_csv("Results/arriba/"+sampleId+"_fusion_report_"+referral+"_arriba.txt", sep="\t")
				print(len(arriba_report))
				if len(arriba_report!=0):
					del arriba_report['Unnamed: 0']
					

				else:
					dict_arriba=({"gene1": referral + "-no fusions found", "gene2": "", "strand1(gene/fusion)": "", "strand2(gene/fusion)": "", "breakpoint1": "", "breakpoint2": "", "site1": "", "site2": "", "type": "", "direction1": "", "direction2": "", "split_reads1": "", "split_reads2": "", "discordant_mates": "", "coverage1": "", "coverage2": "", "confidence": "", "closest_genomic_breakpoint1": "", "closest_genomic_breakpoint2": "", "filters": "", "fusion_transcript": "", "reading_frame" : "", "peptide_sequence": "", "read_identifiers": ""}) 
					arriba_report= pandas.DataFrame(dict_arriba, columns=["gene1", "gene2", "strand1(gene/fusion)", "strand2(gene/fusion)", "breakpoint1", "breakpoint2", "site1", "site2", "type", "direction1", "direction2", "split_reads1", "split_reads2", "discordant_mates", "coverage1", "coverage2", "confidence", "closest_genomic_breakpoint1", "closest_genomic_breakpoint2", "filters", "fusion_transcript", "reading_frame", "peptide_sequence", "read_identifiers"], index=[0]) 	

				arriba_report_final=arriba_report_final.append(arriba_report, sort=False)

	for row in dataframe_to_rows(arriba_report_final, header=True, index=False):
		ws1.append(row)
		
		

def get_ntc_total_coverage():
	if (os.stat("../"+ntc+"/"+ntc+"_coverage.totalCoverage").st_size!=0):
		ntc_total_coverage=pandas.read_csv("../"+ntc+"/"+ntc+"_coverage.totalCoverage", sep="\t")
		ntc_total_average_depth=ntc_total_coverage["AVG_DEPTH"]

	if (os.stat("../"+ntc+"/"+ntc+"_rmdup_coverage.totalCoverage").st_size!=0):
		ntc_total_coverage_rmdup=pandas.read_csv("../"+ntc+"/"+ntc+"_rmdup_coverage.totalCoverage", sep="\t")
		ntc_total_average_depth_rmdup=ntc_total_coverage_rmdup["AVG_DEPTH"]

	return(ntc_total_average_depth, ntc_total_average_depth_rmdup)





def get_total_coverage(ntc_total_average_depth, ntc_total_average_depth_rmdup):
	ws2["A2"]="with duplicates"
	if (os.stat(sampleId+"_coverage.totalCoverage").st_size!=0):
		total_coverage=pandas.read_csv(sampleId+"_coverage.totalCoverage", sep="\t")
		total_coverage["NTC_AVG_DEPTH"]=ntc_total_average_depth
		total_coverage["%NTC contamination"]=total_coverage["NTC_AVG_DEPTH"]/total_coverage["AVG_DEPTH"]
		for row in dataframe_to_rows(total_coverage, header=True, index=False):
			ws2.append(row)
	if (os.stat(sampleId+"_rmdup_coverage.totalCoverage").st_size!=0):
		ws2["A32"]="without duplicates"
		total_coverage_rmdup=pandas.read_csv(sampleId+"_rmdup_coverage.totalCoverage", sep="\t")
		total_coverage_rmdup["NTC_AVG_DEPTH"]=ntc_total_average_depth_rmdup
		total_coverage_rmdup["%NTC contamination"]=total_coverage_rmdup["NTC_AVG_DEPTH"]/total_coverage_rmdup["AVG_DEPTH"]
		for row in dataframe_to_rows(total_coverage_rmdup, header=True, index=False):
			ws2.append(row)
	ws2.column_dimensions['A'].width=50



def ntc_get_coverage():

	if (os.stat("../"+ntc+"/"+ntc+"_coverage.coverage").st_size!=0):
		ntc_coverage=pandas.read_csv("../"+ntc+"/"+ntc+"_coverage.coverage", sep="\t")
		ntc_average_depth=ntc_coverage["AVG_DEPTH"]

	if (os.stat("../"+ntc+"/"+ntc+"_rmdup_coverage.coverage").st_size!=0):
		ntc_coverage_rmdup=pandas.read_csv("../"+ntc+"/"+ntc+"_rmdup_coverage.coverage", sep="\t")
		ntc_average_depth_rmdup=ntc_coverage_rmdup["AVG_DEPTH"]

	return(ntc_average_depth, ntc_average_depth_rmdup)


def get_coverage(referral, ntc_average_depth, ntc_average_depth_rmdup):
	if (os.stat(sampleId+"_coverage.coverage").st_size!=0):
		coverage=pandas.read_csv(sampleId+"_coverage.coverage", sep="\t")
		coverage["NTC_AVG_DEPTH"]=ntc_average_depth
		coverage["%NTC contamination"]=coverage["NTC_AVG_DEPTH"]/coverage["AVG_DEPTH"]
		for row in dataframe_to_rows(coverage, header=True, index=False):
			ws3.append(row)
	if (os.stat(sampleId+"_rmdup_coverage.coverage").st_size!=0):
		coverage_rmdup=pandas.read_csv(sampleId+"_rmdup_coverage.coverage", sep="\t")
		coverage_rmdup["NTC_AVG_DEPTH"]=ntc_average_depth_rmdup
		coverage_rmdup["%NTC contamination"]=coverage_rmdup["NTC_AVG_DEPTH"]/coverage_rmdup["AVG_DEPTH"]
		for row in dataframe_to_rows(coverage_rmdup, header=True, index=False):
			ws4.append(row)



	ws3.column_dimensions['B'].width=15
	ws4.column_dimensions['B'].width=15
	ws3.column_dimensions['C'].width=15
	ws4.column_dimensions['C'].width=15
	ws3.column_dimensions['D'].width=110
	ws4.column_dimensions['D'].width=110

	return(coverage)



def get_quality_metrics(referral, total_coverage):



	border_a=Border(left=Side(border_style=BORDER_MEDIUM), right=Side(border_style=BORDER_MEDIUM), top=Side(border_style=BORDER_MEDIUM), bottom=Side(border_style=BORDER_MEDIUM))
	position=['I7', 'I8', 'J6', 'J7', 'J8', 'K6', 'K7', 'K8', 'L6', 'L7', 'L8' ]
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
	ws5["C7"]="NTC check result"
	ws5["D7"]="NTC check result"
	ws5["E7"]="NTC checker 1"
	ws5["F7"]="NTC checker 2"
	ws5["G7"]="Analysed by:"
	ws5["H7"]="Checked by:"

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


	ws5["A42"]="Gene fusions quality metrics"


	ws5["A44"]="Post PCR1 Qubit"
	ws5["A45"]="Total mapped reads-duplicates removed"



	ws5["B43"]="Value"
	ws5["C43"]="Conclusion checker 1"
	ws5["D43"]="Conclusion checker 2"


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
	ws5["D59"]="No.unique reads"
	ws5["E59"]="Gene_exon"
	ws5["F59"]="Conclusion 1"
	ws5["G59"]="Conclusion 2"


        
	MET_exon_13=total_coverage[total_coverage["START"]==116411698]
	del [MET_exon_13["META"], MET_exon_13["PERC_COVERAGE@100"],MET_exon_13["NTC_AVG_DEPTH"],MET_exon_13["%NTC contamination"]]
	MET_exon_15=total_coverage[total_coverage["END"]==116414944]
	del [MET_exon_15["META"], MET_exon_15["PERC_COVERAGE@100"],MET_exon_15["NTC_AVG_DEPTH"],MET_exon_15["%NTC contamination"]]
	EGFR_exon_1=total_coverage[total_coverage["START"]==55087048]
	del [EGFR_exon_1["META"], EGFR_exon_1["PERC_COVERAGE@100"],EGFR_exon_1["NTC_AVG_DEPTH"],EGFR_exon_1["%NTC contamination"]]
	EGFR_exon_8=total_coverage[total_coverage["END"]==55223532]
	del [EGFR_exon_8["META"], EGFR_exon_8["PERC_COVERAGE@100"], EGFR_exon_8["NTC_AVG_DEPTH"], EGFR_exon_8["%NTC contamination"]]

	for row in dataframe_to_rows(MET_exon_13, header=False, index=False):
		ws5.append(row)
	for row in dataframe_to_rows(MET_exon_15, header=False, index=False):
		ws5.append(row)
	for row in dataframe_to_rows(EGFR_exon_1, header=False, index=False):
		ws5.append(row)
	for row in dataframe_to_rows(EGFR_exon_8, header=False, index=False):
		ws5.append(row)

	ws5["E60"]="MET_Exon13_last10bases"
	ws5["E61"]="MET_Exon15_first10bases"
	ws5["E62"]="EGFR_Exon1_last10bases"
	ws5["E63"]="EGFR_Exon8_first10bases"


	



	border_a=Border(left=Side(border_style=BORDER_MEDIUM), right=Side(border_style=BORDER_MEDIUM), top=Side(border_style=BORDER_MEDIUM), bottom=Side(border_style=BORDER_MEDIUM))
	position=['A4','B4','C4','D4','E4','F4','G4', 'H4',
	'A5','B5','C5','D5','E5','F5', 'G5','H5',
        'A7','B7','C7','D7','E7','F7', 'G7', 'H7',
        'A8','B8','C8','D8','E8','F8', 'G8', 'H8',
        'A14','B14','C14','D14','E14','F14','G14', 'H14', 'I14', 'J14', 'K14', 'L14',
        'A28','B28','C28','D28','E28','F28','G28', 'H28', 'I28', 'J28', 'K28', 'L28', 'M28', 'N28', 'O28', 'P28',
        'B43', 'C43', 'D43',
        'A44', 'B44', 'C44', 'D44',
        'A45', 'B45', 'C45', 'D45',
        'A53','B53','C53','D53','E53','F53','G53', 'H53', 'I53', 'J53', 'K53',
        'A59','B59','C59','D59','E59','F59','G59']
	for cell in position:
		ws5[cell].border=border_a

	
	
	ws5.column_dimensions['A'].width=25
	ws5.column_dimensions['B'].width=25
	ws5.column_dimensions['C'].width=25
	ws5.column_dimensions['D'].width=25
	ws5.column_dimensions['E'].width=25
	ws5.column_dimensions['F'].width=35
	ws5.column_dimensions['G'].width=25
	ws5.column_dimensions['H'].width=25
	ws5.column_dimensions['I'].width=25
	ws5.column_dimensions['J'].width=25
	ws5.column_dimensions['K'].width=25
	ws5.column_dimensions['L'].width=25


	ws1.column_dimensions['A'].width=20
	ws1.column_dimensions['B'].width=20
	ws1.column_dimensions['C'].width=20
	ws1.column_dimensions['D'].width=20
	ws1.column_dimensions['E'].width=20
	ws1.column_dimensions['F'].width=20
	ws1.column_dimensions['G'].width=20
	ws1.column_dimensions['H'].width=20
	ws1.column_dimensions['I'].width=20
	ws1.column_dimensions['J'].width=20
	ws1.column_dimensions['K'].width=20
	ws1.column_dimensions['L'].width=20
	ws1.column_dimensions['M'].width=20
	ws1.column_dimensions['N'].width=20
	ws1.column_dimensions['O'].width=20
	ws1.column_dimensions['P'].width=20
	ws1.column_dimensions['Q'].width=20
	ws1.column_dimensions['R'].width=20
	ws1.column_dimensions['S'].width=20
	ws1.column_dimensions['T'].width=20
	ws1.column_dimensions['U'].width=20
	ws1.column_dimensions['V'].width=20
	ws1.column_dimensions['w'].width=20
	ws1.column_dimensions['X'].width=20
	ws1.column_dimensions['Y'].width=20



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
	position= ['J6', 'K6', 'L6']
	for cell in position:
		ws2[cell].fill=colour




	font_bold=Font(bold=True)
	position= ['A4', 'B4', 'C4', 'D4', 'E4', 'F4', 'G4', 'A39',
	'A10','B10','C10','D10','E10','F10','G10','H10','I10','J10','K10','L10','M10','N10','O10','P10','Q10','R10','S10','T10', 'U10',
	'A24','B24','C24','D24','E24','F24','G24','H24','I24','J24','K24','L24','M24','N24','O24','P24','Q24','R24','S24','T24','U24','V24','W24','X24','Y24', 'Z24', 'AA24']
	for cell in position:
		ws1[cell].font=font_bold


	ws2['G3']="Total mapped reads (including duplicates)"
	ws2['I7']="NTC check 1"
	ws2['I8']="NTC check 2"
	ws2['J6']="Result"
	ws2['K6']= "Comments"
	ws2['L6']="Initials and date"


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




	colour= PatternFill("solid", fgColor="DCDCDC")
	position= ['A14', 'B14', 'C14', 'D14', 'E14', 'F14', 'G14', 'H14', 'I14', 'J14', 
                   'A28', 'B28', 'C28', 'D28', 'E28', 'F28', 'G28', 'H28', 'I28', 'J28', 'K28', 'L28' , 'M28', 'N28',
                   'B43', 'A44', 'A45',
                   'A53', 'B53', 'C53', 'D53', 'E53', 'F53', 'G53', 'H53', 
                   'A59', 'B59', 'C59', 'D59', 'E59' ]
	for cell in position:
		ws5[cell].fill=colour




	colour= PatternFill("solid", fgColor="00CCFFFF")
	position= ['A4','B4','C4','D4','E4','F4', 'G4', 'H4', 'A7','B7','C7','D7']
	for cell in position:
		ws5[cell].fill=colour



	colour= PatternFill("solid", fgColor="00FFCC00")
	position= ['E7', 'F7', 'G7', 'H7', 'K14', 'L14', 'O28', 'P28', 'C43', 'D43', 'I53','J53', 'K53', 'F59' , 'G59']
	for cell in position:
		ws5[cell].fill=colour


	font_bold=Font(bold=True)
	position= ['A4','B4','C4','D4','E4','F4','G4','H4','A7','B7','C7','D7','E7','F7','G7', 'H7','A14','B14','C14','D14','E14','F14','G14','H14','I14', 'J14', 'K14','L14',
		'A28','B28','C28','D28','E28','F28','G28','H28','I28', 'J28', 'K28', 'L28', 'M28', 'N28', 'O28', 'P28',
		'B43', 'C43', 'D43', 'A44', 'A45',
		'A53','B53','C53','D53','E53','F53','G53','H53','I53', 'J53', 'K53',
		'A59','B59','C59','D59','E59','F59','G59',
		'A12', 'A13', 'A27', 'A42', 'A52', 'A58']
	for cell in position:
		ws5[cell].font=font_bold



	ws1['A5']=sampleId
	ws1['B5']='=Patient_demographics!C2'
	ws1['C5']=referral
	ws1['D5']=worksheetId
	ws1['E5']=seqId
	ws1['C1']="Analysis sheet- Pan RNA Cancer Fusion Panel"
	

	ws5['A5']=sampleId
	ws5['B5']= '=Patient_demographics!C2'
	ws5['C5']= '=Patient_demographics!H2'
	ws5['D5']=referral
	ws5['E5']= '=Patient_demographics!L2'
	ws5['F5']= '=Patient_demographics!J2'
	ws5['H5']= '=Patient_demographics!G2'


	ws5['A8']=worksheetId
	ws5['B8']=seqId
	ws5['C8']='=total_coverage!J7'
	ws5['D8']='=total_coverage!J8'
	ws5['E8']='=total_coverage!L7'
	ws5['F8']='=total_coverage!L8'
	ws5['B44']='=Patient_demographics!O2'

	ws7['B2']=sampleId
	ws7['E2']=referral
	ws7['M2']=worksheetId
	ws7['N2']=seqId

	dv = DataValidation(type="list", formula1='"PASS,FAIL"', allow_blank=True)
	ws5.add_data_validation(dv)
	dv.add(ws5['C44'])
	dv.add(ws5['D44'])
	dv.add(ws5['C45'])
	dv.add(ws5['D45'])
	dv.add(ws5['F60'])
	dv.add(ws5['G60'])
	dv.add(ws5['F61'])
	dv.add(ws5['G61'])
	dv.add(ws5['F62'])
	dv.add(ws5['G62'])
	dv.add(ws5['F63'])
	dv.add(ws5['G63'])



	ws7.column_dimensions['A'].width=30
	ws7.column_dimensions['B'].width=20
	ws7.column_dimensions['C'].width=30
	ws7.column_dimensions['N'].width=40


	ws1['S10']="Conclusion checker 1"
	ws1['T10']="Conclusion checker 2"
	ws1['U10']="Comments"
	ws1['Y24']="Conclusion checker 1"
	ws1['Z24']="Conclusion checker 2"
	ws1['AA24']="Comments"



	dv = DataValidation(type="list", formula1='"Genuine,Artefact,Fail_QC,no fusions"', allow_blank=True)
	ws1.add_data_validation(dv)
	dv.add(ws1['S11'])
	dv.add(ws1['S12'])
	dv.add(ws1['S13'])
	dv.add(ws1['S14'])
	dv.add(ws1['S15'])
	dv.add(ws1['S16'])
	dv.add(ws1['S17'])
	dv.add(ws1['S18'])
	dv.add(ws1['S19'])
	dv.add(ws1['S20'])
	dv.add(ws1['S21'])

	dv.add(ws1['T11'])
	dv.add(ws1['T12'])
	dv.add(ws1['T13'])
	dv.add(ws1['T14'])
	dv.add(ws1['T15'])
	dv.add(ws1['T16'])
	dv.add(ws1['T17'])
	dv.add(ws1['T18'])
	dv.add(ws1['T19'])
	dv.add(ws1['T20'])
	dv.add(ws1['T21'])

	dv.add(ws1['Y25'])
	dv.add(ws1['Y26'])
	dv.add(ws1['Y27'])
	dv.add(ws1['Y28'])
	dv.add(ws1['Y29'])
	dv.add(ws1['Y30'])
	dv.add(ws1['Y31'])
	dv.add(ws1['Y32'])
	dv.add(ws1['Y33'])
	dv.add(ws1['Y34'])
	dv.add(ws1['Y35'])

	dv.add(ws1['Z25'])
	dv.add(ws1['Z26'])
	dv.add(ws1['Z27'])
	dv.add(ws1['Z28'])
	dv.add(ws1['Z29'])
	dv.add(ws1['Z30'])
	dv.add(ws1['Z31'])
	dv.add(ws1['Z32'])
	dv.add(ws1['Z33'])
	dv.add(ws1['Z34'])
	dv.add(ws1['Z35'])




	ws5['A15']= '=IF(OR(Gene_fusion_report!T11="Genuine", Gene_fusion_report!T11="no_fusions"),Gene_fusion_report!A11, "")'
	ws5['B15']= '=IF(OR(Gene_fusion_report!T11="Genuine", Gene_fusion_report!T11="no_fusions"),Gene_fusion_report!B11, "")'
	ws5['C15']= '=IF(OR(Gene_fusion_report!T11="Genuine",Gene_fusion_report!T11="no_fusions"),Gene_fusion_report!C11, "")'
	ws5['D15']= '=IF(OR(Gene_fusion_report!T11="Genuine",Gene_fusion_report!T11="no_fusions"),Gene_fusion_report!D11, "")'
	ws5['E15']= '=IF(OR(Gene_fusion_report!T11="Genuine",Gene_fusion_report!T11="no_fusions"),Gene_fusion_report!E11, "")'
	ws5['F15']= '=IF(OR(Gene_fusion_report!T11="Genuine",Gene_fusion_report!T11="no_fusions"),Gene_fusion_report!F11, "")'
	ws5['G15']= '=IF(OR(Gene_fusion_report!T11="Genuine",Gene_fusion_report!T11="no_fusions"),Gene_fusion_report!G11, "")'
	ws5['H15']= '=IF(OR(Gene_fusion_report!T11="Genuine",Gene_fusion_report!T11="no_fusions"),Gene_fusion_report!H11, "")'
	ws5['I15']= '=IF(OR(Gene_fusion_report!T11="Genuine",Gene_fusion_report!T11="no_fusions"),Gene_fusion_report!I11, "")'
	ws5['J15']= '=IF(OR(Gene_fusion_report!T11="Genuine",Gene_fusion_report!T11="no_fusions"),Gene_fusion_report!R11, "")'
	ws5['K15']= '=IF(OR(Gene_fusion_report!T11="Genuine",Gene_fusion_report!T11="no_fusions"),Gene_fusion_report!S11, "")'
	ws5['L15']= '=IF(OR(Gene_fusion_report!T11="Genuine",Gene_fusion_report!T11="no_fusions"),Gene_fusion_report!T11, "")'
	ws5['M15']= '=IF(OR(Gene_fusion_report!T11="Genuine",Gene_fusion_report!T11="no_fusions"),Gene_fusion_report!U11, "")'

	ws5['A16']= '=IF(OR(Gene_fusion_report!T12="Genuine", Gene_fusion_report!T12="no_fusions"),Gene_fusion_report!A12, "")'
	ws5['B16']= '=IF(OR(Gene_fusion_report!T12="Genuine",Gene_fusion_report!T12="no_fusions"), Gene_fusion_report!B12, "")'
	ws5['C16']= '=IF(OR(Gene_fusion_report!T12="Genuine",Gene_fusion_report!T12="no_fusions"),Gene_fusion_report!C12, "")'
	ws5['D16']= '=IF(OR(Gene_fusion_report!T12="Genuine",Gene_fusion_report!T12="no_fusions"),Gene_fusion_report!D12, "")'
	ws5['E16']= '=IF(OR(Gene_fusion_report!T12="Genuine",Gene_fusion_report!T12="no_fusions"),Gene_fusion_report!E12, "")'
	ws5['F16']= '=IF(OR(Gene_fusion_report!T12="Genuine",Gene_fusion_report!T12="no_fusions"),Gene_fusion_report!F12, "")'
	ws5['G16']= '=IF(OR(Gene_fusion_report!T12="Genuine",Gene_fusion_report!T12="no_fusions"),Gene_fusion_report!G12, "")'
	ws5['H16']= '=IF(OR(Gene_fusion_report!T12="Genuine",Gene_fusion_report!T12="no_fusions"),Gene_fusion_report!H12, "")'
	ws5['I16']= '=IF(OR(Gene_fusion_report!T12="Genuine",Gene_fusion_report!T12="no_fusions"),Gene_fusion_report!I12, "")'
	ws5['J16']= '=IF(OR(Gene_fusion_report!T12="Genuine",Gene_fusion_report!T12="no_fusions"),Gene_fusion_report!R12, "")'
	ws5['K16']= '=IF(OR(Gene_fusion_report!T12="Genuine",Gene_fusion_report!T12="no_fusions"),Gene_fusion_report!S12, "")'
	ws5['L16']= '=IF(OR(Gene_fusion_report!T12="Genuine",Gene_fusion_report!T12="no_fusions"),Gene_fusion_report!T12, "")'
	ws5['M16']= '=IF(OR(Gene_fusion_report!T11="Genuine",Gene_fusion_report!T12="no_fusions"),Gene_fusion_report!U12, "")'

	ws5['A17']= '=IF(OR(Gene_fusion_report!T13="Genuine",Gene_fusion_report!T13="no_fusions"),Gene_fusion_report!A13, "")'
	ws5['B17']= '=IF(OR(Gene_fusion_report!T13="Genuine",Gene_fusion_report!T13="no_fusions"),Gene_fusion_report!B13, "")'
	ws5['C17']= '=IF(OR(Gene_fusion_report!T13="Genuine",Gene_fusion_report!T13="no_fusions"),Gene_fusion_report!C13, "")'
	ws5['D17']= '=IF(OR(Gene_fusion_report!T13="Genuine",Gene_fusion_report!T13="no_fusions"),Gene_fusion_report!D13, "")'
	ws5['E17']= '=IF(OR(Gene_fusion_report!T13="Genuine",Gene_fusion_report!T13="no_fusions"),Gene_fusion_report!E13, "")'
	ws5['F17']= '=IF(OR(Gene_fusion_report!T13="Genuine",Gene_fusion_report!T13="no_fusions"),Gene_fusion_report!F13, "")'
	ws5['G17']= '=IF(OR(Gene_fusion_report!T13="Genuine",Gene_fusion_report!T13="no_fusions"),Gene_fusion_report!G13, "")'
	ws5['H17']= '=IF(OR(Gene_fusion_report!T13="Genuine",Gene_fusion_report!T13="no_fusions"),Gene_fusion_report!H13, "")'
	ws5['I17']= '=IF(OR(Gene_fusion_report!T13="Genuine",Gene_fusion_report!T13="no_fusions"),Gene_fusion_report!I13, "")'
	ws5['J17']= '=IF(OR(Gene_fusion_report!T13="Genuine",Gene_fusion_report!T13="no_fusions"),Gene_fusion_report!R13, "")'
	ws5['K17']= '=IF(OR(Gene_fusion_report!T13="Genuine",Gene_fusion_report!T13="no_fusions"),Gene_fusion_report!S13, "")'
	ws5['L17']= '=IF(OR(Gene_fusion_report!T13="Genuine",Gene_fusion_report!T13="no_fusions"),Gene_fusion_report!T13, "")'
	ws5['M17']= '=IF(OR(Gene_fusion_report!T13="Genuine",Gene_fusion_report!T13="no_fusions"),Gene_fusion_report!U13, "")'

	ws5['A18']= '=IF(OR(Gene_fusion_report!T14="Genuine",Gene_fusion_report!T14="no_fusions"),Gene_fusion_report!A14, "")'
	ws5['B18']= '=IF(OR(Gene_fusion_report!T14="Genuine",Gene_fusion_report!T14="no_fusions"),Gene_fusion_report!B14, "")'
	ws5['C18']= '=IF(OR(Gene_fusion_report!T14="Genuine",Gene_fusion_report!T14="no_fusions"),Gene_fusion_report!C14, "")'
	ws5['D18']= '=IF(OR(Gene_fusion_report!T14="Genuine",Gene_fusion_report!T14="no_fusions"),Gene_fusion_report!D14, "")'
	ws5['E18']= '=IF(OR(Gene_fusion_report!T14="Genuine",Gene_fusion_report!T14="no_fusions"),Gene_fusion_report!E14, "")'
	ws5['F18']= '=IF(OR(Gene_fusion_report!T14="Genuine",Gene_fusion_report!T14="no_fusions"),Gene_fusion_report!F14, "")'
	ws5['G18']= '=IF(OR(Gene_fusion_report!T14="Genuine",Gene_fusion_report!T14="no_fusions"),Gene_fusion_report!G14, "")'
	ws5['H18']= '=IF(OR(Gene_fusion_report!T14="Genuine",Gene_fusion_report!T14="no_fusions"),Gene_fusion_report!H14, "")'
	ws5['I18']= '=IF(OR(Gene_fusion_report!T14="Genuine",Gene_fusion_report!T14="no_fusions"),Gene_fusion_report!I14, "")'
	ws5['J18']= '=IF(OR(Gene_fusion_report!T14="Genuine",Gene_fusion_report!T14="no_fusions"),Gene_fusion_report!R14, "")'
	ws5['K18']= '=IF(OR(Gene_fusion_report!T14="Genuine",Gene_fusion_report!T14="no_fusions"),Gene_fusion_report!S14, "")'
	ws5['L18']= '=IF(OR(Gene_fusion_report!T14="Genuine",Gene_fusion_report!T14="no_fusions"),Gene_fusion_report!T14, "")'
	ws5['M18']= '=IF(OR(Gene_fusion_report!T14="Genuine",Gene_fusion_report!T14="no_fusions"),Gene_fusion_report!U14, "")'

	ws5['A19']= '=IF(OR(Gene_fusion_report!T15="Genuine",Gene_fusion_report!T15="no_fusions"),Gene_fusion_report!A15, "")'
	ws5['B19']= '=IF(OR(Gene_fusion_report!T15="Genuine",Gene_fusion_report!T15="no_fusions"),Gene_fusion_report!B15, "")'
	ws5['C19']= '=IF(OR(Gene_fusion_report!T15="Genuine",Gene_fusion_report!T15="no_fusions"),Gene_fusion_report!C15, "")'
	ws5['D19']= '=IF(OR(Gene_fusion_report!T15="Genuine",Gene_fusion_report!T15="no_fusions"),Gene_fusion_report!D15, "")'
	ws5['E19']= '=IF(OR(Gene_fusion_report!T15="Genuine",Gene_fusion_report!T15="no_fusions"),Gene_fusion_report!E15, "")'
	ws5['F19']= '=IF(OR(Gene_fusion_report!T15="Genuine",Gene_fusion_report!T15="no_fusions"),Gene_fusion_report!F15, "")'
	ws5['G19']= '=IF(OR(Gene_fusion_report!T15="Genuine",Gene_fusion_report!T15="no_fusions"),Gene_fusion_report!G15, "")'
	ws5['H19']= '=IF(OR(Gene_fusion_report!T15="Genuine",Gene_fusion_report!T15="no_fusions"),Gene_fusion_report!H15, "")'
	ws5['I19']= '=IF(OR(Gene_fusion_report!T15="Genuine",Gene_fusion_report!T15="no_fusions"),Gene_fusion_report!I15, "")'
	ws5['J19']= '=IF(OR(Gene_fusion_report!T15="Genuine",Gene_fusion_report!T15="no_fusions"),Gene_fusion_report!R15, "")'
	ws5['K19']= '=IF(OR(Gene_fusion_report!T15="Genuine",Gene_fusion_report!T15="no_fusions"),Gene_fusion_report!S15, "")'
	ws5['L19']= '=IF(OR(Gene_fusion_report!T15="Genuine",Gene_fusion_report!T15="no_fusions"),Gene_fusion_report!T15, "")'
	ws5['M19']= '=IF(OR(Gene_fusion_report!T15="Genuine",Gene_fusion_report!T15="no_fusions"),Gene_fusion_report!U15, "")'

	ws5['A20']= '=IF(OR(Gene_fusion_report!T16="Genuine",Gene_fusion_report!T16="no_fusions"),Gene_fusion_report!A16, "")'
	ws5['B20']= '=IF(OR(Gene_fusion_report!T16="Genuine",Gene_fusion_report!T16="no_fusions"),Gene_fusion_report!B16, "")'
	ws5['C20']= '=IF(OR(Gene_fusion_report!T16="Genuine",Gene_fusion_report!T16="no_fusions"),Gene_fusion_report!C16, "")'
	ws5['D20']= '=IF(OR(Gene_fusion_report!T16="Genuine",Gene_fusion_report!T16="no_fusions"),Gene_fusion_report!D16, "")'
	ws5['E20']= '=IF(OR(Gene_fusion_report!T16="Genuine",Gene_fusion_report!T16="no_fusions"),Gene_fusion_report!E16, "")'
	ws5['F20']= '=IF(OR(Gene_fusion_report!T16="Genuine",Gene_fusion_report!T16="no_fusions"),Gene_fusion_report!F16, "")'
	ws5['G20']= '=IF(OR(Gene_fusion_report!T16="Genuine",Gene_fusion_report!T16="no_fusions"),Gene_fusion_report!G16, "")'
	ws5['H20']= '=IF(OR(Gene_fusion_report!T16="Genuine",Gene_fusion_report!T16="no_fusions"),Gene_fusion_report!H16, "")'
	ws5['I20']= '=IF(OR(Gene_fusion_report!T16="Genuine",Gene_fusion_report!T16="no_fusions"),Gene_fusion_report!I16, "")'
	ws5['J20']= '=IF(OR(Gene_fusion_report!T16="Genuine",Gene_fusion_report!T16="no_fusions"),Gene_fusion_report!R16, "")'
	ws5['K20']= '=IF(OR(Gene_fusion_report!T16="Genuine",Gene_fusion_report!T16="no_fusions"),Gene_fusion_report!S16, "")'
	ws5['L20']= '=IF(OR(Gene_fusion_report!T16="Genuine",Gene_fusion_report!T16="no_fusions"),Gene_fusion_report!T16, "")'
	ws5['M20']= '=IF(OR(Gene_fusion_report!T16="Genuine",Gene_fusion_report!T16="no_fusions"),Gene_fusion_report!U16, "")'


	ws5['A21']= '=IF(OR(Gene_fusion_report!T17="Genuine",Gene_fusion_report!T17="no_fusions"),Gene_fusion_report!A17, "")'
	ws5['B21']= '=IF(OR(Gene_fusion_report!T17="Genuine",Gene_fusion_report!T17="no_fusions"),Gene_fusion_report!B17, "")'
	ws5['C21']= '=IF(OR(Gene_fusion_report!T17="Genuine",Gene_fusion_report!T17="no_fusions"),Gene_fusion_report!C17, "")'
	ws5['D21']= '=IF(OR(Gene_fusion_report!T17="Genuine",Gene_fusion_report!T17="no_fusions"),Gene_fusion_report!D17, "")'
	ws5['E21']= '=IF(OR(Gene_fusion_report!T17="Genuine",Gene_fusion_report!T17="no_fusions"),Gene_fusion_report!E17, "")'
	ws5['F21']= '=IF(OR(Gene_fusion_report!T17="Genuine",Gene_fusion_report!T17="no_fusions"),Gene_fusion_report!F17, "")'
	ws5['G21']= '=IF(OR(Gene_fusion_report!T17="Genuine",Gene_fusion_report!T17="no_fusions"),Gene_fusion_report!G17, "")'
	ws5['H21']= '=IF(OR(Gene_fusion_report!T17="Genuine",Gene_fusion_report!T17="no_fusions"),Gene_fusion_report!H17, "")'
	ws5['I21']= '=IF(OR(Gene_fusion_report!T17="Genuine",Gene_fusion_report!T17="no_fusions"),Gene_fusion_report!I17, "")'
	ws5['J21']= '=IF(OR(Gene_fusion_report!T17="Genuine",Gene_fusion_report!T17="no_fusions"),Gene_fusion_report!R17, "")'
	ws5['K21']= '=IF(OR(Gene_fusion_report!T17="Genuine",Gene_fusion_report!T17="no_fusions"),Gene_fusion_report!S17, "")'
	ws5['L21']= '=IF(OR(Gene_fusion_report!T17="Genuine",Gene_fusion_report!T17="no_fusions"),Gene_fusion_report!T17, "")'
	ws5['M21']= '=IF(OR(Gene_fusion_report!T17="Genuine",Gene_fusion_report!T17="no_fusions"),Gene_fusion_report!U17, "")'


	ws5['A22']= '=IF(OR(Gene_fusion_report!T18="Genuine",Gene_fusion_report!T18="no_fusions"),Gene_fusion_report!A18, "")'
	ws5['B22']= '=IF(OR(Gene_fusion_report!T18="Genuine",Gene_fusion_report!T18="no_fusions"),Gene_fusion_report!B18, "")'
	ws5['C22']= '=IF(OR(Gene_fusion_report!T18="Genuine",Gene_fusion_report!T18="no_fusions"),Gene_fusion_report!C18, "")'
	ws5['D22']= '=IF(OR(Gene_fusion_report!T18="Genuine",Gene_fusion_report!T18="no_fusions"),Gene_fusion_report!D18, "")'
	ws5['E22']= '=IF(OR(Gene_fusion_report!T18="Genuine",Gene_fusion_report!T18="no_fusions"),Gene_fusion_report!E18, "")'
	ws5['F22']= '=IF(OR(Gene_fusion_report!T18="Genuine",Gene_fusion_report!T18="no_fusions"),Gene_fusion_report!F18, "")'
	ws5['G22']= '=IF(OR(Gene_fusion_report!T18="Genuine",Gene_fusion_report!T18="no_fusions"),Gene_fusion_report!G18, "")'
	ws5['H22']= '=IF(OR(Gene_fusion_report!T18="Genuine",Gene_fusion_report!T18="no_fusions"),Gene_fusion_report!H18, "")'
	ws5['I22']= '=IF(OR(Gene_fusion_report!T18="Genuine",Gene_fusion_report!T18="no_fusions"),Gene_fusion_report!I18, "")'
	ws5['J22']= '=IF(OR(Gene_fusion_report!T18="Genuine",Gene_fusion_report!T18="no_fusions"),Gene_fusion_report!R18, "")'
	ws5['K22']= '=IF(OR(Gene_fusion_report!T18="Genuine",Gene_fusion_report!T18="no_fusions"),Gene_fusion_report!S18, "")'
	ws5['L22']= '=IF(OR(Gene_fusion_report!T18="Genuine",Gene_fusion_report!T18="no_fusions"),Gene_fusion_report!T18, "")'
	ws5['M22']= '=IF(OR(Gene_fusion_report!T18="Genuine",Gene_fusion_report!T18="no_fusions"),Gene_fusion_report!U18, "")'


	ws5['A23']= '=IF(OR(Gene_fusion_report!T19="Genuine",Gene_fusion_report!T19="no_fusions"),Gene_fusion_report!A19, "")'
	ws5['B23']= '=IF(OR(Gene_fusion_report!T19="Genuine",Gene_fusion_report!T19="no_fusions"),Gene_fusion_report!B19, "")'
	ws5['C23']= '=IF(OR(Gene_fusion_report!T19="Genuine",Gene_fusion_report!T19="no_fusions"),Gene_fusion_report!C19, "")'
	ws5['D23']= '=IF(OR(Gene_fusion_report!T19="Genuine",Gene_fusion_report!T19="no_fusions"),Gene_fusion_report!D19, "")'
	ws5['E23']= '=IF(OR(Gene_fusion_report!T19="Genuine",Gene_fusion_report!T19="no_fusions"),Gene_fusion_report!E19, "")'
	ws5['F23']= '=IF(OR(Gene_fusion_report!T19="Genuine",Gene_fusion_report!T19="no_fusions"),Gene_fusion_report!F19, "")'
	ws5['G23']= '=IF(OR(Gene_fusion_report!T19="Genuine",Gene_fusion_report!T19="no_fusions"),Gene_fusion_report!G19, "")'
	ws5['H23']= '=IF(OR(Gene_fusion_report!T19="Genuine",Gene_fusion_report!T19="no_fusions"),Gene_fusion_report!H19, "")'
	ws5['I23']= '=IF(OR(Gene_fusion_report!T19="Genuine",Gene_fusion_report!T19="no_fusions"),Gene_fusion_report!I19, "")'
	ws5['J23']= '=IF(OR(Gene_fusion_report!T19="Genuine",Gene_fusion_report!T19="no_fusions"),Gene_fusion_report!R19, "")'
	ws5['K23']= '=IF(OR(Gene_fusion_report!T19="Genuine",Gene_fusion_report!T19="no_fusions"),Gene_fusion_report!S19, "")'
	ws5['L23']= '=IF(OR(Gene_fusion_report!T19="Genuine",Gene_fusion_report!T19="no_fusions"),Gene_fusion_report!T19, "")'
	ws5['M23']= '=IF(OR(Gene_fusion_report!T19="Genuine",Gene_fusion_report!T19="no_fusions"),Gene_fusion_report!U19, "")'


	ws5['A24']= '=IF(OR(Gene_fusion_report!T20="Genuine",Gene_fusion_report!T20="no_fusions"),Gene_fusion_report!A20, "")'
	ws5['B24']= '=IF(OR(Gene_fusion_report!T20="Genuine",Gene_fusion_report!T20="no_fusions"),Gene_fusion_report!B20, "")'
	ws5['C24']= '=IF(OR(Gene_fusion_report!T20="Genuine",Gene_fusion_report!T20="no_fusions"),Gene_fusion_report!C20, "")'
	ws5['D24']= '=IF(OR(Gene_fusion_report!T20="Genuine",Gene_fusion_report!T20="no_fusions"),Gene_fusion_report!D20, "")'
	ws5['E24']= '=IF(OR(Gene_fusion_report!T20="Genuine",Gene_fusion_report!T20="no_fusions"),Gene_fusion_report!E20, "")'
	ws5['F24']= '=IF(OR(Gene_fusion_report!T20="Genuine",Gene_fusion_report!T20="no_fusions"),Gene_fusion_report!F20, "")'
	ws5['G24']= '=IF(OR(Gene_fusion_report!T20="Genuine",Gene_fusion_report!T20="no_fusions"),Gene_fusion_report!G20, "")'
	ws5['H24']= '=IF(OR(Gene_fusion_report!T20="Genuine",Gene_fusion_report!T20="no_fusions"),Gene_fusion_report!H20, "")'
	ws5['I24']= '=IF(OR(Gene_fusion_report!T20="Genuine",Gene_fusion_report!T20="no_fusions"),Gene_fusion_report!I20, "")'
	ws5['J24']= '=IF(OR(Gene_fusion_report!T20="Genuine",Gene_fusion_report!T20="no_fusions"),Gene_fusion_report!R20, "")'
	ws5['K24']= '=IF(OR(Gene_fusion_report!T20="Genuine",Gene_fusion_report!T20="no_fusions"),Gene_fusion_report!S20, "")'
	ws5['L24']= '=IF(OR(Gene_fusion_report!T20="Genuine",Gene_fusion_report!T20="no_fusions"),Gene_fusion_report!T20, "")'
	ws5['M24']= '=IF(OR(Gene_fusion_report!T20="Genuine",Gene_fusion_report!T20="no_fusions"),Gene_fusion_report!U20, "")'


	ws5['A25']= '=IF(OR(Gene_fusion_report!T21="Genuine",Gene_fusion_report!T21="no_fusions"),Gene_fusion_report!A21, "")'
	ws5['B25']= '=IF(OR(Gene_fusion_report!T21="Genuine",Gene_fusion_report!T21="no_fusions"),Gene_fusion_report!B21, "")'
	ws5['C25']= '=IF(OR(Gene_fusion_report!T21="Genuine",Gene_fusion_report!T21="no_fusions"),Gene_fusion_report!C21, "")'
	ws5['D25']= '=IF(OR(Gene_fusion_report!T21="Genuine",Gene_fusion_report!T21="no_fusions"),Gene_fusion_report!D21, "")'
	ws5['E25']= '=IF(OR(Gene_fusion_report!T21="Genuine",Gene_fusion_report!T21="no_fusions"),Gene_fusion_report!E21, "")'
	ws5['F25']= '=IF(OR(Gene_fusion_report!T21="Genuine",Gene_fusion_report!T21="no_fusions"),Gene_fusion_report!F21, "")'
	ws5['G25']= '=IF(OR(Gene_fusion_report!T21="Genuine",Gene_fusion_report!T21="no_fusions"),Gene_fusion_report!G21, "")'
	ws5['H25']= '=IF(OR(Gene_fusion_report!T21="Genuine",Gene_fusion_report!T21="no_fusions"),Gene_fusion_report!H21, "")'
	ws5['I25']= '=IF(OR(Gene_fusion_report!T21="Genuine",Gene_fusion_report!T21="no_fusions"),Gene_fusion_report!I21, "")'
	ws5['J25']= '=IF(OR(Gene_fusion_report!T21="Genuine",Gene_fusion_report!T21="no_fusions"),Gene_fusion_report!R21, "")'
	ws5['K25']= '=IF(OR(Gene_fusion_report!T21="Genuine",Gene_fusion_report!T21="no_fusions"),Gene_fusion_report!S21, "")'
	ws5['L25']= '=IF(OR(Gene_fusion_report!T21="Genuine",Gene_fusion_report!T21="no_fusions"),Gene_fusion_report!T21, "")'
	ws5['M25']= '=IF(OR(Gene_fusion_report!T21="Genuine",Gene_fusion_report!T21="no_fusions"),Gene_fusion_report!U21, "")'


	ws5['A26']= '=IF(OR(Gene_fusion_report!T22="Genuine",Gene_fusion_report!T22="no_fusions"),Gene_fusion_report!A22, "")'
	ws5['B26']= '=IF(OR(Gene_fusion_report!T22="Genuine",Gene_fusion_report!T22="no_fusions"),Gene_fusion_report!B22, "")'
	ws5['C26']= '=IF(OR(Gene_fusion_report!T22="Genuine",Gene_fusion_report!T22="no_fusions"),Gene_fusion_report!C22, "")'
	ws5['D26']= '=IF(OR(Gene_fusion_report!T22="Genuine",Gene_fusion_report!T22="no_fusions"),Gene_fusion_report!D22, "")'
	ws5['E26']= '=IF(OR(Gene_fusion_report!T22="Genuine",Gene_fusion_report!T22="no_fusions"),Gene_fusion_report!E22, "")'
	ws5['F26']= '=IF(OR(Gene_fusion_report!T22="Genuine",Gene_fusion_report!T22="no_fusions"),Gene_fusion_report!F22, "")'
	ws5['G26']= '=IF(OR(Gene_fusion_report!T22="Genuine",Gene_fusion_report!T22="no_fusions"),Gene_fusion_report!G22, "")'
	ws5['H26']= '=IF(OR(Gene_fusion_report!T22="Genuine",Gene_fusion_report!T22="no_fusions"),Gene_fusion_report!H22, "")'
	ws5['I26']= '=IF(OR(Gene_fusion_report!T22="Genuine",Gene_fusion_report!T22="no_fusions"),Gene_fusion_report!I22, "")'
	ws5['J26']= '=IF(OR(Gene_fusion_report!T22="Genuine",Gene_fusion_report!T22="no_fusions"),Gene_fusion_report!R22, "")'
	ws5['K26']= '=IF(OR(Gene_fusion_report!T22="Genuine",Gene_fusion_report!T22="no_fusions"),Gene_fusion_report!S22, "")'
	ws5['L26']= '=IF(OR(Gene_fusion_report!T22="Genuine",Gene_fusion_report!T22="no_fusions"),Gene_fusion_report!T22, "")'
	ws5['M26']= '=IF(OR(Gene_fusion_report!T22="Genuine",Gene_fusion_report!T22="no_fusions"),Gene_fusion_report!U22, "")'





#Transfer arriba genuine results to summary tab


	ws5['A29']= '=IF(OR(Gene_fusion_report!Z25="Genuine",Gene_fusion_report!Z25="no_fusions"),Gene_fusion_report!A25, "")'
	ws5['B29']= '=IF(OR(Gene_fusion_report!Z25="Genuine",Gene_fusion_report!Z25="no_fusions"),Gene_fusion_report!B25, "")'
	ws5['C29']= '=IF(OR(Gene_fusion_report!Z25="Genuine",Gene_fusion_report!Z25="no_fusions"),Gene_fusion_report!C25, "")'
	ws5['D29']= '=IF(OR(Gene_fusion_report!Z25="Genuine",Gene_fusion_report!Z25="no_fusions"),Gene_fusion_report!D25, "")'
	ws5['E29']= '=IF(OR(Gene_fusion_report!Z25="Genuine",Gene_fusion_report!Z25="no_fusions"),Gene_fusion_report!E25, "")'
	ws5['F29']= '=IF(OR(Gene_fusion_report!Z25="Genuine",Gene_fusion_report!Z25="no_fusions"),Gene_fusion_report!F25, "")'
	ws5['G29']= '=IF(OR(Gene_fusion_report!Z25="Genuine",Gene_fusion_report!Z25="no_fusions"),Gene_fusion_report!G25, "")'
	ws5['H29']= '=IF(OR(Gene_fusion_report!Z25="Genuine",Gene_fusion_report!Z25="no_fusions"),Gene_fusion_report!H25, "")'
	ws5['I29']= '=IF(OR(Gene_fusion_report!Z25="Genuine",Gene_fusion_report!Z25="no_fusions"),Gene_fusion_report!I25, "")'
	ws5['J29']= '=IF(OR(Gene_fusion_report!Z25="Genuine",Gene_fusion_report!Z25="no_fusions"),Gene_fusion_report!L25, "")'
	ws5['K29']= '=IF(OR(Gene_fusion_report!Z25="Genuine",Gene_fusion_report!Z25="no_fusions"),Gene_fusion_report!M25, "")'
	ws5['L29']= '=IF(OR(Gene_fusion_report!Z25="Genuine",Gene_fusion_report!Z25="no_fusions"),Gene_fusion_report!N25, "")'
	ws5['M29']= '=IF(OR(Gene_fusion_report!Z25="Genuine",Gene_fusion_report!Z25="no_fusions"),Gene_fusion_report!Q25, "")'
	ws5['N29']= '=IF(OR(Gene_fusion_report!Z25="Genuine",Gene_fusion_report!Z25="no_fusions"),Gene_fusion_report!T25, "")'
	ws5['O29']= '=IF(OR(Gene_fusion_report!Z25="Genuine",Gene_fusion_report!Z25="no_fusions"),Gene_fusion_report!Y25, "")'
	ws5['P29']= '=IF(OR(Gene_fusion_report!Z25="Genuine",Gene_fusion_report!Z25="no_fusions"),Gene_fusion_report!Z25, "")'


	ws5['A30']= '=IF(OR(Gene_fusion_report!Z26="Genuine",Gene_fusion_report!Z26="no_fusions"),Gene_fusion_report!A26, "")'
	ws5['B30']= '=IF(OR(Gene_fusion_report!Z26="Genuine",Gene_fusion_report!Z26="no_fusions"),Gene_fusion_report!B26, "")'
	ws5['C30']= '=IF(OR(Gene_fusion_report!Z26="Genuine",Gene_fusion_report!Z26="no_fusions"),Gene_fusion_report!C26, "")'
	ws5['D30']= '=IF(OR(Gene_fusion_report!Z26="Genuine",Gene_fusion_report!Z26="no_fusions"),Gene_fusion_report!D26, "")'
	ws5['E30']= '=IF(OR(Gene_fusion_report!Z26="Genuine",Gene_fusion_report!Z26="no_fusions"),Gene_fusion_report!E26, "")'
	ws5['F30']= '=IF(OR(Gene_fusion_report!Z26="Genuine",Gene_fusion_report!Z26="no_fusions"),Gene_fusion_report!F26, "")'
	ws5['G30']= '=IF(OR(Gene_fusion_report!Z26="Genuine",Gene_fusion_report!Z26="no_fusions"),Gene_fusion_report!G26, "")'
	ws5['H30']= '=IF(OR(Gene_fusion_report!Z26="Genuine",Gene_fusion_report!Z26="no_fusions"),Gene_fusion_report!H26, "")'
	ws5['I30']= '=IF(OR(Gene_fusion_report!Z26="Genuine",Gene_fusion_report!Z26="no_fusions"),Gene_fusion_report!I26, "")'
	ws5['J30']= '=IF(OR(Gene_fusion_report!Z26="Genuine",Gene_fusion_report!Z26="no_fusions"),Gene_fusion_report!L26, "")'
	ws5['K30']= '=IF(OR(Gene_fusion_report!Z26="Genuine",Gene_fusion_report!Z26="no_fusions"),Gene_fusion_report!M26, "")'
	ws5['L30']= '=IF(OR(Gene_fusion_report!Z26="Genuine",Gene_fusion_report!Z26="no_fusions"),Gene_fusion_report!N26, "")'
	ws5['M30']= '=IF(OR(Gene_fusion_report!Z26="Genuine",Gene_fusion_report!Z26="no_fusions"),Gene_fusion_report!Q26, "")'
	ws5['N30']= '=IF(OR(Gene_fusion_report!Z26="Genuine",Gene_fusion_report!Z26="no_fusions"),Gene_fusion_report!T26, "")'
	ws5['O30']= '=IF(OR(Gene_fusion_report!Z26="Genuine",Gene_fusion_report!Z26="no_fusions"),Gene_fusion_report!Y26, "")'
	ws5['P30']= '=IF(OR(Gene_fusion_report!Z26="Genuine",Gene_fusion_report!Z26="no_fusions"),Gene_fusion_report!Z26, "")'



	ws5['A31']= '=IF(OR(Gene_fusion_report!Z27="Genuine",Gene_fusion_report!Z27="no_fusions"),Gene_fusion_report!A27, "")'
	ws5['B31']= '=IF(OR(Gene_fusion_report!Z27="Genuine",Gene_fusion_report!Z27="no_fusions"),Gene_fusion_report!B27, "")'
	ws5['C31']= '=IF(OR(Gene_fusion_report!Z27="Genuine",Gene_fusion_report!Z27="no_fusions"),Gene_fusion_report!C27, "")'
	ws5['D31']= '=IF(OR(Gene_fusion_report!Z27="Genuine",Gene_fusion_report!Z27="no_fusions"),Gene_fusion_report!D27, "")'
	ws5['E31']= '=IF(OR(Gene_fusion_report!Z27="Genuine",Gene_fusion_report!Z27="no_fusions"),Gene_fusion_report!E27, "")'
	ws5['F31']= '=IF(OR(Gene_fusion_report!Z27="Genuine",Gene_fusion_report!Z27="no_fusions"),Gene_fusion_report!F27, "")'
	ws5['G31']= '=IF(OR(Gene_fusion_report!Z27="Genuine",Gene_fusion_report!Z27="no_fusions"),Gene_fusion_report!G27, "")'
	ws5['H31']= '=IF(OR(Gene_fusion_report!Z27="Genuine",Gene_fusion_report!Z27="no_fusions"),Gene_fusion_report!H27, "")'
	ws5['I31']= '=IF(OR(Gene_fusion_report!Z27="Genuine",Gene_fusion_report!Z27="no_fusions"),Gene_fusion_report!I27, "")'
	ws5['J31']= '=IF(OR(Gene_fusion_report!Z27="Genuine",Gene_fusion_report!Z27="no_fusions"),Gene_fusion_report!L27, "")'
	ws5['K31']= '=IF(OR(Gene_fusion_report!Z27="Genuine",Gene_fusion_report!Z27="no_fusions"),Gene_fusion_report!M27, "")'
	ws5['L31']= '=IF(OR(Gene_fusion_report!Z27="Genuine",Gene_fusion_report!Z27="no_fusions"),Gene_fusion_report!N27, "")'
	ws5['M31']= '=IF(OR(Gene_fusion_report!Z27="Genuine",Gene_fusion_report!Z27="no_fusions"),Gene_fusion_report!Q27, "")'
	ws5['N31']= '=IF(OR(Gene_fusion_report!Z27="Genuine",Gene_fusion_report!Z27="no_fusions"),Gene_fusion_report!T27, "")'
	ws5['O31']= '=IF(OR(Gene_fusion_report!Z27="Genuine",Gene_fusion_report!Z27="no_fusions"),Gene_fusion_report!Y27, "")'
	ws5['P31']= '=IF(OR(Gene_fusion_report!Z27="Genuine",Gene_fusion_report!Z27="no_fusions"),Gene_fusion_report!Z27, "")'


	ws5['A32']= '=IF(OR(Gene_fusion_report!Z28="Genuine",Gene_fusion_report!Z28="no_fusions"),Gene_fusion_report!A28, "")'
	ws5['B32']= '=IF(OR(Gene_fusion_report!Z28="Genuine",Gene_fusion_report!Z28="no_fusions"),Gene_fusion_report!B28, "")'
	ws5['C32']= '=IF(OR(Gene_fusion_report!Z28="Genuine",Gene_fusion_report!Z28="no_fusions"),Gene_fusion_report!C28, "")'
	ws5['D32']= '=IF(OR(Gene_fusion_report!Z28="Genuine",Gene_fusion_report!Z28="no_fusions"),Gene_fusion_report!D28, "")'
	ws5['E32']= '=IF(OR(Gene_fusion_report!Z28="Genuine",Gene_fusion_report!Z28="no_fusions"),Gene_fusion_report!E28, "")'
	ws5['F32']= '=IF(OR(Gene_fusion_report!Z28="Genuine",Gene_fusion_report!Z28="no_fusions"),Gene_fusion_report!F28, "")'
	ws5['G32']= '=IF(OR(Gene_fusion_report!Z28="Genuine",Gene_fusion_report!Z28="no_fusions"),Gene_fusion_report!G28, "")'
	ws5['H32']= '=IF(OR(Gene_fusion_report!Z28="Genuine",Gene_fusion_report!Z28="no_fusions"),Gene_fusion_report!H28, "")'
	ws5['I32']= '=IF(OR(Gene_fusion_report!Z28="Genuine",Gene_fusion_report!Z28="no_fusions"),Gene_fusion_report!I28, "")'
	ws5['J32']= '=IF(OR(Gene_fusion_report!Z28="Genuine",Gene_fusion_report!Z28="no_fusions"),Gene_fusion_report!L28, "")'
	ws5['K32']= '=IF(OR(Gene_fusion_report!Z28="Genuine",Gene_fusion_report!Z28="no_fusions"),Gene_fusion_report!M28, "")'
	ws5['L32']= '=IF(OR(Gene_fusion_report!Z28="Genuine",Gene_fusion_report!Z28="no_fusions"),Gene_fusion_report!N28, "")'
	ws5['M32']= '=IF(OR(Gene_fusion_report!Z28="Genuine",Gene_fusion_report!Z28="no_fusions"),Gene_fusion_report!Q28, "")'
	ws5['N32']= '=IF(OR(Gene_fusion_report!Z28="Genuine",Gene_fusion_report!Z28="no_fusions"),Gene_fusion_report!T28, "")'
	ws5['O32']= '=IF(OR(Gene_fusion_report!Z28="Genuine",Gene_fusion_report!Z28="no_fusions"),Gene_fusion_report!Y28, "")'
	ws5['P32']= '=IF(OR(Gene_fusion_report!Z28="Genuine",Gene_fusion_report!Z28="no_fusions"),Gene_fusion_report!Z28, "")'



	ws5['A33']= '=IF(OR(Gene_fusion_report!Z29="Genuine",Gene_fusion_report!Z29="no_fusions"),Gene_fusion_report!A29, "")'
	ws5['B33']= '=IF(OR(Gene_fusion_report!Z29="Genuine",Gene_fusion_report!Z29="no_fusions"),Gene_fusion_report!B29, "")'
	ws5['C33']= '=IF(OR(Gene_fusion_report!Z29="Genuine",Gene_fusion_report!Z29="no_fusions"),Gene_fusion_report!C29, "")'
	ws5['D33']= '=IF(OR(Gene_fusion_report!Z29="Genuine",Gene_fusion_report!Z29="no_fusions"),Gene_fusion_report!D29, "")'
	ws5['E33']= '=IF(OR(Gene_fusion_report!Z29="Genuine",Gene_fusion_report!Z29="no_fusions"),Gene_fusion_report!E29, "")'
	ws5['F33']= '=IF(OR(Gene_fusion_report!Z29="Genuine",Gene_fusion_report!Z29="no_fusions"),Gene_fusion_report!F29, "")'
	ws5['G33']= '=IF(OR(Gene_fusion_report!Z29="Genuine",Gene_fusion_report!Z29="no_fusions"),Gene_fusion_report!G29, "")'
	ws5['H33']= '=IF(OR(Gene_fusion_report!Z29="Genuine",Gene_fusion_report!Z29="no_fusions"),Gene_fusion_report!H29, "")'
	ws5['I33']= '=IF(OR(Gene_fusion_report!Z29="Genuine",Gene_fusion_report!Z29="no_fusions"),Gene_fusion_report!I29, "")'
	ws5['J33']= '=IF(OR(Gene_fusion_report!Z29="Genuine",Gene_fusion_report!Z29="no_fusions"),Gene_fusion_report!L29, "")'
	ws5['K33']= '=IF(OR(Gene_fusion_report!Z29="Genuine",Gene_fusion_report!Z29="no_fusions"),Gene_fusion_report!M29, "")'
	ws5['L33']= '=IF(OR(Gene_fusion_report!Z29="Genuine",Gene_fusion_report!Z29="no_fusions"),Gene_fusion_report!N29, "")'
	ws5['M33']= '=IF(OR(Gene_fusion_report!Z29="Genuine",Gene_fusion_report!Z29="no_fusions"),Gene_fusion_report!Q29, "")'
	ws5['N33']= '=IF(OR(Gene_fusion_report!Z29="Genuine",Gene_fusion_report!Z29="no_fusions"),Gene_fusion_report!T29, "")'
	ws5['O33']= '=IF(OR(Gene_fusion_report!Z29="Genuine",Gene_fusion_report!Z29="no_fusions"),Gene_fusion_report!Y29, "")'
	ws5['P33']= '=IF(OR(Gene_fusion_report!Z29="Genuine",Gene_fusion_report!Z29="no_fusions"),Gene_fusion_report!Z29, "")'


	ws5['A34']= '=IF(OR(Gene_fusion_report!Z30="Genuine",Gene_fusion_report!Z30="no_fusions"),Gene_fusion_report!A30, "")'
	ws5['B34']= '=IF(OR(Gene_fusion_report!Z30="Genuine",Gene_fusion_report!Z30="no_fusions"),Gene_fusion_report!B30, "")'
	ws5['C34']= '=IF(OR(Gene_fusion_report!Z30="Genuine",Gene_fusion_report!Z30="no_fusions"),Gene_fusion_report!C30, "")'
	ws5['D34']= '=IF(OR(Gene_fusion_report!Z30="Genuine",Gene_fusion_report!Z30="no_fusions"),Gene_fusion_report!D30, "")'
	ws5['E34']= '=IF(OR(Gene_fusion_report!Z30="Genuine",Gene_fusion_report!Z30="no_fusions"),Gene_fusion_report!E30, "")'
	ws5['F34']= '=IF(OR(Gene_fusion_report!Z30="Genuine",Gene_fusion_report!Z30="no_fusions"),Gene_fusion_report!F30, "")'
	ws5['G34']= '=IF(OR(Gene_fusion_report!Z30="Genuine",Gene_fusion_report!Z30="no_fusions"),Gene_fusion_report!G30, "")'
	ws5['H34']= '=IF(OR(Gene_fusion_report!Z30="Genuine",Gene_fusion_report!Z30="no_fusions"),Gene_fusion_report!H30, "")'
	ws5['I34']= '=IF(OR(Gene_fusion_report!Z30="Genuine",Gene_fusion_report!Z30="no_fusions"),Gene_fusion_report!I30, "")'
	ws5['J34']= '=IF(OR(Gene_fusion_report!Z30="Genuine",Gene_fusion_report!Z30="no_fusions"),Gene_fusion_report!L30, "")'
	ws5['K34']= '=IF(OR(Gene_fusion_report!Z30="Genuine",Gene_fusion_report!Z30="no_fusions"),Gene_fusion_report!M30, "")'
	ws5['L34']= '=IF(OR(Gene_fusion_report!Z30="Genuine",Gene_fusion_report!Z30="no_fusions"),Gene_fusion_report!N30, "")'
	ws5['M34']= '=IF(OR(Gene_fusion_report!Z30="Genuine",Gene_fusion_report!Z30="no_fusions"),Gene_fusion_report!Q30, "")'
	ws5['N34']= '=IF(OR(Gene_fusion_report!Z30="Genuine",Gene_fusion_report!Z30="no_fusions"),Gene_fusion_report!T30, "")'
	ws5['O34']= '=IF(OR(Gene_fusion_report!Z30="Genuine",Gene_fusion_report!Z30="no_fusions"),Gene_fusion_report!Y30, "")'
	ws5['P34']= '=IF(OR(Gene_fusion_report!Z30="Genuine",Gene_fusion_report!Z30="no_fusions"),Gene_fusion_report!Z30, "")'


	ws5['A35']= '=IF(OR(Gene_fusion_report!Z31="Genuine",Gene_fusion_report!Z31="no_fusions"),Gene_fusion_report!A31, "")'
	ws5['B35']= '=IF(OR(Gene_fusion_report!Z31="Genuine",Gene_fusion_report!Z31="no_fusions"),Gene_fusion_report!B31, "")'
	ws5['C35']= '=IF(OR(Gene_fusion_report!Z31="Genuine",Gene_fusion_report!Z31="no_fusions"),Gene_fusion_report!C31, "")'
	ws5['D35']= '=IF(OR(Gene_fusion_report!Z31="Genuine",Gene_fusion_report!Z31="no_fusions"),Gene_fusion_report!D31, "")'
	ws5['E35']= '=IF(OR(Gene_fusion_report!Z31="Genuine",Gene_fusion_report!Z31="no_fusions"),Gene_fusion_report!E31, "")'
	ws5['F35']= '=IF(OR(Gene_fusion_report!Z31="Genuine",Gene_fusion_report!Z31="no_fusions"),Gene_fusion_report!F31, "")'
	ws5['G35']= '=IF(OR(Gene_fusion_report!Z31="Genuine",Gene_fusion_report!Z31="no_fusions"),Gene_fusion_report!G31, "")'
	ws5['H35']= '=IF(OR(Gene_fusion_report!Z31="Genuine",Gene_fusion_report!Z31="no_fusions"),Gene_fusion_report!H31, "")'
	ws5['I35']= '=IF(OR(Gene_fusion_report!Z31="Genuine",Gene_fusion_report!Z31="no_fusions"),Gene_fusion_report!I31, "")'
	ws5['J35']= '=IF(OR(Gene_fusion_report!Z31="Genuine",Gene_fusion_report!Z31="no_fusions"),Gene_fusion_report!L31, "")'
	ws5['K35']= '=IF(OR(Gene_fusion_report!Z31="Genuine",Gene_fusion_report!Z31="no_fusions"),Gene_fusion_report!M31, "")'
	ws5['L35']= '=IF(OR(Gene_fusion_report!Z31="Genuine",Gene_fusion_report!Z31="no_fusions"),Gene_fusion_report!N31, "")'
	ws5['M35']= '=IF(OR(Gene_fusion_report!Z31="Genuine",Gene_fusion_report!Z31="no_fusions"),Gene_fusion_report!Q31, "")'
	ws5['N35']= '=IF(OR(Gene_fusion_report!Z31="Genuine",Gene_fusion_report!Z31="no_fusions"),Gene_fusion_report!T31, "")'
	ws5['O35']= '=IF(OR(Gene_fusion_report!Z31="Genuine",Gene_fusion_report!Z31="no_fusions"),Gene_fusion_report!Y31, "")'
	ws5['P35']= '=IF(OR(Gene_fusion_report!Z31="Genuine",Gene_fusion_report!Z31="no_fusions"),Gene_fusion_report!Z31, "")'


	ws5['A36']= '=IF(OR(Gene_fusion_report!Z32="Genuine",Gene_fusion_report!Z32="no_fusions"),Gene_fusion_report!A32, "")'
	ws5['B36']= '=IF(OR(Gene_fusion_report!Z32="Genuine",Gene_fusion_report!Z32="no_fusions"),Gene_fusion_report!B32, "")'
	ws5['C36']= '=IF(OR(Gene_fusion_report!Z32="Genuine",Gene_fusion_report!Z32="no_fusions"),Gene_fusion_report!C32, "")'
	ws5['D36']= '=IF(OR(Gene_fusion_report!Z32="Genuine",Gene_fusion_report!Z32="no_fusions"),Gene_fusion_report!D32, "")'
	ws5['E36']= '=IF(OR(Gene_fusion_report!Z32="Genuine",Gene_fusion_report!Z32="no_fusions"),Gene_fusion_report!E32, "")'
	ws5['F36']= '=IF(OR(Gene_fusion_report!Z32="Genuine",Gene_fusion_report!Z32="no_fusions"),Gene_fusion_report!F32, "")'
	ws5['G36']= '=IF(OR(Gene_fusion_report!Z32="Genuine",Gene_fusion_report!Z32="no_fusions"),Gene_fusion_report!G32, "")'
	ws5['H36']= '=IF(OR(Gene_fusion_report!Z32="Genuine",Gene_fusion_report!Z32="no_fusions"),Gene_fusion_report!H32, "")'
	ws5['I36']= '=IF(OR(Gene_fusion_report!Z32="Genuine",Gene_fusion_report!Z32="no_fusions"),Gene_fusion_report!I32, "")'
	ws5['J36']= '=IF(OR(Gene_fusion_report!Z32="Genuine",Gene_fusion_report!Z32="no_fusions"),Gene_fusion_report!L32, "")'
	ws5['K36']= '=IF(OR(Gene_fusion_report!Z32="Genuine",Gene_fusion_report!Z32="no_fusions"),Gene_fusion_report!M32, "")'
	ws5['L36']= '=IF(OR(Gene_fusion_report!Z32="Genuine",Gene_fusion_report!Z32="no_fusions"),Gene_fusion_report!N32, "")'
	ws5['M36']= '=IF(OR(Gene_fusion_report!Z32="Genuine",Gene_fusion_report!Z32="no_fusions"),Gene_fusion_report!Q32, "")'
	ws5['N36']= '=IF(OR(Gene_fusion_report!Z32="Genuine",Gene_fusion_report!Z32="no_fusions"),Gene_fusion_report!T32, "")'
	ws5['O36']= '=IF(OR(Gene_fusion_report!Z32="Genuine",Gene_fusion_report!Z32="no_fusions"),Gene_fusion_report!Y32, "")'
	ws5['P36']= '=IF(OR(Gene_fusion_report!Z32="Genuine",Gene_fusion_report!Z32="no_fusions"),Gene_fusion_report!Z32, "")'


	ws5['A37']= '=IF(OR(Gene_fusion_report!Z33="Genuine",Gene_fusion_report!Z33="no_fusions"),Gene_fusion_report!A33, "")'
	ws5['B37']= '=IF(OR(Gene_fusion_report!Z33="Genuine",Gene_fusion_report!Z33="no_fusions"),Gene_fusion_report!B33, "")'
	ws5['C37']= '=IF(OR(Gene_fusion_report!Z33="Genuine",Gene_fusion_report!Z33="no_fusions"),Gene_fusion_report!C33, "")'
	ws5['D37']= '=IF(OR(Gene_fusion_report!Z33="Genuine",Gene_fusion_report!Z33="no_fusions"),Gene_fusion_report!D33, "")'
	ws5['E37']= '=IF(OR(Gene_fusion_report!Z33="Genuine",Gene_fusion_report!Z33="no_fusions"),Gene_fusion_report!E33, "")'
	ws5['F37']= '=IF(OR(Gene_fusion_report!Z33="Genuine",Gene_fusion_report!Z33="no_fusions"),Gene_fusion_report!F33, "")'
	ws5['G37']= '=IF(OR(Gene_fusion_report!Z33="Genuine",Gene_fusion_report!Z33="no_fusions"),Gene_fusion_report!G33, "")'
	ws5['H37']= '=IF(OR(Gene_fusion_report!Z33="Genuine",Gene_fusion_report!Z33="no_fusions"),Gene_fusion_report!H33, "")'
	ws5['I37']= '=IF(OR(Gene_fusion_report!Z33="Genuine",Gene_fusion_report!Z31="no_fusions"),Gene_fusion_report!I33, "")'
	ws5['J37']= '=IF(OR(Gene_fusion_report!Z33="Genuine",Gene_fusion_report!Z33="no_fusions"),Gene_fusion_report!L33, "")'
	ws5['K37']= '=IF(OR(Gene_fusion_report!Z33="Genuine",Gene_fusion_report!Z33="no_fusions"),Gene_fusion_report!M33, "")'
	ws5['L37']= '=IF(OR(Gene_fusion_report!Z33="Genuine",Gene_fusion_report!Z33="no_fusions"),Gene_fusion_report!N33, "")'
	ws5['M37']= '=IF(OR(Gene_fusion_report!Z33="Genuine",Gene_fusion_report!Z33="no_fusions"),Gene_fusion_report!Q33, "")'
	ws5['N37']= '=IF(OR(Gene_fusion_report!Z33="Genuine",Gene_fusion_report!Z33="no_fusions"),Gene_fusion_report!T33, "")'
	ws5['O37']= '=IF(OR(Gene_fusion_report!Z33="Genuine",Gene_fusion_report!Z33="no_fusions"),Gene_fusion_report!Y33, "")'
	ws5['P37']= '=IF(OR(Gene_fusion_report!Z33="Genuine",Gene_fusion_report!Z33="no_fusions"),Gene_fusion_report!Z33, "")'


	ws5['A38']= '=IF(OR(Gene_fusion_report!Z34="Genuine",Gene_fusion_report!Z34="no_fusions"),Gene_fusion_report!A34, "")'
	ws5['B38']= '=IF(OR(Gene_fusion_report!Z34="Genuine",Gene_fusion_report!Z34="no_fusions"),Gene_fusion_report!B34, "")'
	ws5['C38']= '=IF(OR(Gene_fusion_report!Z34="Genuine",Gene_fusion_report!Z34="no_fusions"),Gene_fusion_report!C34, "")'
	ws5['D38']= '=IF(OR(Gene_fusion_report!Z34="Genuine",Gene_fusion_report!Z34="no_fusions"),Gene_fusion_report!D34, "")'
	ws5['E38']= '=IF(OR(Gene_fusion_report!Z34="Genuine",Gene_fusion_report!Z34="no_fusions"),Gene_fusion_report!E34, "")'
	ws5['F38']= '=IF(OR(Gene_fusion_report!Z34="Genuine",Gene_fusion_report!Z34="no_fusions"),Gene_fusion_report!F34, "")'
	ws5['G38']= '=IF(OR(Gene_fusion_report!Z34="Genuine",Gene_fusion_report!Z34="no_fusions"),Gene_fusion_report!G34, "")'
	ws5['H38']= '=IF(OR(Gene_fusion_report!Z34="Genuine",Gene_fusion_report!Z34="no_fusions"),Gene_fusion_report!H34, "")'
	ws5['I38']= '=IF(OR(Gene_fusion_report!Z34="Genuine",Gene_fusion_report!Z34="no_fusions"),Gene_fusion_report!I34, "")'
	ws5['J38']= '=IF(OR(Gene_fusion_report!Z34="Genuine",Gene_fusion_report!Z34="no_fusions"),Gene_fusion_report!L34, "")'
	ws5['K38']= '=IF(OR(Gene_fusion_report!Z34="Genuine",Gene_fusion_report!Z34="no_fusions"),Gene_fusion_report!M34, "")'
	ws5['L38']= '=IF(OR(Gene_fusion_report!Z34="Genuine",Gene_fusion_report!Z34="no_fusions"),Gene_fusion_report!N34, "")'
	ws5['M38']= '=IF(OR(Gene_fusion_report!Z34="Genuine",Gene_fusion_report!Z34="no_fusions"),Gene_fusion_report!Q34, "")'
	ws5['N38']= '=IF(OR(Gene_fusion_report!Z34="Genuine",Gene_fusion_report!Z34="no_fusions"),Gene_fusion_report!T34, "")'
	ws5['O38']= '=IF(OR(Gene_fusion_report!Z34="Genuine",Gene_fusion_report!Z34="no_fusions"),Gene_fusion_report!Y34, "")'
	ws5['P38']= '=IF(OR(Gene_fusion_report!Z34="Genuine",Gene_fusion_report!Z34="no_fusions"),Gene_fusion_report!Z34, "")'


	ws5['A39']= '=IF(OR(Gene_fusion_report!Z35="Genuine",Gene_fusion_report!Z35="no_fusions"),Gene_fusion_report!A35, "")'
	ws5['B39']= '=IF(OR(Gene_fusion_report!Z35="Genuine",Gene_fusion_report!Z35="no_fusions"),Gene_fusion_report!B35, "")'
	ws5['C39']= '=IF(OR(Gene_fusion_report!Z35="Genuine",Gene_fusion_report!Z35="no_fusions"),Gene_fusion_report!C35, "")'
	ws5['D39']= '=IF(OR(Gene_fusion_report!Z35="Genuine",Gene_fusion_report!Z35="no_fusions"),Gene_fusion_report!D35, "")'
	ws5['E39']= '=IF(OR(Gene_fusion_report!Z35="Genuine",Gene_fusion_report!Z35="no_fusions"),Gene_fusion_report!E35, "")'
	ws5['F39']= '=IF(OR(Gene_fusion_report!Z35="Genuine",Gene_fusion_report!Z35="no_fusions"),Gene_fusion_report!F35, "")'
	ws5['G39']= '=IF(OR(Gene_fusion_report!Z35="Genuine",Gene_fusion_report!Z35="no_fusions"),Gene_fusion_report!G35, "")'
	ws5['H39']= '=IF(OR(Gene_fusion_report!Z35="Genuine",Gene_fusion_report!Z35="no_fusions"),Gene_fusion_report!H35, "")'
	ws5['I39']= '=IF(OR(Gene_fusion_report!Z35="Genuine",Gene_fusion_report!Z35="no_fusions"),Gene_fusion_report!I35, "")'
	ws5['J39']= '=IF(OR(Gene_fusion_report!Z35="Genuine",Gene_fusion_report!Z35="no_fusions"),Gene_fusion_report!L35, "")'
	ws5['K39']= '=IF(OR(Gene_fusion_report!Z35="Genuine",Gene_fusion_report!Z35="no_fusions"),Gene_fusion_report!M35, "")'
	ws5['L39']= '=IF(OR(Gene_fusion_report!Z35="Genuine",Gene_fusion_report!Z35="no_fusions"),Gene_fusion_report!N35, "")'
	ws5['M39']= '=IF(OR(Gene_fusion_report!Z35="Genuine",Gene_fusion_report!Z35="no_fusions"),Gene_fusion_report!Q35, "")'
	ws5['N39']= '=IF(OR(Gene_fusion_report!Z35="Genuine",Gene_fusion_report!Z35="no_fusions"),Gene_fusion_report!T35, "")'
	ws5['O39']= '=IF(OR(Gene_fusion_report!Z35="Genuine",Gene_fusion_report!Z35="no_fusions"),Gene_fusion_report!Y35, "")'
	ws5['P39']= '=IF(OR(Gene_fusion_report!Z35="Genuine",Gene_fusion_report!Z35="no_fusions"),Gene_fusion_report!Z35, "")'


	ws5['A40']= '=IF(OR(Gene_fusion_report!Z36="Genuine",Gene_fusion_report!Z36="no_fusions"),Gene_fusion_report!A36, "")'
	ws5['B40']= '=IF(OR(Gene_fusion_report!Z36="Genuine",Gene_fusion_report!Z36="no_fusions"),Gene_fusion_report!B36, "")'
	ws5['C40']= '=IF(OR(Gene_fusion_report!Z36="Genuine",Gene_fusion_report!Z36="no_fusions"),Gene_fusion_report!C36, "")'
	ws5['D40']= '=IF(OR(Gene_fusion_report!Z36="Genuine",Gene_fusion_report!Z36="no_fusions"),Gene_fusion_report!D36, "")'
	ws5['E40']= '=IF(OR(Gene_fusion_report!Z36="Genuine",Gene_fusion_report!Z36="no_fusions"),Gene_fusion_report!E36, "")'
	ws5['F40']= '=IF(OR(Gene_fusion_report!Z36="Genuine",Gene_fusion_report!Z36="no_fusions"),Gene_fusion_report!F36, "")'
	ws5['G40']= '=IF(OR(Gene_fusion_report!Z36="Genuine",Gene_fusion_report!Z36="no_fusions"),Gene_fusion_report!G36, "")'
	ws5['H40']= '=IF(OR(Gene_fusion_report!Z36="Genuine",Gene_fusion_report!Z36="no_fusions"),Gene_fusion_report!H36, "")'
	ws5['I40']= '=IF(OR(Gene_fusion_report!Z36="Genuine",Gene_fusion_report!Z36="no_fusions"),Gene_fusion_report!I36, "")'
	ws5['J40']= '=IF(OR(Gene_fusion_report!Z36="Genuine",Gene_fusion_report!Z36="no_fusions"),Gene_fusion_report!L36, "")'
	ws5['K40']= '=IF(OR(Gene_fusion_report!Z36="Genuine",Gene_fusion_report!Z36="no_fusions"),Gene_fusion_report!M36, "")'
	ws5['L40']= '=IF(OR(Gene_fusion_report!Z36="Genuine",Gene_fusion_report!Z36="no_fusions"),Gene_fusion_report!N36, "")'
	ws5['M40']= '=IF(OR(Gene_fusion_report!Z36="Genuine",Gene_fusion_report!Z36="no_fusions"),Gene_fusion_report!Q36, "")'
	ws5['N40']= '=IF(OR(Gene_fusion_report!Z36="Genuine",Gene_fusion_report!Z36="no_fusions"),Gene_fusion_report!T36, "")'
	ws5['O40']= '=IF(OR(Gene_fusion_report!Z36="Genuine",Gene_fusion_report!Z36="no_fusions"),Gene_fusion_report!Y36, "")'
	ws5['P40']= '=IF(OR(Gene_fusion_report!Z36="Genuine",Gene_fusion_report!Z36="no_fusions"),Gene_fusion_report!Z36, "")'







def get_alignment_metrics():
	with open ("../"+sampleId+"_AlignmentSummaryMetrics.txt") as file:
		for line in file:
			if line.startswith("CATEGORY"):
				headers=line.split('\t')
			if line.startswith("PAIR"):
				pair_list=line.split('\t')


	alignment_metrics=pandas.DataFrame([pair_list], columns=headers)
	total_reads=alignment_metrics[['PF_READS']]
	total_reads_aligned=alignment_metrics[['PF_HQ_ALIGNED_READS']]
	total_reads_value=total_reads.iloc[0,0]
	#ws5["G14"]=total_reads_value
	aligned_reads_value=total_reads_aligned.iloc[0,0]
	#ws5["G15"]=aligned_reads_value
	


def get_alignment_metrics_rmdup():
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
	ws5["B45"]=aligned_reads_value_rmdup
	



def get_met_exon_skipping():




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
	for row in dataframe_to_rows(rmats_dataframe, header=True, index=False):
		ws8.append(row)





	ws8["N9"]="Conclusion checker 1"
	ws8["O9"]="Conclusion checker 2"
	ws8["P9"]="Comments"



	dv = DataValidation(type="list", formula1='"Genuine,Artefact,Fail_QC"', allow_blank=True)
	ws8.add_data_validation(dv)
	dv.add(ws8['N10'])
	dv.add(ws8['N11'])
	dv.add(ws8['N12'])


	dv.add(ws8['O10'])
	dv.add(ws8['O11'])
	dv.add(ws8['O12'])


	ws5['A54']= '=IF(OR(RMATS!O10="Genuine",RMATS!O10="no_fusions"),RMATS!B10, "")'
	ws5['B54']= '=IF(OR(RMATS!O10="Genuine",RMATS!O10="no_fusions"),RMATS!C10, "")'
	ws5['C54']= '=IF(OR(RMATS!O10="Genuine",RMATS!O10="no_fusions"),RMATS!D10, "")'
	ws5['D54']= '=IF(OR(RMATS!O10="Genuine",RMATS!O10="no_fusions"),RMATS!E10, "")'
	ws5['E54']= '=IF(OR(RMATS!O10="Genuine",RMATS!O10="no_fusions"),RMATS!F10, "")'
	ws5['F54']= '=IF(OR(RMATS!O10="Genuine",RMATS!O10="no_fusions"),RMATS!G10, "")'
	ws5['G54']= '=IF(OR(RMATS!O10="Genuine",RMATS!O10="no_fusions"),RMATS!J10, "")'
	ws5['H54']= '=IF(OR(RMATS!O10="Genuine",RMATS!O10="no_fusions"),RMATS!M10, "")'
	ws5['I54']= '=IF(OR(RMATS!O10="Genuine",RMATS!O10="no_fusions"),RMATS!N10, "")'
	ws5['J54']= '=IF(OR(RMATS!O10="Genuine",RMATS!O10="no_fusions"),RMATS!O10, "")'
	ws5['K54']= '=IF(OR(RMATS!O10="Genuine",RMATS!O10="no_fusions"),RMATS!P10, "")'


	ws5['A55']= '=IF(OR(RMATS!O11="Genuine",RMATS!O11="no_fusions"),RMATS!B11, "")'
	ws5['B55']= '=IF(OR(RMATS!O11="Genuine",RMATS!O11="no_fusions"),RMATS!C11, "")'
	ws5['C55']= '=IF(OR(RMATS!O11="Genuine",RMATS!O11="no_fusions"),RMATS!D11, "")'
	ws5['D55']= '=IF(OR(RMATS!O11="Genuine",RMATS!O11="no_fusions"),RMATS!E11, "")'
	ws5['E55']= '=IF(OR(RMATS!O11="Genuine",RMATS!O11="no_fusions"),RMATS!F11, "")'
	ws5['F55']= '=IF(OR(RMATS!O11="Genuine",RMATS!O11="no_fusions"),RMATS!G11, "")'
	ws5['G55']= '=IF(OR(RMATS!O11="Genuine",RMATS!O11="no_fusions"),RMATS!J11, "")'
	ws5['H55']= '=IF(OR(RMATS!O11="Genuine",RMATS!O11="no_fusions"),RMATS!M11, "")'
	ws5['I55']= '=IF(OR(RMATS!O11="Genuine",RMATS!O11="no_fusions"),RMATS!N11, "")'
	ws5['J55']= '=IF(OR(RMATS!O11="Genuine",RMATS!O11="no_fusions"),RMATS!O11, "")'
	ws5['K55']= '=IF(OR(RMATS!O11="Genuine",RMATS!O11="no_fusions"),RMATS!P11, "")'

	ws5['A56']= '=IF(OR(RMATS!O12="Genuine",RMATS!O12="no_fusions"),RMATS!B12, "")'
	ws5['B56']= '=IF(OR(RMATS!O12="Genuine",RMATS!O12="no_fusions"),RMATS!C12, "")'
	ws5['C56']= '=IF(OR(RMATS!O12="Genuine",RMATS!O12="no_fusions"),RMATS!D12, "")'
	ws5['D56']= '=IF(OR(RMATS!O12="Genuine",RMATS!O12="no_fusions"),RMATS!E12, "")'
	ws5['E56']= '=IF(OR(RMATS!O12="Genuine",RMATS!O12="no_fusions"),RMATS!F12, "")'
	ws5['F56']= '=IF(OR(RMATS!O12="Genuine",RMATS!O12="no_fusions"),RMATS!G12, "")'
	ws5['G56']= '=IF(OR(RMATS!O12="Genuine",RMATS!O12="no_fusions"),RMATS!J12, "")'
	ws5['H56']= '=IF(OR(RMATS!O12="Genuine",RMATS!O12="no_fusions"),RMATS!M12, "")'
	ws5['I56']= '=IF(OR(RMATS!O12="Genuine",RMATS!O12="no_fusions"),RMATS!N12, "")'
	ws5['J56']= '=IF(OR(RMATS!O12="Genuine",RMATS!O12="no_fusions"),RMATS!O12, "")'
	ws5['K56']= '=IF(OR(RMATS!O12="Genuine",RMATS!O12="no_fusions"),RMATS!P12, "")'
	


	ws8.column_dimensions['A'].width=30
	ws8.column_dimensions['B'].width=15
	ws8.column_dimensions['C'].width=15
	ws8.column_dimensions['D'].width=20
	ws8.column_dimensions['E'].width=20
	ws8.column_dimensions['F'].width=20
	ws8.column_dimensions['G'].width=20
	ws8.column_dimensions['H'].width=70
	ws8.column_dimensions['I'].width=40
	ws8.column_dimensions['J'].width=20
	ws8.column_dimensions['K'].width=20
	ws8.column_dimensions['L'].width=40
	ws8.column_dimensions['M'].width=20
	ws8.column_dimensions['N'].width=20
	ws8.column_dimensions['O'].width=20
	ws8.column_dimensions['P'].width=20
	ws8.column_dimensions['Q'].width=20


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


			


if __name__ == "__main__":

	seqId=sys.argv[1]
	sampleId=sys.argv[2]
	referral=sys.argv[3]
	ntc=sys.argv[4]
	worksheetId=sys.argv[5]


	referral_list=get_referral_list(referral)
	ntc_star_fusion_report, ntc_arriba_report=get_NTC_fusion_report(ntc)
	get_star_fusion_report(referral_list, ntc_star_fusion_report)
	get_arriba_fusion_report(referral_list)
	ntc_total_average_depth, ntc_total_average_depth_rmdup=get_ntc_total_coverage()
	get_total_coverage(ntc_total_average_depth, ntc_total_average_depth_rmdup)
	ntc_average_depth, ntc_average_depth_rmdup=ntc_get_coverage()
	coverage=get_coverage(referral, ntc_average_depth, ntc_average_depth_rmdup)
	get_quality_metrics(referral, coverage)
	get_alignment_metrics()
	get_alignment_metrics_rmdup()
	if (referral=="referral1" or referral=="referral6" or referral=="referral7" or referral=="referral8" or referral=="referral11"):
		get_met_exon_skipping()
	wb.save(sampleId+"_"+referral +"_updated.xlsx")
