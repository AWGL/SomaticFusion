import pandas
from openpyxl import Workbook
import os
import numpy
from openpyxl.utils.dataframe import dataframe_to_rows
import sys
from openpyxl.styles.borders import Border, Side, BORDER_MEDIUM, BORDER_THIN, BORDER_THICK
from openpyxl.styles import Font
from openpyxl.styles import PatternFill


wb=Workbook()
ws7=wb.create_sheet("Patient demographics")
ws5=wb.create_sheet("Summary")
ws1= wb.create_sheet("Gene fusion report")
ws8=wb.create_sheet("RMATS")
ws6= wb.create_sheet("NTC fusion report")
ws2=wb.create_sheet("total_coverage")
ws3=wb.create_sheet("coverage_with_duplicates")
ws4=wb.create_sheet("coverage_without_duplicates")

ws1["A2"]="Lab number"
ws1["B2"]="Patient name"
ws1["C2"]="Reason for referral"
ws1["D2"]="NGS wks"
ws1["E2"]="NextSeq runId"
ws1["F2"]="Checker 1 initials and date"
ws1["G2"]="Checker 2 initials and date"


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
ws7['K1']='Post extraction qubit RNA conc'
ws7['L1']='Pre-processing (dil/zymo)'
ws7['M1']='Qubit RNA conc after pre-processing'
ws7['N1']='NGS Worksheet'
ws7['O1']='NextSeq run ID'
ws7['P1']='Result'
ws7['Q1']='Comments'

ws8['A4']='Lab number'
ws8['B4']='Patient Name'
ws8['C4']='Reason for referral'
ws8['D4']='NGS Wks'
ws8['E4']='Nextseq run ID'
ws8['F4']='Checker 1 initials and date'
ws8['G4']='Checker 2 initials and date'

ws8['A8']='RMATS output'

ws8['C1']='Analysis Sheet- Pan RNA Cancer Fusion Panel Validation'



def get_referral_list(referral):
	
	referral_file=pandas.read_csv(/data/diagnostics/pipelines/SomaticFusion/SomaticFusion-0.0.4/Referrals/referral+".txt", sep="\t")
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


	ws6["A39"]="Arriba results"
	for referral in referral_list:
	
		if (os.stat("../"+ntc+"/Results/arriba/"+ntc+"_fusion_report_"+referral+"_arriba.txt").st_size!=0):
      	
			NTC_arriba_report=pandas.read_csv("../"+ntc+"/Results/arriba/"+ntc+"_fusion_report_"+referral+"_arriba.txt", sep="\t")

	
			for row in dataframe_to_rows(NTC_arriba_report, header=True, index=False):
				ws6.append(row)


	return(NTC_star_fusion_report, NTC_arriba_report)

	
		




def get_star_fusion_report(referral_list, ntc_star_fusion_report):
	ws1["A4"]="STAR-Fusion results"



	#create empty pandas dataframe and append

	star_fusion_report_final= pandas.DataFrame(columns=["Fusion_Name", "Split_Read_Count", "Spanning_Read_Count", "Left_Breakpoint", "Right_Breakpoint", "SpliceType", "LargeAnchorSupport", "FFPM", "LeftBreakEntropy", "RightBreakEntropy","CDS_Left_ID", "CDS_Left_Range", "CDS_Right_ID", "CDS_Right_Range", "Prot_Fusion_Type","Num_WT_Fragments_Left", "Num_WT_Fragments_Right", "Fusion_Allelic_Fraction"]) 

	for referral in referral_list:

		if (os.stat("./Results/STAR_Fusion/"+sampleId+"_fusion_report_"+referral+".txt").st_size!=0):
      
			star_fusion_report=pandas.read_csv("Results/STAR_Fusion/"+sampleId+"_fusion_report_"+referral+".txt", sep="\t")
			star_fusion_report=pandas.DataFrame(star_fusion_report)
			del star_fusion_report['Unnamed: 0']
			star_fusion_report_final=star_fusion_report_final.append(star_fusion_report, sort=False)
	

	for row in dataframe_to_rows(star_fusion_report_final, header=True, index=False):
		ws1.append(row)


	
		




def get_arriba_fusion_report(referral_list):

	ws1["A39"]="Arriba results"

	arriba_report_final= pandas.DataFrame(columns=["gene1", "gene2", "strand1(gene/fusion)", "strand2(gene/fusion)", "breakpoint1", "breakpoint2", "site1", "site2", "type", "direction1", "direction2", "split_reads1", "split_reads2", "discordant_mates", "coverage1", "coverage2", "confidence", "closest_genomic_breakpoint1", "closest_genomic_breakpoint2", "filters", "fusion_transcript", "reading_frame", "peptide_sequence", "read_identifiers"])

	for referral in referral_list:

		if (os.stat("./Results/arriba/"+sampleId+"_fusion_report_"+referral+"_arriba.txt").st_size!=0):
      	
			arriba_report=pandas.read_csv("Results/arriba/"+sampleId+"_fusion_report_"+referral+"_arriba.txt", sep="\t")
			del arriba_report['Unnamed: 0']
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
		ws2["A20"]="without duplicates"
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



def get_quality_metrics(referral):
	ws5['C1']='Analysis Sheet-Pan RNA Cancer Fusion Panel Validation'
	ws5['C1'].font=Font(size=16)
	ws5["A4"]="Lab number" 
	ws5["B4"]="Patient name"
	ws5["C4"]="Tumour %"
	ws5["D4"]="Reason for Referral"
	ws5["E4"]="Qubit RNA conc after pre-processing (ng/ul)"
	ws5["F4"]="DV200"
	ws5["G4"]="Due date"

	ws5["A7"]="NGS wks"
	ws5["B7"]="NextSeq run ID"
	ws5["C7"]="NTC check 1"
	ws5["D7"]="NTC check 2"
	ws5["E7"]="Analysed by:"
	ws5["F7"]="Checked by:"

	ws5["A14"]= "Gene fusion genuine calls"
	ws5["A15"]= "STAR-Fusion results"
	
	ws5["A16"]="Fusion_Name"
	ws5["B16"]="Split_Read_Count"
	ws5["C16"]="Spanning_Read_Count"
	ws5["D16"]="Left_Breakpoint"
	ws5["E16"]="Right_Breakpoint:"
	ws5["F16"]="SpliceType"
	ws5["G16"]="LargeAnchorSupport"
	ws5["H16"]="FFPM"
	ws5["I16"]="Prot_Fusion_Type"
	ws5["J16"]="Fusion_Allelic_Fraction"
	ws5["K16"]="Conclusion checker 1"
	ws5["L16"]="Conclusion checker 2"

	
	ws5["A22"]="Arriba results"
	ws5["A23"]="#gene1"
	ws5["B23"]="gene2"
	ws5["C23"]="strand1(gene/fusion)"
	ws5["D23"]="strand2(gene/fusion)"
	ws5["E23"]="breakpoint1"
	ws5["F23"]="breakpoint2"
	ws5["G23"]="site1"
	ws5["H23"]="site2"
	ws5["I23"]="type"
	ws5["J23"]="split_reads1"
	ws5["K23"]="split_reads2"
	ws5["L23"]="discordant_mates"
	ws5["M23"]="confidence"
	ws5["N23"]="filters"
	ws5["O23"]="Conclusion checker 1"
	ws5["P23"]="Conclusion checker 2"


	ws5["A28"]="Gene fusions quality metrics"


	ws5["A30"]="Post PCR1 Qubit"
	ws5["A31"]="Total reads-duplicates removed"
	ws5["A32"]="% mapped reads- duplicates removed"


	ws5["B29"]="Value"
	ws5["C29"]="Conclusion checker 1"
	ws5["D29"]="Conclusion checker 2"


	ws5["A38"]="Intragenic fusions(RMATS) genuine calls"
	ws5["A39"]="geneSymbol"
	ws5["B39"]="chr"
	ws5["C39"]="exonStart_0base"
	ws5["D39"]="exonEnd"
	ws5["E39"]="IJC_SAMPLE_1"
	ws5["F39"]="SJC_SAMPLE_1"
	ws5["G39"]="FDR"
	ws5["H39"]="IncLevelDifference"
	ws5["I39"]="Fusion Level"
	ws5["J39"]="Conclusion Checker 1"
	ws5["K39"]="Conclusion Checker 2"

	ws5["A45"]="Gene_exon"
	ws5["A45"]="Gene_exon"
	ws5["B45"]="CHR"
	ws5["C45"]="START"
	ws5["D45"]="END"
	ws5["E45"]="No.unique reads"
	ws5["F45"]="Conclusion 1"
	ws5["G45"]="Conclusion 2"












	border_a=Border(left=Side(border_style=BORDER_MEDIUM), right=Side(border_style=BORDER_MEDIUM), top=Side(border_style=BORDER_MEDIUM), bottom=Side(border_style=BORDER_MEDIUM))
	position=['A4','B4','C4','D4','E4','F4','G4',
	'A5','B5','C5','D5','E5','F5', 'G5',
        'A7','B7','C7','D7','E7','F7',
        'A8','B8','C8','D8','E8','F8',
        'A16','B16','C16','D16','E16','F16','G16', 'H16', 'I16', 'J16', 'K16', 'L16',
        'A23','B23','C23','D23','E23','F23','G23', 'H23', 'I23', 'J23', 'K23', 'L23', 'M23', 'N23', 'O23', 'P23',
        'B29', 'C29', 'D29',
        'A30', 'B30', 'C30', 'D30',
        'A31', 'B31', 'C31', 'D31',
        'A32', 'B32', 'C32', 'D32',
        'A39','B39','C39','D39','E39','F39','G39', 'H39', 'I39', 'J39', 'K39',
        'A45','B45','C45','D45','E45','F45','G45']
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
	position=['A5','B5','C5','D5','E5','F5','G5','H5','I5','J5','K5','L5','M5','N5','O5','P5','Q5','R5','S5',
	'A40','B40','C40','D40','E40','F40','G40','H40','I40','J40','K40','L40','M40','N40','O40','P40','Q40','R40','S40','T40','U40','V40','W40','X40','Y40']
	for cell in position:
		ws1[cell].border=border_a

	colour= PatternFill("solid", fgColor="DCDCDC")
	position= ['A5','B5','C5','D5','E5','F5','G5','H5','I5','J5','K5','L5','M5','N5','O5','P5','Q5','R5','S5',
	'A40','B40','C40','D40','E40','F40','G40','H40','I40','J40','K40','L40','M40','N40','O40','P40','Q40','R40','S40','T40','U40','V40','W40','X40','Y40']
	for cell in position:
		ws1[cell].fill=colour



	font_bold=Font(bold=True)
	position= ['A3','A4','B4','C4','D4','E4','F4','G4', 'A7','B7','C7','D7','E7','F7','A13', 'F13', 'J13']
	for cell in position:
		ws5[cell].font=font_bold


	font_bold=Font(bold=True)
	position= ['A4', 'A39',
	'A5','B5','C5','D5','E5','F5','G5','H5','I5','J5','K5','L5','M5','N5','O5','P5','Q5','R5','S5',
	'A40','B40','C40','D40','E40','F40','G40','H40','I40','J40','K40','L40','M40','N40','O40','P40','Q40','R40','S40','T40','U40','V40','W40','X40','Y40']
	for cell in position:
		ws1[cell].font=font_bold




	font_bold=Font(bold=True)
	position= ['A1','B1','C1','D1','E1','F1','G1','H1','I1','J1','K1','L1','M1','N1','O1','P1','Q1','R1','S1','T1','U1','V1']
	for cell in position:
		ws7[cell].font=font_bold


	colour= PatternFill("solid", fgColor="DCDCDC")
	position= ['A1','B1','C1','D1','E1','F1','G1','H1','I1','J1','K1','L1','M1','N1','O1','P1','Q1','R1','S1','T1','U1','V1']
	for cell in position:
		ws7[cell].fill=colour



	border_a=Border(left=Side(border_style=BORDER_MEDIUM), right=Side(border_style=BORDER_MEDIUM), top=Side(border_style=BORDER_MEDIUM), bottom=Side(border_style=BORDER_MEDIUM))
	position=['A1','B1','C1','D1','E1','F1','G1','H1','I1','J1','K1','L1','M1','N1','O1','P1','Q1','R1','S1','T1','U1','V1']
	for cell in position:
		ws7[cell].border=border_a




	colour= PatternFill("solid", fgColor="00CCFFFF")
	position= ['A4','B4','C4','D4','E4','F4', 'G4','A7','B7','C7','D7']
	for cell in position:
		ws5[cell].fill=colour






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
	ws5["G14"]=total_reads_value
	aligned_reads_value=total_reads_aligned.iloc[0,0]
	ws5["G15"]=aligned_reads_value
	


def get_alignment_metrics_rmdup():
	with open ("../"+sampleId+"_AlignmentSummaryMetrics_removed_duplicates.txt") as file_rmdup:
		for line in file_rmdup:
			if line.startswith("CATEGORY"):
				headers_rmdup=line.split('\t')
			if line.startswith("PAIR"):
				pair_list_rmdup=line.split('\t')


	alignment_metrics_rmdup=pandas.DataFrame([pair_list_rmdup], columns=headers_rmdup)
	total_reads_rmdup=alignment_metrics_rmdup[['PF_READS']]
	total_reads_aligned_rmdup=alignment_metrics_rmdup[['PF_HQ_ALIGNED_READS']]
	total_reads_value_rmdup=total_reads_rmdup.iloc[0,0]
	ws5["G16"]=total_reads_value_rmdup
	aligned_reads_value_rmdup=total_reads_aligned_rmdup.iloc[0,0]
	ws5["G17"]=aligned_reads_value_rmdup
	



def get_met_exon_skipping():
	

	
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

	ws8["N9"]="Fusion level"
	ws8["O9"]="Conclusion checker 1"
	ws8["P9"]="Conclusion checker 2"
	ws8["Q9"]="Comments"


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
	position=['A4','B4','C4','D4','E4','F4','G4','A5','B5','C5','D5','E5','F5','G5','A9','B9','C9','D9','E9','F9','G9','H9','I9','J9','K9','L9','M9','N9','O9','P9','Q9']
	for cell in position:
		ws8[cell].border=border_a

	colour= PatternFill("solid", fgColor="DCDCDC")
	position= ['A9','B9','C9','D9','E9','F9','G9','H9','I9','J9','K9','L9','M9','N9','O9','P9','Q9']
	for cell in position:
		ws8[cell].fill=colour


	colour= PatternFill("solid", fgColor="00CCFFFF")
	position= ['A4','B4','C4','D4','E4','F4', 'G4']
	for cell in position:
		ws8[cell].fill=colour

			


if __name__ == "__main__":

	seqId=sys.argv[1]
	sampleId=sys.argv[2]
	referral=sys.argv[3]
	ntc=sys.argv[4]


	referral_list=get_referral_list(referral)
	ntc_star_fusion_report, ntc_arriba_report=get_NTC_fusion_report(ntc)
	get_star_fusion_report(referral_list, ntc_star_fusion_report)
	get_arriba_fusion_report(referral_list)
	ntc_total_average_depth, ntc_total_average_depth_rmdup=get_ntc_total_coverage()
	get_total_coverage(ntc_total_average_depth, ntc_total_average_depth_rmdup)
	ntc_average_depth, ntc_average_depth_rmdup=ntc_get_coverage()
	get_coverage(referral, ntc_average_depth, ntc_average_depth_rmdup)
	get_quality_metrics(referral)
	get_alignment_metrics()
	get_alignment_metrics_rmdup()
	if (referral=="MET" or referral=="ALL"):
		get_met_exon_skipping()
	wb.save(sampleId+"_"+referral +"-TESTER.xlsx")
