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
ws5=wb.create_sheet("Quality metrics")
ws1= wb.create_sheet("fusion report")
ws6= wb.create_sheet("NTC fusion report")
ws2=wb.create_sheet("total_coverage")
ws3=wb.create_sheet("coverage_with_duplicates")
ws4=wb.create_sheet("coverage_without_duplicates")

ws1["E2"]="Checker 1 initals and date"
ws1["G2"]="Checker 2 initals and date"


ws7['A1']='Date Received'
ws7['B1']='Date Requested'
ws7['C1']='Due Date'
ws7['D1']='LABNO'
ws7['E1']='Patient name'
ws7['F1']='DOB'
ws7['G1']='Reason for referral'
ws7['H1']='Referring Clinician'
ws7['I1']='Hospital'
ws7['J1']='Date reported'
ws7['K1']='TAT'
ws7['L1']='No of days in histo'
ws7['M1']='Block/Slide/RNA'
ws7['N1']='% Tumour'
ws7['O1']='Result'
ws7['P1']='NGS Worksheet'
ws7['Q1']='Qubit RNA conc. (ng/ul)'
ws7['R1']='Total RNA input'
ws7['S1']='Post PCR1 Qubit'
ws7['T1']='Date of NextSeq run'
ws7['U1']='NextSeq run ID'
ws7['V1']='Comments'



def get_NTC_fusion_report(ntc):
	
	ws6["A4"]="NTC STAR-Fusion results"
	if (referral=="ALL"):
		NTC_star_fusion_report=pandas.read_csv("../"+ntc+"/fusionReport/"+ntc+"_fusionReport.txt", sep="\t")
		

	else:
		if (os.stat("../"+ntc+"/Results/STAR_Fusion/"+ntc+"_fusion_report_"+referral+".txt").st_size!=0):
      
			NTC_star_fusion_report=pandas.read_csv("../"+ntc+"/Results/STAR_Fusion/"+ntc+"_fusion_report_"+referral+".txt", sep="\t")

			NTC_star_fusion_report=pandas.DataFrame(NTC_star_fusion_report)
	

	for row in dataframe_to_rows(NTC_star_fusion_report, header=True, index=False):
		ws6.append(row)


	ws6["A39"]="Arriba results"
	if (referral=="ALL"):
		NTC_arriba_report=pandas.read_csv("../"+ntc+"/"+ntc+"_fusions_adapted.tsv", sep="\t")
	else:
		if (os.stat("../"+ntc+"/Results/arriba/"+ntc+"_fusion_report_"+referral+"_arriba.txt").st_size!=0):
      	
			NTC_arriba_report=pandas.read_csv("../"+ntc+"/Results/arriba/"+ntc+"_fusion_report_"+referral+"_arriba.txt", sep="\t")

	
	for row in dataframe_to_rows(NTC_arriba_report, header=True, index=False):
		ws6.append(row)
	



	return(NTC_star_fusion_report, NTC_arriba_report)

	
		




def get_star_fusion_report(referral, ntc_star_fusion_report):
	ws1["A4"]="STAR-Fusion results"
	if (referral=="ALL"):
		star_fusion_report=pandas.read_csv("fusionReport/"+sampleId+"_fusionReport.txt", sep="\t")
		

	else:
		if (os.stat("./Results/STAR_Fusion/"+sampleId+"_fusion_report_"+referral+".txt").st_size!=0):
      
			star_fusion_report=pandas.read_csv("Results/STAR_Fusion/"+sampleId+"_fusion_report_"+referral+".txt", sep="\t")

	star_fusion_report=pandas.DataFrame(star_fusion_report)
	

	for row in dataframe_to_rows(star_fusion_report, header=True, index=False):
		ws1.append(row)


	
		




def get_arriba_fusion_report(referral):

	ws1["A39"]="Arriba results"
	if (referral=="ALL"):
		arriba_report=pandas.read_csv(sampleId+"_fusions_adapted.tsv", sep="\t")
	else:
		if (os.stat("./Results/arriba/"+sampleId+"_fusion_report_"+referral+"_arriba.txt").st_size!=0):
      	
			arriba_report=pandas.read_csv("Results/arriba/"+sampleId+"_fusion_report_"+referral+"_arriba.txt", sep="\t")

	for row in dataframe_to_rows(arriba_report, header=True, index=False):
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
	ws5["A3"]="Patient information"
	ws5["A4"]="Lab number"
	ws5["A5"]=sampleId
	ws5["B4"]="Patient name"
	ws5["C4"]="Referral genes/variants"
	ws5["D4"]="NGS wks"
	ws5["E4"]="NextSeq runId"
	ws5["F4"]="Checker 1 initials and date"
	ws5["G4"]="Checker 2 initials and date"
	ws5["C5"]=referral

	ws5["A13"]="Pre-run quality metrics"
	ws5["A14"]="Tumour %"
	ws5["A15"]="Qubit[RNA] ng/ul"
	ws5["A16"]="Input amount ng"
	ws5["A17"]="DV200"
	ws5["A18"]="Post PCR qubit"
	ws5["C13"]="Threshold"
	ws5["H13"]="Threshold"

	ws5["F13"]="Post-run quality metrics"
	ws5["F14"]="Total reads"
	ws5["F15"]="Total mapped reads"
	ws5["F16"]="Total reads-duplicates removed"
	ws5["F17"]="Total mapped reads-duplicates removed"
	ws5["F18"]="Mean insert size"

	ws5["J13"]="Run quality metrics"
	ws5["J14"]="Cluster Density"
	ws5["J15"]="Cluster passing filter"
	ws5["J16"]="Q30"
	ws5["J17"]="Yield"

	border_a=Border(left=Side(border_style=BORDER_MEDIUM), right=Side(border_style=BORDER_MEDIUM), top=Side(border_style=BORDER_MEDIUM), bottom=Side(border_style=BORDER_MEDIUM))
	position=['A4','B4','C4','D4','E4','F4','G4',
	'A5','B5','C5','D5','E5','F5','G5',
	'A13','A14','A15','A16','A17','A18',
	'B13','B14','B15','B16','B17','B18',
	'F13','F14','F15','F16','F17','F18',
	'G13','G14','G15','G16','G17','G18',
	'J13','J14','J15','J16','J17',
	'K13','K14','K15','K16','K17']
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
	position= ['A3','A4','B4','C4','D4','E4','F4','G4', 'A13', 'F13', 'J13']
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
	position= ['A4','B4','C4','D4','E4','F4', 'G4']
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
	ws5["G16"]=total_reads_value_rmdup
	aligned_reads_value_rmdup=total_reads_aligned_rmdup.iloc[0,0]
	ws5["G17"]=aligned_reads_value_rmdup
	



def get_met_exon_skipping():
	
	ws8=wb.create_sheet("MET_exon_skipping")
	ws8["D3"]="Checker 1 initials and date"
	ws8["F3"]="Checker 2 initials and date"
	ws8["D5"]=""
	ws8["F5"]=""

	
	rmats=pandas.DataFrame()

	with open (seqId+"_"+sampleId+"_RMATS_Report.txt") as file:
		for line in file:
			if line.startswith("#"):
				continue
			else:
				rmats_line=line.split('\t')
				ws8.append(rmats_line)

	ws8.column_dimensions['A'].width=30
	ws8.column_dimensions['B'].width=15
	ws8.column_dimensions['C'].width=15
	ws8.column_dimensions['D'].width=20
	ws8.column_dimensions['E'].width=20
	ws8.column_dimensions['F'].width=20
	ws8.column_dimensions['G'].width=20
	ws8.column_dimensions['H'].width=70
	ws8.column_dimensions['I'].width=70
	ws8.column_dimensions['J'].width=20
	ws8.column_dimensions['K'].width=20
	ws8.column_dimensions['L'].width=70
	ws8.column_dimensions['M'].width=20



			


if __name__ == "__main__":

	seqId=sys.argv[1]
	sampleId=sys.argv[2]
	referral=sys.argv[3]
	ntc=sys.argv[4]



	ntc_star_fusion_report, ntc_arriba_report=get_NTC_fusion_report(ntc)
	get_star_fusion_report(referral, ntc_star_fusion_report)
	get_arriba_fusion_report(referral)
	ntc_total_average_depth, ntc_total_average_depth_rmdup=get_ntc_total_coverage()
	get_total_coverage(ntc_total_average_depth, ntc_total_average_depth_rmdup)
	ntc_average_depth, ntc_average_depth_rmdup=ntc_get_coverage()
	get_coverage(referral, ntc_average_depth, ntc_average_depth_rmdup)
	get_quality_metrics(referral)
	get_alignment_metrics()
	get_alignment_metrics_rmdup()
#	if (referral=="MET" or referral=="ALL"):
#		get_met_exon_skipping()
	wb.save(sampleId+"_"+referral +".xlsx")
