
import pandas
import sys



'''
 Description: combines the information for total reads and total aligned reads for all samples on a run
 Author: Laura McCluskey
 Date: 10th July 2020

'''


seq_id = sys.argv[1]

aligned_total_reads=[]
percentage_aligned=[]
aligned_total_reads_rmdup=[]
percentage_aligned_rmdup=[]

samples = pandas.read_csv("samples_list.txt", sep="\t")
sampleList=samples["sampleId"].tolist()

for sampleId in sampleList:
	with open (sampleId+"_AlignmentSummaryMetrics.txt") as file:
		for line in file:
			if line.startswith("CATEGORY"):
				headers=line.split('\t')
			if line.startswith("PAIR"):
				pair_list=line.split('\t')


	alignment_metrics=pandas.DataFrame([pair_list], columns=headers)
	total_reads=alignment_metrics[['PF_READS']]
	total_reads_value=total_reads.iloc[0,0]
	total_reads_aligned=alignment_metrics[['PF_HQ_ALIGNED_READS']]
	aligned_reads_value=total_reads_aligned.iloc[0,0]
	aligned_total_reads.append(aligned_reads_value)
	percentage_reads_aligned=float((float(aligned_reads_value)/float(total_reads_value))*100)
	percentage_aligned.append(percentage_reads_aligned)
	

	with open (sampleId+"_AlignmentSummaryMetrics_rmdup.txt") as file_rmdup:
		for line in file_rmdup:
			if line.startswith("CATEGORY"):
				headers_rmdup=line.split('\t')
			if line.startswith("PAIR"):
				pair_list_rmdup=line.split('\t')


	alignment_metrics_rmdup=pandas.DataFrame([pair_list_rmdup], columns=headers_rmdup)
	total_reads_rmdup=alignment_metrics_rmdup[['PF_READS']]
	total_reads_value_rmdup=total_reads_rmdup.iloc[0,0]
	total_reads_aligned_rmdup=alignment_metrics_rmdup[['PF_HQ_ALIGNED_READS']]
	aligned_reads_value_rmdup=total_reads_aligned_rmdup.iloc[0,0]
	aligned_total_reads_rmdup.append(aligned_reads_value_rmdup)
	percentage_reads_aligned_rmdup=float((float(aligned_reads_value_rmdup)/float(total_reads_value_rmdup))*100)
	percentage_aligned_rmdup.append(percentage_reads_aligned_rmdup)



dict={'sample':sampleList, 'aligned_reads':aligned_total_reads, '%reads_aligned': percentage_aligned, 'unique_reads_aligned':aligned_total_reads_rmdup, '%unique_reads_aligned': percentage_aligned_rmdup}
reads_table=pandas.DataFrame(dict, columns=['sample', 'aligned_reads', '%reads_aligned', 'unique_reads_aligned', '%unique_reads_aligned'])

reads_table.to_csv(seq_id+'-aligned_reads.csv', index=False)