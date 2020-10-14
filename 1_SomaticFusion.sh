
#!/bin/bash
# set -euo pipefail
#PBS -l walltime=40:00:00
#PBS -l ncpus=12
PBS_O_WORKDIR=(`echo $PBS_O_WORKDIR | sed "s/^\/state\/partition1//" `)
cd $PBS_O_WORKDIR

# Description: Wrapper script for calling fusion genes in RNA-seq data \
# from clinical samples using STAR-Fusion.
# Authors: Christopher Medway and Laura McCluskey
# Date: 27th March 2020
# Usage: qsub run_star-fusion.sh [inside sample dir with .variables and \
# .fastq.gz files]

version=1.0.0

# source variables file
. *.variables

# copy the panel & pipeline variables locally and source
cp /data/diagnostics/pipelines/SomaticFusion/SomaticFusion-$version/$panel/$panel.variables . 
cp /data/diagnostics/pipelines/SomaticFusion/SomaticFusion-$version/SomaticFusion.config .
cp /data/diagnostics/pipelines/SomaticFusion/SomaticFusion-$version/make-fusion-report.py .
cp /data/diagnostics/pipelines/SomaticFusion/SomaticFusion-$version/RNA_fusion_group_file.txt .
cp /data/diagnostics/pipelines/SomaticFusion/SomaticFusion-$version/RNAFusion-ROI.bed .
cp /data/diagnostics/pipelines/SomaticFusion/SomaticFusion-$version/fusion_report_referrals.py .
cp /data/diagnostics/pipelines/SomaticFusion/SomaticFusion-$version/make_worksheets.py .
cp /data/diagnostics/pipelines/SomaticFusion/SomaticFusion-$version/rmats/RMATS.config .

gatk3=/share/apps/GATK-distros/GATK_3.8.0/GenomeAnalysisTK.jar
minMQS=20
minBQS=10

vendorCaptureBed=./180702_HG19_PanCancer_EZ_capture_targets.bed
vendorPrimaryBed=./180702_HG19_PanCancer_EZ_primary_targets.bed

. $panel.variables
. SomaticFusion.config

# set conda env
source "$conda_bin_path"/activate SomaticFusion


#######################
# Preprocessing FASTQ #
#######################


for fastqPair in $(ls "$sampleId"_S*.fastq.gz | cut -d_ -f1-4 | sort | uniq); do

    #parse fastq filenames
    laneId=$(echo "$fastqPair" | cut -d_ -f4)
    read1Fastq=$(ls "$fastqPair"_R1_*fastq.gz)
    read2Fastq=$(ls "$fastqPair"_R2_*fastq.gz)

    #trim adapters
    cutadapt -a $read1Adapter -A $read2Adapter -m 35 -o "$seqId"_"$sampleId"_"$laneId"_R1.fastq -p "$seqId"_"$sampleId"_"$laneId"_R2.fastq "$read1Fastq" "$read2Fastq"


    fastqc -d /state/partition1/tmpdir --threads 12 --extract "$seqId"_"$sampleId"_"$laneId"_R1.fastq
    fastqc -d /state/partition1/tmpdir --threads 12 --extract "$seqId"_"$sampleId"_"$laneId"_R2.fastq

    mv "$seqId"_"$sampleId"_"$laneId"_R1_fastqc/summary.txt "$seqId"_"$sampleId"_"$laneId"_R1_fastqc.txt
    mv "$seqId"_"$sampleId"_"$laneId"_R2_fastqc/summary.txt "$seqId"_"$sampleId"_"$laneId"_R2_fastqc.txt




############
# Run STAR #
############

sampleId_without_conc="$(echo $sampleId | cut -d'_' -f1)"


STAR --chimSegmentMin 12 \
     --chimJunctionOverhangMin 12 \
     --chimSegmentReadGapMax 3 \
     --alignSJDBoverhangMin 10 \
     --alignMatesGapMax 200000 \
     --chimOutType WithinBAM \
     --alignIntronMax 200000  \
     --alignSJstitchMismatchNmax 5 -1 5 5  \
     --twopassMode Basic \
     --outSAMtype BAM Unsorted \
     --readFilesIn "$seqId"_"$sampleId"_"$laneId"_R1.fastq "$seqId"_"$sampleId"_"$laneId"_R2.fastq \
     --genomeDir /share/apps/star-fusion/GRCh37_gencode_v19_CTAT_lib_Mar272019.plug-n-play/ctat_genome_lib_build_dir/ref_genome.fa.star.idx  \
     --outFileNamePrefix "$seqId"_"$sampleId"_"$laneId"_ \
     --outSAMattrRGline ID:"$sampleId" SM:"$sampleId_without_conc"
done


###################
# Run STAR-Fusion #
###################


## STAR-fusion required a samples_file in order to handle multi-lane runs
R1=$(for i in ./"$sampleId"_*.fastq.gz; do echo $i | grep "R1"; done)
R2=$(for i in ./"$sampleId"_*.fastq.gz; do echo $i | grep "R2"; done)
SAMPLE=$(for i in $R1;do echo $sampleId;done)
paste <(printf %s "$SAMPLE") <(printf %s "$R1") <(printf %s "$R2") > "$sampleId".samples

STAR-Fusion --genome_lib_dir $starfusion_lib \
            --samples_file $sampleId.samples \
            --output_dir ./STAR-Fusion/ \
            --denovo_reconstruct \
            --examine_coding_effect \
            --FusionInspector inspect \
            --CPU $ncpus \
            --min_FFPM 1 \
            --require_LDAS 0

cd STAR-Fusion

/home/transfer/miniconda3/envs/SomaticFusion/lib/STAR-Fusion/FusionInspector/FusionInspector --fusions star-fusion.fusion_predictions.abridged.coding_effect.tsv  --out_prefix finspector  --min_junction_reads 1  --min_novel_junction_support 3  --min_spanning_frags_only 5  --vis  --max_promiscuity 10  --out_dir /share/data/results/"$seqId"/RocheSTFusion/"$sampleId"/STAR-Fusion/FusionInspector-validate  --genome_lib_dir /share/apps/star-fusion/GRCh37_gencode_v19_CTAT_lib_Mar272019.plug-n-play/ctat_genome_lib_build_dir/  --CPU 12  --samples_file /share/data/results/"$seqId"/RocheSTFusion/"$sampleId"/STAR-Fusion/starF.target_samples.txt  --include_Trinity  --annotate  --examine_coding_effect --require_LDAS 0

cd ..

###########################
# Generate Fusion Reports #
###########################


python make-fusion-report.py \
    --sampleId $sampleId \
    --seqId $seqId \
    --panel $panel \
    --ip $(hostname --ip-address)

#deactivate conda env
source "$conda_bin_path"/deactivate



############################
# merge and sort bam files #
############################


samtools merge "$sampleId"_Aligned_out.bam *Aligned.out.bam 

samtools sort "$sampleId"_Aligned_out.bam "$sampleId"_Aligned_sorted

samtools index "$sampleId"_Aligned_sorted.bam



##############
# Run arriba #
##############

source "$conda_bin_path"/activate SomaticFusion

arriba -x "$sampleId"_Aligned_sorted.bam \
-g /share/apps/star-fusion/GRCh37_gencode_v19_CTAT_lib_Mar272019.plug-n-play/ctat_genome_lib_build_dir/ref_annot.gtf \
-a /share/apps/star-fusion/GRCh37_gencode_v19_CTAT_lib_Mar272019.plug-n-play/ctat_genome_lib_build_dir/ref_genome.fa \
-o "$sampleId"_fusions_adapted.tsv \
-b /home/transfer/miniconda3/envs/arriba/var/lib/arriba/blacklist_hg19_hs37d5_GRCh37_2018-11-04.tsv.gz \
-O "$sampleId"_fusions_discarded_adapted.tsv \
-R 0 

source "$conda_bin_path"/deactivate


###############################
# Calculate depth of coverage #
###############################


/share/apps/jre-distros/jre1.8.0_131/bin/java -XX:GCTimeLimit=50 -XX:GCHeapFreeLimit=10 -Djava.io.tmpdir=/state/partition1/tmpdir -Xmx4g -jar $gatk3 \
   -T DepthOfCoverage \
   -R /share/apps/star-fusion/star-fusion-ref.fa \
   -I  "$sampleId"_Aligned_sorted.bam \
   -L RNAFusion-ROI.bed \
   -o "$seqId"_"$sampleId"_DepthOfCoverage \
   --countType COUNT_FRAGMENTS \
   --minMappingQuality $minMQS \
   --minBaseQuality $minBQS \
   -ct 30  \
   --omitLocusTable \
   -dt NONE \
   -U ALLOW_N_CIGAR_READS \
   -rf ReassignOneMappingQuality -RMQF 255 -RMQT 60 \
   -rf MappingQualityUnavailable



############################
# Run CoverageCalculatorPy #
############################

sed 's/:/\t/g'  "$seqId"_"$sampleId"_DepthOfCoverage | grep -v 'Locus' | sort -k1,1 -k2,2n | bgzip > "$seqId"_"$sampleId"_DepthOfCoverage.gz
tabix -b 2 -e 2 -s 1 "$seqId"_"$sampleId"_DepthOfCoverage.gz


source "$conda_bin_path"/activate CoverageCalculatorPy


python /home/transfer/pipelines/CoverageCalculatorPy/CoverageCalculatorPy.py \
-B RNAFusion-ROI.bed \
-D /data/results/"$seqId"/RocheSTFusion/"$sampleId"/"$seqId"_"$sampleId"_DepthOfCoverage.gz \
--padding 0 \
--groupfile RNA_fusion_group_file.txt \
--outname "$sampleId"_coverage 



########################
# Calculate qc metrics #
########################


if [ ! -f ../"$seqId"_"$sampleId"_HsMetrics.txt ]; then

/share/apps/jre-distros/jre1.8.0_131/bin/java -XX:GCTimeLimit=50 -XX:GCHeapFreeLimit=10 -Djava.io.tmpdir=/state/partition1/tmpdir -Xmx2g \
     -jar /share/apps/picard-tools-distros/picard-tools-2.18.5/picard.jar CollectHsMetrics \
     I=/data/results/"$seqId"/RocheSTFusion/"$sampleId"/"$sampleId"_Aligned_sorted.bam \
     O=../"$seqId"_"$sampleId"_HsMetrics.txt \
     R=/share/apps/star-fusion/GRCh37_gencode_v19_CTAT_lib_Mar272019.plug-n-play/ctat_genome_lib_build_dir/ref_genome.fa \
     BAIT_INTERVALS=/share/apps/star-fusion/RochePanCancer_capture.interval_list \
     TARGET_INTERVALS=/share/apps/star-fusion/RochePanCancer_primary.interval_list \
     MAX_RECORDS_IN_RAM=2000000 \
     TMP_DIR=/state/partition1/tmpdir \
     MINIMUM_MAPPING_QUALITY=$minMQS \
     MINIMUM_BASE_QUALITY=$minBQS \
     CLIP_OVERLAPPING_READS=false

fi



#Alignment metrics: library sequence similarity
/share/apps/jre-distros/jre1.8.0_131/bin/java -XX:GCTimeLimit=50 -XX:GCHeapFreeLimit=10 -Djava.io.tmpdir=/state/partition1/tmpdir -Xmx2g \
    -jar /share/apps/picard-tools-distros/picard-tools-2.18.5/picard.jar CollectAlignmentSummaryMetrics \
    R=/share/apps/star-fusion/GRCh37_gencode_v19_CTAT_lib_Mar272019.plug-n-play/ctat_genome_lib_build_dir/ref_genome.fa \
    ADAPTER_SEQUENCE=AGATCGGAAGAGC \
    I="$sampleId"_Aligned_sorted.bam \
    O=../"$sampleId"_AlignmentSummaryMetrics.txt \
    MAX_RECORDS_IN_RAM=2000000 \
TMP_DIR=/state/partition1/tmpdir


source "$conda_bin_path"/deactivate




#######################################################
#remove duplicates and recalculate qc metrics/coverage#
#######################################################

#remove duplicates

/share/apps/jre-distros/jre1.8.0_131/bin/java \
    -XX:GCTimeLimit=50 \
    -XX:GCHeapFreeLimit=10 \
    -Djava.io.tmpdir=/state/partition1/tmpdir \
    -Xmx2g \
    -jar /share/apps/picard-tools-distros/picard-tools-2.18.5/picard.jar \
    MarkDuplicates \
    I="$sampleId"_Aligned_sorted.bam \
    OUTPUT="$sampleId"_rmdup.bam \
    METRICS_FILE="$sampleId"_markDuplicatesMetrics.txt \
    CREATE_INDEX=true \
    MAX_RECORDS_IN_RAM=2000000 \
    VALIDATION_STRINGENCY=SILENT \
    TMP_DIR=/state/partition1/tmpdir \
    QUIET=true \
    VERBOSITY=ERROR \
    REMOVE_DUPLICATES=TRUE


#Alignment metrics: library sequence similarity
/share/apps/jre-distros/jre1.8.0_131/bin/java -XX:GCTimeLimit=50 -XX:GCHeapFreeLimit=10 -Djava.io.tmpdir=/state/partition1/tmpdir -Xmx2g \
    -jar /share/apps/picard-tools-distros/picard-tools-2.18.5/picard.jar CollectAlignmentSummaryMetrics \
    R=/share/apps/star-fusion/GRCh37_gencode_v19_CTAT_lib_Mar272019.plug-n-play/ctat_genome_lib_build_dir/ref_genome.fa \
    ADAPTER_SEQUENCE=AGATCGGAAGAGC \
    I="$sampleId"_rmdup.bam \
    O=../"$sampleId"_AlignmentSummaryMetrics_rmdup.txt \
    MAX_RECORDS_IN_RAM=2000000 \
TMP_DIR=/state/partition1/tmpdir



#Calculate depth of coverage

/share/apps/jre-distros/jre1.8.0_131/bin/java -XX:GCTimeLimit=50 -XX:GCHeapFreeLimit=10 -Djava.io.tmpdir=/state/partition1/tmpdir -Xmx4g -jar $gatk3 \
   -T DepthOfCoverage \
   -R /share/apps/star-fusion/star-fusion-ref.fa \
   -I  "$sampleId"_rmdup.bam \
   -L RNAFusion-ROI.bed \
   -o "$sampleId"_rmdup_DepthOfCoverage \
   --countType COUNT_FRAGMENTS \
   --minMappingQuality 20  \
   --minBaseQuality 10 \
   -ct 30  \
   --omitLocusTable \
   -dt NONE \
   -U ALLOW_N_CIGAR_READS \
   -rf ReassignOneMappingQuality -RMQF 255 -RMQT 60 \
   -rf MappingQualityUnavailable



#run coverageCalculatorPy

source /home/transfer/miniconda3/bin/activate CoverageCalculatorPy

sed 's/:/\t/g'  "$sampleId"_rmdup_DepthOfCoverage | grep -v 'Locus' | sort -k1,1 -k2,2n | bgzip > "$sampleId"_rmdup_DepthOfCoverage.gz

tabix -b 2 -e 2 -s 1 "$sampleId"_rmdup_DepthOfCoverage.gz



python /home/transfer/pipelines/CoverageCalculatorPy/CoverageCalculatorPy.py \
-B RNAFusion-ROI.bed \
-D /data/results/"$seqId"/RocheSTFusion/"$sampleId"/"$sampleId"_rmdup_DepthOfCoverage.gz \
--padding 0 \
--groupfile RNA_fusion_group_file.txt \
--outname "$sampleId"_rmdup_coverage

source "$conda_bin_path"/deactivate



###########
#Run RMATS#
###########

source ~/miniconda3/bin/activate rmats

# source variables
. RMATS.config
OUTPUT_DIR=/data/results/$seqId/$panel/$sampleId/

# rmats output here
if [ -d "$OUTPUT_DIR"/RMATS ]; then
  rm -r "$OUTPUT_DIR"/RMATS
fi

mkdir -p "$OUTPUT_DIR"/RMATS

# create text file with query sample bam location - required by RMATS
echo /data/results/"$seqId"/"$panel"/"$sampleId"/"$sampleId"_Aligned_sorted.bam > "$OUTPUT_DIR"/RMATS/query_sample.txt

# rmats calling
rmats.py \
      --b1 "$OUTPUT_DIR"/RMATS/query_sample.txt \
      --b2 $RMATS_REF_SAMPLES \
      -t paired \
      --gtf $GTF_PATH \
      --variable-read-length \
      --od "$OUTPUT_DIR"/RMATS \
      --tmp "$OUTPUT_DIR"/RMATS \
      --readLength $READ_LENGTH \
      --nthread $THREADS \
      --tstat $THREADS


#########################
# GENERATE RMATS REPORT #
#########################

# ENSEMBL GENE ID
# GENE SYMBOL
# CHROMOSOME
# START OF EVENT
# END OF EVENT
# IJC_SAMPLE = NUMBER OF JUNCTION READS SUPPORTING EXON (I)NCLUSION IN SAMPLE
# SJC_SAMPLE = NUMBER OF JUNCTION READS SUPPORTING EXON (S)KIPPING  IN SAMPLE
# IJC_REF = NUMBER OF JUNCTION READS SUPPORTING EXON (I)NCLUSION IN POOLED REF
# SJC_REF = NUMBER OF JUNCTION READS SUPPORTING EXON (S)KIPPING  IN POOLED REF
# FDR = P-Value (FALSE DISCOVERY RATE)
# INC_LEVEL_SAMPLE = PROPORTION OF READS SUPPORTING EXON INCLUSION IN SAMPLE (X)
# INC_LEV_REF = PROPORTION OF READS SUPPORTING EXON INCLUSION IN REF (Y)
# INC_LEV_DIFF = mean(X) - mean(Y)"


# remove legacy report file
if [ -f "$OUTPUT_DIR"/"$seqId"_"$sampleId"_RMATS_Report.tsv ]; then
  rm  "$OUTPUT_DIR"/"$seqId"_"$sampleId"_RMATS_Report.tsv
fi


if [ -f "$OUTPUT_DIR"/RMATS/SE.MATS.JC.txt ]; then

  # if no fusion are called assume error with RMATS
  n_fusions=$(cat "$OUTPUT_DIR"/RMATS/SE.MATS.JC.txt | wc -l)
  echo "$n_fusions"
  if [[ "$n_fusions" -lt 3 ]]; then
    echo "EXITING!!"
    exit 1
  fi


  # read RMATS exon skipping report, extract MET and EGFR events
  while read ln; do

    if [[ $ln == ID* ]]; then

      header=$(echo $ln | cut -f 2-4,6-7,13-16,20-23)
      echo $header "Sample1_Perc_SJC" | sed -e 's/ /\t/g' > "$OUTPUT_DIR"/"$seqId"_"$sampleId"_RMATS_Report.tsv 

    elif [[ $ln =~ "chr7	+	116411902	116412043" ]] || [[ $ln =~ "chr7	+	55209978	55221845" ]]; then

      main=$(echo $ln | cut -f 2-4,6-7,13-16,20-23)

      IJC=$(echo $ln | cut -d " " -f 13)
      SJC=$(echo $ln | cut -d " " -f 14)
      
      # calculate proportion metric (as requested by HR)
      PROP=$(awk "BEGIN {print "$SJC"/("$IJC"+"$SJC")*100}")

      echo -e $main $PROP | sed -e 's/ /\t/g' >> "$OUTPUT_DIR"/"$seqId"_"$sampleId"_RMATS_Report.tsv
    fi

  done < "$OUTPUT_DIR"/RMATS/SE.MATS.JC.txt

fi

source ~/miniconda3/bin/deactivate




#######################################
#Separate reports into referral types #
#######################################

mkdir ./Results
mkdir ./Results/arriba
mkdir ./Results/arriba_discarded
mkdir ./Results/STAR_Fusion

source "$conda_bin_path"/activate SomaticFusion

python fusion_report_referrals.py /data/results/"$seqId"/RocheSTFusion/"$sampleId"/ $sampleId

source "$conda_bin_path"/deactivate



################################
#Append sampleId to sample list#
################################


if [ ! -e /data/results/"$seqId"/"$panel"/samples_list.txt ]
then
    echo "sampleId" >> /data/results/"$seqId"/RocheSTFusion/samples_list.txt
fi



if [ -e /data/results/"$seqId"/"$panel"/*/Results/arriba_discarded/"$sampleId"_fusion_report_NTRK3_arriba_discarded.txt ]
then
    echo $sampleId >> /data/results/"$seqId"/RocheSTFusion/samples_list.txt
fi



#################################################################
#create total_reads_list, contamination_list and analysis sheets#
#################################################################

cd /data/results/"$seqId"/RocheSTFusion

if [ ! -e /data/results/"$seqId"/"$panel"/samples_correct_order.txt ]
then
    

#create samples_list in the same order as the samplesheet for the contamination check
scp /data/archive/fastq/"$seqId"/SampleSheet.csv .

grep -A 100 "^Sample_ID" SampleSheet.csv >> samples.csv

awk -F',' '{print $1}' samples.csv >>samples_correct_order.txt

fi

grep -v "sampleId" samples_list.txt > samples_list_without_header.txt


#if the sample is the final one on the run- calculate the total reads, contamination, quality and create the analysis sheets

expected=$(cat samples_correct_order.txt| wc -l)

complete=$(cat samples.csv| uniq | wc -l)

echo $complete
echo $expected

if [ "$complete" -eq "$expected" ]; then

    #combine the total reads and total aligned reads information for all samples on the run
    python /data/diagnostics/pipelines/SomaticFusion/SomaticFusion-$version/total_reads_list.py $seqId


    #calculate contamination for run
    python /data/diagnostics/pipelines/SomaticFusion/SomaticFusion-$version/contamination_check_arriba.py $version
    python /data/diagnostics/pipelines/SomaticFusion/SomaticFusion-$version/contamination_check_star_fusion.py $version
    python /data/diagnostics/pipelines/SomaticFusion/SomaticFusion-$version/contamination_check_RMATS.py $seqId $version
    python /data/diagnostics/pipelines/SomaticFusion/SomaticFusion-$version/combine_contamination_files.py


    #Create combinedQC file 
    bash /data/diagnostics/pipelines/SomaticFusion/SomaticFusion-$version/compileQcReport.sh $seqId $panel


    cat samples_list_without_header.txt | while read sample; do

    cd /data/results/"$seqId"/RocheSTFusion/$sample

    . *.variables


    #create analysis spreadsheets

    source /home/transfer/miniconda3/bin/activate VirtualHood

    if [[ "$sample" != *"NTC"* ]]; then

        NTC_variable=NTC_"$worklistId"

        python make_worksheets.py $seqId $sample $referral "$NTC_variable" "$worklistId"
    fi
done
fi

