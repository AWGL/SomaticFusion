
#!/bin/bash
# set -euo pipefail
#PBS -l walltime=12:00:00
#PBS -l ncpus=12

PBS_O_WORKDIR=(`echo $PBS_O_WORKDIR | sed "s/^\/state\/partition1//" `)
cd $PBS_O_WORKDIR

# Description: Wrapper script for calling fusion genes in RNA-seq data \
# from clinical samples using STAR-Fusion.
# Authors: Christopher Medway and Laura McCluskey
# Date: 27th March 2020
# Usage: qsub run_star-fusion.sh [inside sample dir with .variables and \
# .fastq.gz files]

version=0.0.1

# source variables file
. *.variables

# copy the panel & pipeline variables locally and source
cp /data/diagnostics/pipelines/SomaticFusion/SomaticFusion-$version/$panel/$panel.variables . 
cp /data/diagnostics/pipelines/SomaticFusion/SomaticFusion-$version/SomaticFusion.config .
cp /data/diagnostics/pipelines/SomaticFusion/SomaticFusion-$version/make-fusion-report.py .
cp /data/diagnostics/pipelines/SomaticFusion/SomaticFusion-$version/RNA_fusion_group_file.txt .


vendorCaptureBed=./180702_HG19_PanCancer_EZ_capture_targets.bed
vendorPrimaryBed=./180702_HG19_PanCancer_EZ_primary_targets.bed

gatk3=/share/apps/GATK-distros/GATK_3.8.0/GenomeAnalysisTK.jar

. $panel.variables
. SomaticFusion.config

# set conda env
source "$conda_bin_path"/activate SomaticFusion


#######################
# Preprocessing FASTQ #
#######################

#count how many core FASTQC tests failed
countQCFlagFails() {
    grep -E "Basic Statistics|Per base sequence quality|Per tile sequence quality|Per sequence quality scores|Per base N content" "$1" | \
   grep -v ^PASS | \
    grep -v ^WARN | \
    wc -l | \
    sed 's/^[[:space:]]*//g'
}

#record FASTQC pass/fail
rawSequenceQuality=PASS



for fastqPair in $(ls "$sampleId"_S*.fastq.gz | cut -d_ -f1-4 | sort | uniq); do

    #parse fastq filenames
    laneId=$(echo "$fastqPair" | cut -d_ -f4)
    read1Fastq=$(ls "$fastqPair"_R1_*fastq.gz)
    read2Fastq=$(ls "$fastqPair"_R2_*fastq.gz)

    #trim adapters
    cutadapt -a AGATCGGAAGAGC -A AGATCGGAAGAGC -m 35 -o "$seqId"_"$sampleId"_"$laneId"_R1.fastq -p "$seqId"_"$sampleId"_"$laneId"_R2.fastq "$read1Fastq" "$read2Fastq"


    #fastqc -d /state/partition1/tmpdir --threads 12 --extract "$seqId"_"$sampleId"_"$laneId"_R1.fastq
    fastqc -d /state/partition1/tmpdir --threads 12 --extract "$seqId"_"$sampleId"_"$laneId"_R2.fastq

    mv "$seqId"_"$sampleId"_"$laneId"_R1_fastqc/summary.txt "$seqId"_"$sampleId"_"$laneId"_R1_fastqc.txt
    mv "$seqId"_"$sampleId"_"$laneId"_R2_fastqc/summary.txt "$seqId"_"$sampleId"_"$laneId"_R2_fastqc.txt

    #check FASTQC output
    if [ $(countQCFlagFails "$seqId"_"$sampleId"_"$laneId"_R1_fastqc.txt) -gt 0 ] || [ $(countQCFlagFails "$seqId"_"$sampleId"_"$laneId"_R2_fastqc.txt) -gt 0 ]; then
        rawSequenceQuality=FAIL
fi



 

############
# Run STAR #
############

sampleId_without_conc="$(echo $sampleId | cut -d'_' -f1)"


STAR --chimSegmentMin 12 \
     --chimJunctionOverhangMin 12 \
     --chimSegmentReadGapMax 3 \
     --alignSJDBoverhangMin 10 \
     --alignMatesGapMax 200000 \
     --chimOutType SeparateSAMold \
     --alignIntronMax 200000  \
     --alignSJstitchMismatchNmax 5 -1 5 5  \
     --twopassMode Basic \
     --readFilesIn "$seqId"_"$sampleId"_"$laneId"_R1.fastq "$seqId"_"$sampleId"_"$laneId"_R2.fastq \
     --genomeDir /share/apps/star-fusion/GRCh37_gencode_v19_CTAT_lib_Mar272019.plug-n-play/ctat_genome_lib_build_dir/ref_genome.fa.star.idx  \
     --outFileNamePrefix "$seqId"_"$sampleId"_"$laneId"_ \
     --outSAMattrRGline ID:$sampleId SM:$sampleId_without_conc
done


###################
# Run STAR-Fusion #
###################


## STAR-fusion required a samples_file in order to handle multi-lane runs
R1=$(for i in ./"$sampleId"_*.fastq.gz; do echo $i | grep "R1"; done)
R2=$(for i in ./"$sampleId"_*.fastq.gz; do echo $i | grep "R2"; done)
SAMPLE=$(for i in $R1;do echo $sampleId;done)
paste <(printf %s "$SAMPLE") <(printf %s "$R1") <(printf %s "$R2") > "$sampleId".samples

# run STAR-Fusion
STAR-Fusion --genome_lib_dir $starfusion_lib \
            --samples_file $sampleId.samples \
            --output_dir ./STAR-Fusion/ \
            --FusionInspector validate \
            --denovo_reconstruct \
            --examine_coding_effect \
            --CPU $ncpus \
            --min_FFPM 1


###########################
# Generate Fusion Reports #
###########################


python make-fusion-report.py \
    --sampleId $sampleId \
    --seqId $seqId \
    --panel $panel \
    --ip $(hostname --ip-address)

# deactivate conda env
source "$conda_bin_path"/deactivate



#################################################
# Covert sam files to bam files and merge lanes #
#################################################


for sam_file in *.Aligned.out.sam; do

    laneId=$(ls "$sam_file" | cut -d  "_" -f7)
    samtools view -S -b "$sam_file" > "$seqId"_"$sampleId"_"$laneId"_Aligned.out.bam

done


samtools merge "$sampleId"_Aligned_out.bam *Aligned.out.bam 

samtools sort "$sampleId"_Aligned_out.bam "$sampleId"_Aligned_sorted

samtools index "$sampleId"_Aligned_sorted.bam



for sam_file in *Chimeric.out.sam; do

    laneId=$(ls "$sam_file" | cut -d  "_" -f7)

    samtools view -S -b "$sam_file" > "$seqId"_"$sampleId"_"$laneId"_Chimeric.out.bam

done

samtools merge "$sampleId"_Chimeric.out.bam *Chimeric.out.bam



##############
# Run arriba #
##############

arriba -x "$sampleId"_Aligned_sorted.bam \
-g /share/apps/star-fusion/GRCh37_gencode_v19_CTAT_lib_Mar272019.plug-n-play/ctat_genome_lib_build_dir/ref_annot.gtf \
-a /share/apps/star-fusion/GRCh37_gencode_v19_CTAT_lib_Mar272019.plug-n-play/ctat_genome_lib_build_dir/ref_genome.fa \
-o "$sampleId"_fusions_adapted.tsv -b /home/transfer/miniconda3/envs/arriba/var/lib/arriba/blacklist_hg19_hs37d5_GRCh37_2018_11-04.tsv.gz \
-c "$sampleId"_Chimeric.out.bam \
-O "$sampleId"_fusions_discarded_adapted.tsv \
-R 0 \
-f intragenic,same_gene



###############################
# Calculate depth of coverage #
###############################

## DOC
/share/apps/jre-distros/jre1.8.0_131/bin/java -XX:GCTimeLimit=50 -XX:GCHeapFreeLimit=10 -Djava.io.tmpdir=/state/partition1/tmpdir -Xmx4g -jar $gatk3 \
   -T DepthOfCoverage \
   -R /share/apps/star-fusion/star-fusion-ref.fa \
   -I  "$sampleId"_Aligned_sorted.bam \
   -L ../RNAFusion-ROI_adapted.bed \
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


source /home/transfer/miniconda3/bin/activate CoverageCalculatorPy


python /home/transfer/pipelines/CoverageCalculatorPy/CoverageCalculatorPy.py \
-B RNAFusion-ROI_adapted.bed \
-D /data/results/"$seqId"/RocheSTFusion/"$sampleId"/"$seqId"_"$sampleId"_DepthOfCoverage.gz \
--padding 0 \
--groupfile RNA_fusion_group_file.txt \
--outname "$sampleId"_coverage 



#################################################
# Calculate qc metrics #
#################################################


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


