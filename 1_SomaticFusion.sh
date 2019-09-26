#!/bin/bash
# set -euo pipefail
#PBS -l walltime=12:00:00
#PBS -l ncpus=12

PBS_O_WORKDIR=(`echo $PBS_O_WORKDIR | sed "s/^\/state\/partition1//" `)
cd $PBS_O_WORKDIR

# Description: Wrapper script for calling fusion genes in RNA-seq data \
# from clinical samples using STAR-Fusion.
# Arthor: Christopher Medway <christopher.medway@wales.nhs.uk>
# Date: 6th June 2019
# Useage: qsub run_star-fusion.sh [inside sample dir with .variables and \
# .fastq.gz files]

version=0.0.1

# source variables file
. *.variables

# copy the panel & pipeline variables locally and source
cp /data/diagnostics/pipelines/SomaticFusion/SomaticFusion-$version/$panel/$panel.variables . 
cp /data/diagnostics/pipelines/SomaticFusion/SomaticFusion-$version/SomaticFusion.config .
cp /data/diagnostics/pipelines/SomaticFusion/SomaticFusion-$version/make-fusion-report.py .

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

#convert FASTQ to uBAM & add RGIDs
for fastqPair in $(ls "$sampleId"_S*.fastq.gz | cut -d_ -f1-3 | sort | uniq); do

    #parse fastq filenames
    laneId=$(echo "$fastqPair" | cut -d_ -f3)
    read1Fastq=$(ls "$fastqPair"_R1_*fastq.gz)
    read2Fastq=$(ls "$fastqPair"_R2_*fastq.gz)

    #trim adapters
    cutadapt \
    -a "$read1Adapter" \
    -A "$read2Adapter" \
    -m 35 \
    --cores $ncpus \
    -o "$seqId"_"$sampleId"_"$laneId"_R1.fastq \
    -p "$seqId"_"$sampleId"_"$laneId"_R2.fastq \
    "$read1Fastq" \
    "$read2Fastq"

    #fastqc
    mkdir -p FASTQC

    fastqc -d /state/partition1/tmpdir --threads 12 --extract "$seqId"_"$sampleId"_"$laneId"_R1.fastq -o FASTQC
    fastqc -d /state/partition1/tmpdir --threads 12 --extract "$seqId"_"$sampleId"_"$laneId"_R2.fastq -o FASTQC

    #check FASTQC output
    if [ $(countQCFlagFails "$seqId"_"$sampleId"_"$laneId"_R1_fastqc.txt) -gt 0 ] || [ $(countQCFlagFails "$seqId"_"$sampleId"_"$laneId"_R2_fastqc.txt) -gt 0 ]; then
        rawSequenceQuality=FAIL
    fi
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
