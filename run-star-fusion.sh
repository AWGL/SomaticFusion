#!/bin/bash
#PBS -l walltime=12:00:00
#PBS -l ncpus=12
#PBS_O_WORKDIR=(`echo $PBS_O_WORKDIR | sed "s/^\/state\/partition1//" `)
#cd $PBS_O_WORKDIR

# Description: Wrapper script for calling fusion genes in RNA-seq data \
# from clinical samples using STAR-Fusion.
# Arthor: Christopher Medway <christopher.medway@wales.nhs.uk>
# Date: 6th June 2019
# Useage: qsub run_star-fusion.sh [inside sample dir with .variables and \
# .fastq.gz files]

version=0.0.1
conda_bin_path=/home/cm/anaconda3/bin
starfusion_lib=/share/apps/star-fusion/GRCh37_gencode_v19_CTAT_lib_Mar272019.plug-n-play/ctat_genome_lib_build_dir/
ncpus=12

. *.variables

echo $sampleId

## STAR-fusion required a samples_file in order to handle multi-lane runs
R1=$(for i in ./"$sampleId"_*.fastq.gz; do echo $i | grep "R1"; done)
R2=$(for i in ./"$sampleId"_*.fastq.gz; do echo $i | grep "R2"; done)
SAMPLE=$(for i in $R1;do echo $sampleId;done)
paste <(printf %s "$SAMPLE") <(printf %s "$R1") <(printf %s "$R2") > "$sampleId".samples

# activate star-fusion environment
source "$conda_bin_path"/activate star-fusion

STAR-Fusion --genome_lib_dir $starfusion_lib \
            --samples_file $sampleId.samples \
            --output_dir ./STAR-Fusion/ \
            --FusionInspector validate \
            --denovo_reconstruct \
            --examine_coding_effect \
            --CPU $ncpus \
            --min_FFPM 1

source "$conda_bin_path"/deactivate
