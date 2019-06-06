#!/bin/bash

sampleId=18M80495

## make samples_file for multi lane runs
R1=$(for i in "$sampleId"_*.fastq.gz; do echo $i | grep "R1"; done)
R2=$(for i in "$sampleId"_*.fastq.gz; do echo $i | grep "R2"; done)
SAMPLE=$(for i in $R1;do echo $sampleId;done)

# checj R1 and R2 are equal

paste <(printf %s "$SAMPLE") <(printf %s "$R1") <(printf %s "$R2") > "$sampleId".samples



source /home/cm/anaconda3/bin/activate star-fusion

STAR-Fusion --genome_lib_dir /home/cm/pipeline_development/panCancerFusion/GRCh37_gencode_v19_CTAT_lib_Mar272019.plug-n-play/ctat_genome_lib_build_dir \
            --samples_file /home/cm/pipeline_development/panCancerFusion/"$sampleId".samples \
            --output_dir /home/cm/pipeline_development/panCancerFusion/results/$sampleId/ \
            --FusionInspector validate

source /home/cm/anaconda3/bin/deactivate


