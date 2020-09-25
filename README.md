# SomaticFusion

## Description:

Somatic pipeline used to call intergenic fusions using STAR-Fusion and arriba, and exon skipping events using RMATS.



## Running the pipeline:

In the sample directory run:

qsub 1_SomaticFusion.sh


## Requirements
  
### Samplesheet

There may be 2 worksheets on each run, with each run containing up to 48 samples, and a maximum of 24 samples per worksheet. 

Each worksheet will have its own NTC which needs to be in the format NTC-<worksheetid>
  
  
### Sample directory

Each sample directory must contain:
- A variables file in the format <sampleId>.varibles
- Zipped fastq files
