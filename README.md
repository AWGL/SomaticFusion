# SomaticFusion

##IMPORTANT
### Updating version on the cluster:
When updating the pipeline version on the cluster, the rmats folder should be copied from the previous version into the new version of the pipeline. This is not on github due to it containing sensitive information but is required for the pipeline to run without errors.


## Description:

Somatic pipeline used to call intergenic fusions using STAR-Fusion and arriba, and exon skipping events using RMATS.



## Running the pipeline:

In the sample directory run:

qsub 1_SomaticFusion.sh


## Requirements
  
### Samplesheet

There may be 2 worksheets on each run, with each run containing up to 48 samples, and a maximum of 24 samples per worksheet. 

Each worksheet will have its own NTC which needs to be in the format NTC-{worksheetid}
  
  
### Sample directory

Each sample directory must contain:
- A variables file in the format {sampleId}.variables
- Zipped fastq files

