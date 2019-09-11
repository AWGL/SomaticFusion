cp ../make-fusion-report.py ./SampleB/

cd ./SampleB

python make-fusion-report.py \
    --sampleId SampleB \
    --seqId 190401_M02641_0198_000000000-CCCFJ

rm make-fusion-report.py

cd ../
