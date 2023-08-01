# Collision Report Extraction
A collision report extractor to extract Information from California DMV AV collision reports. 

The California DMV receives collision reports for Autonomous Vehicles in PDF format, making it challenging to manually compile all the data from these reports. To address this issue, an extractor has been developed to automatically extract relevant information from the collision reports and consolidate it into a single Excel file. This tool aims to simplify the process of gathering and analyzing collision data for Autonomous Vehicles.

[![DOI](https://zenodo.org/badge/658437595.svg)](https://zenodo.org/badge/latestdoi/658437595)


Author: Saquib M Haroon and Alyssa Ryan @ University of Arizona
## Libraries used
The successful implementation of the extractor relied on the utilization of specific libraries. These libraries played a crucial role in generating the final Excel file from the Autonomous Vehicle collision reports.  

easyOCR  
pdf2image  
openpyxl  
OpenCV  
NumPy 

## Extracted Dataset
Find the latest extracted dataset upto June 2023 [here](Cal_DMV_AV_Dataset_2019+.xlsx)

## Future Work

Use NLP models to automatically extract Injury information from the description.  
Geocode the Address so as to identify collision coordinates

## Lab website
Visit our Website: [Ryan Research Lab](https://www.alyssaryan.co)
