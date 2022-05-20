# BSQA
Package to analyse MRI data acquired for the purpose of assessing scanners breast imaging performance.  The package is capable of producing metrics for SNR (sensitive to individual element failure), uniformity, suppression effectiveness and contrast. And publish a report with the results.  

BSQA_analysis requires a config file to be placed in the same directory and named BSQA_config.ini.  An example config file is available. Please refer to the document BSQA_analysis_code_user_guide.pdf for further details on how the code functions. 

The source code will require editing before use at different sites, in particular the algorithm for sorting the DICOM images. While every attempt has been made to ensure the program functions as expected and produces accurate results, users will need to staisfy themselves that this is the case. 
