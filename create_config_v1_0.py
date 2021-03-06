# -*- coding: utf-8 -*-
"""
Created on Thu Apr 14 11:01:08 2022

@author: James.Harkin
"""

import configparser

config = configparser.ConfigParser()
config.optionxform = str

config["user_input"] = {"ask_overwrite_analysis": False}

config["default_paths"] = {
    "default_search_path" : "S:/Radiology-Directorate/IRIS/Non-Ionising/MRI/QA and Calibration/NHSBSP/DICOMS/to_analyse",
    "default_archive_path" : "S:/Radiology-Directorate/IRIS/Non-Ionising/MRI/QA and Calibration/NHSBSP/DICOMS/",
    "default_results_path" : "S:/Radiology-Directorate/IRIS/Non-Ionising/MRI/QA and Calibration/NHSBSP/Results/NHSBSP_Results.xlsx",
    "default_baseline_path" : "S:/Radiology-Directorate/IRIS/Non-Ionising/MRI/QA and Calibration/NHSBSP/Results/NHSBSP_Baseline.xlsx",
    "default_report_path" : "S:/Radiology-Directorate/IRIS/Non-Ionising/MRI/QA and Calibration/NHSBSP/Reports/",
    }
config["sequence_names"] = {"T2":"DelRec - SE_SNR","T1": "T1_GRE_NO_SPIR", "T1_sup": "T1_GRE_SPIR"}   

config["coil_spec"] = {"Elements":15, "Lower_Threshold":0.01, "Upper_Threshold":0.1}
    
config["sheet_names_to_report"] = {
    "SNR" : "GA_SNR_T2_noise_av",
    "uniformity" : "GA_Uniformity_T2",
    "suppression" : "GA_Suppression_T1w",
    "contrast" : "GA_Contrast_T1w"
    }
config["pdf_header_image"] = {"file_path":"S:/Radiology-Directorate/IRIS/Non-Ionising/MRI/QA and Calibration/NHSBSP/Reports/NIR Logo.PNG"}
config["pdf_text_sizes"] = {"title_size":24, "heading_size":16, "body_txt_size":10}
config["pdf_default_comment"] = {"comment" : "Default PASS comment -  The results from today's testing indicates that the 3.0 T Philips Elition X meets the requirements specified in the NHS Breast Cancer Screening Program Technical Guidelines for Magnetic Resonance Imaging (MRI)."}
config["pdf_glossary"] = {
    "SNR" : "(K_mean)(<TM>/<air>)(BW/BW0)^0.5. Where K_mean is a scale factor dependant on the number of elements, BW is the receive bandwidth of the acquired image, BW0 = 130 Hz/Px, <TM> is the average magnitude of pixels within the phantom mask and <air>is the average magnitude of pixels within the air mask.",
    "Uniformity": "The Uniformity (Uint) is calculated using the method outlined in IPEM Report 112 and is referred to as the integral uniformity method.",
    "CE" :  "Combined Element: These images are generated using a sum of squares combination of the individual element images.",
    "IE" : "Individual Element: IE_1 indicates the element 1.  Elements are numbered sequentially according to the DICOM Tag: 07A1,103E.",
    "Contrast" : "<TM1> - <TM2>.  Where <TM1> is the average magnitude of pixels within the TM1 mask and <TM2> is the average magnitude of pixels within the TM2 mask.",
    "Scaled Contrast" : "The Contrast divided by <TM>", 
    "Contrast Ratio" : "<TM1>/<TM2>",
    "Ratio of contrast ratios" : "The contrast ratio for the T1w suppressed image divided by the contrast ratio for the non suppressed image.",
    "Ratio of contrasts": "The contrast of the suppressed image divided by the contrast of the non suppressed image.",
    "Scaled ratio of contrasts": "The scaled contrast of the suppressed image divided by the scaled contrast of the non suppressed image."
    }
config["report_sections"] = {
    "T2W SE Image Analysis" : "This section summarises the results of SNR and uniformity testing of the T2W images. SNR is reported for each element individually, and all elements combined. Uniformity is reported for the combined element (CE) image.",
    "T1W GRE Image Analysis" : "This section summarises the results of the suppression effectiveness and contrast testing of the CE T1w images.",
    "Appendix A - T2W SE Magnitude Images" : "This section displays the T2W magnitude images, for each element individually and all elements combined; on the testing date and at acceptance."
    }



with open('BSQA_config.ini', 'w') as configfile:
    config.write(configfile)