# Master_Thesis
Master's Thesis in Sociology at Concordia University 

The R code used for the Age, Period, and Cohort - Interaction Model (APC-I) (file name: analysis.R in R_scripts) was repurposed, transformed, and adapted from the code provided in the following journal article: 
Lu, Yunmei. “Examining the Stability and Change in Age-Crime Relation in South Korea, 1980–2019: An Age-Period-Cohort Analysis.” PLOS ONE 19, no. 3 (March 29, 2024): e0299852. https://doi.org/10.1371/journal.pone.0299852.

Their GitHub account containing the code is found here: https://github.com/irisyunmeilu/SK_APC

All other code for the descriptive statistics, cleaning, and manipulation is my own, in addition the graphics, tables, and Excel spreadsheets.

People are free do as they please with the items in accordance with the License attached.

**File Guide**

The APC_results folder contains the results from the Poisson regression, predicted rates, and more -- see the variable dictionary.

The R_scripts folder contains the files for the data manipulation/analysis, separated between the APC analysis performed on the unrounded counts at the Mcgill-Concordia QICSS laboratory. There is also an R script for all figures.

The data_files folder contains the *rounded* counts that were withdrawn from the QICSS lab, and the population data used for the APC analysis and to calculate rates.

**Data Sources**

The suicide count data were acquired from official Statistics Canada data through the Canadian Vital Statistics Deaths database for years 1974-2019. The data for 1950-1973 were acquired and digitised from here: 
Expert Working Group on the Revision and Updating of the Original Task Force Report on Suicide in Canada, and Canada, eds. Suicide in Canada: Update of the Report of the Task Force on Suicide in Canada. Ottawa: Health Canada, 1994.

Population data:

Yearly population estimates, 1950-1970:
Statistics Canada. Table 17-10-0029-01  Estimates of population, by age group and sex, Canada, provinces and territories (x 1,000). https://doi.org/10.25318/1710002901-eng

Yearly population estimates, 1971-2019:
Statistics Canada. Table 17-10-0005-01  Population estimates on July 1, by age and gender. https://doi.org/10.25318/1710000501-eng

Acknowledgement for QICSS access:

The analyses contained in this text were carried out at the Centre interuniversitaire québécois de statistiques sociales (QICSS), a member of the Canadian Research Data Centre Network (CRDCN). The activities of the QICSS are made possible by the financial support of the Social Sciences and Humanities Research Council (SSHRC), the Canadian Institutes of Health Research (CIHR), the Canada Foundation for Innovation (CFI), Statistics Canada, the Fonds de recherche du Québec and all the Quebec universities that contribute to their funding. The ideas expressed in this text are those of the authors and not necessarily those of the CRDCN, the QICSS or their partners.

