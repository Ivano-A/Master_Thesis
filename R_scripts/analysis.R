###_________CITATION_________####
# Lu, Yunmei. 2024. “Examining the Stability and Change in Age-Crime Relation in South Korea, 1980–2019: 
# An Age-Period-Cohort Analysis.” PLOS ONE 19(3):e0299852. doi: 10.1371/journal.pone.0299852.

# The code was taken from the GitHub linked in the article cited above and modified in part for my needs. 


library(here)
library(readr)
library(foreign)
library(purrr)
library(MESS)
library(data.table)
library(readxl)
library(gtsummary)
library(ggplot2)
library(gnm)
library(lme4)
library(MASS)
library(Matrix)
library(gridExtra)
library(ggthemes)
library(haven)
library(tidyverse)
library(writexl)


# Set the project directory to the current working directory.
projDir <- here::here() # File path to this project's directory
dataDir <- file.path(projDir, "Data")  # File path to where data will be downloaded
outDir <- "output_xlsx"           # Name of the sub-folder to save results
figDir <- file.path(outDir, "figs")   # Name of the sub-folder to save generated figures

# Create sub-directory folders if they don't exist
if (!dir.exists(here::here(outDir))) {
  dir.create(outDir)
} else {
  print("Output directory already exists!")
}

if (!dir.exists(here::here(figDir))) {
  dir.create(figDir)
} else {
  print("Figure directory already exists!")
}


file_paths <- c("Data/Analysis_data/female_AB.xlsx",
                "Data/Analysis_data/female_atlantic.xlsx",
                "Data/Analysis_data/female_BC.xlsx",
                "Data/Analysis_data/female_MB.xlsx",
                "Data/Analysis_data/female_ON.xlsx",
                "Data/Analysis_data/female_SK.xlsx",
                "Data/Analysis_data/female_QC.xlsx",
                "Data/Analysis_data/male_AB.xlsx",
                "Data/Analysis_data/male_atlantic.xlsx",
                "Data/Analysis_data/male_BC.xlsx",
                "Data/Analysis_data/male_MB.xlsx",
                "Data/Analysis_data/male_ON.xlsx",
                "Data/Analysis_data/male_QC.xlsx",
                "Data/Analysis_data/male_SK.xlsx")

# Step 1: Global Deviance Test ####
# Initialise a list
test_results <- list()

# For loop to run the two poisson regression and anova test
for (file_path in file_paths) {
  
  dt <- read_xlsx(file_path)
  
  dt$age.f <- factor(dt$age.n)
  dt$period.f <- factor(dt$yr.n)
  dt$pop <- dt$tot_p
  
  apci <- glm(count ~ age.f * period.f + offset(log(pop)), family = "poisson", data = dt)
  
  ap <- glm(count ~ age.f + period.f, family = "poisson", data = dt)
  
  anova_results <- anova(ap, apci, test = "LRT")
  
  test_results[[file_path]] <- anova_results 
}

# For loop to save the anova test results
for (item_name in names(test_results)) {
  
  base_name <- str_remove(str_remove(item_name, "\\.xlsx$"), "Data/Analysis_data")
  
  df <- as.data.frame(test_results[[item_name]])
  
  save_selected_table <- df
  
  file_name <- file.path(outDir, paste0("global_deviance_test", base_name, "_global_dev.xlsx"))
  
  write_xlsx(save_selected_table, path = file_name)
}

# Beginning of analysis ####

# Initialize empty lists to store results
table1_list <- list()
table2a_list <- list()
table2b_list <- list()
table3a_list <- list()
table3b_list <- list()
cohort_plot_l <- list()
age_plot_l <- list()
period_plot_l <- list()
final_dataset <- list()
model_list <- list()
age.main.coh.per.constant <- list()
cohort_with_confidence <- list()
age_main_with_confidence <- list()
period_main_with_confidence <- list()

# Define sequences for period, age, and cohort to use throughout the code
period_years <- seq(1952, 2017, by = 5)
age_age <- seq(17, 82, by = 5)
cohort <- seq(1870, 2000, by = 5)

# Loop through each file path and perform analysis
for (file_path in file_paths) {
  
  # Read the xlsx file
  dt <- read_xlsx(file_path)
  
  # Ensure dt has the required columns: age.n, yr.n, tot_p, rate, cohort
  dt$age.f <- factor(dt$age.n)
  dt$period.f <- factor(dt$yr.n)
  dt$pop <- dt$tot_p
  
  a.n <- length(levels(dt$age.f))  # number of age groups
  p.n <- length(levels(dt$period.f))  # number of periods
  c.n <- a.n + p.n - 1  # number of cohorts
  
  table1 <- data.frame(
    age.group = rep(c("15-19", "20-24", "25-29", "30-34", 
                      "35-39", "40-44", "45-49", "50-54",
                      "55-59", "60-64", "65-69", "70-74",
                      "75-79", "80-84"), each = 2),
    y1950 = rep(NA, a.n * 2),
    y1955 = rep(NA, a.n * 2),
    y1960 = rep(NA, a.n * 2),
    y1965 = rep(NA, a.n * 2),
    y1970 = rep(NA, a.n * 2),
    y1975 = rep(NA, a.n * 2),
    y1980 = rep(NA, a.n * 2),
    y1985 = rep(NA, a.n * 2),
    y1990 = rep(NA, a.n * 2),
    y1995 = rep(NA, a.n * 2),
    y2000 = rep(NA, a.n * 2),
    y2005 = rep(NA, a.n * 2),
    y2010 = rep(NA, a.n * 2),
    y2015 = rep(NA, a.n * 2))
  
  
  # Defining the sequences for odd and even indices (28 is for the number of age rows * 2, one row for rate and one for cohort category)
  odd <- seq(1, 28, 2)
  even <- seq(2, 28, 2)
  
  # Looping through each column in p.n, adding the rates and cohorts to their respective columns and rows
  s <-seq(1, p.n)
  for (i in s) {
    table1[odd, i + 1] <- dt$rate[(1 + a.n * (i - 1)):(a.n * i)]
    table1[even, i + 1] <- dt$cohort[(1 + a.n * (i - 1)):(a.n * i)]
  }
  
  # ------------------------------------------------------------------- #
  
  # Step 2. Poisson APC-I model, offset term to control for population ####
  
  # Sum to zero constraint
  options(contrasts = c("contr.sum", "contr.sum"), 
          na.action = na.omit)
  
  api <- glm(count ~ age.f * period.f + offset(log(pop)), family = "poisson", data = dt)
  summary(api)
  
  api_deviance <- api$deviance
  api_null_deviance <- api$null.deviance
  api_aic <- api$aic
  
  api_model <- cbind(api_deviance, api_null_deviance, api_aic) %>% 
    as.data.frame()
  
  # Cohort matrix
  cindex <- array(rep(0, a.n * p.n), dim = c(a.n, p.n))
  for (j in 1:p.n) {
    cindex[, j] <- seq((a.n + j - 1), j, - 1)
  }
  
  # Step 3: Get Age, Period main effect ####
  co2 <- api$coefficients[1:c.n] # coefficients
  se2 <- summary(api)$coef[, 2]  # standard error
  a.co2 <- c(co2[2:a.n], -sum(co2[2:a.n]))  # age coefficients, all ages sum to 0
  p.co2 <- c(co2[(a.n + 1):c.n], -sum(co2[(a.n + 1):c.n]))  # period coefficients, all periods sum to 0
  
  # Age effect SE
  A <- contr.sum(a.n, contrasts = TRUE)
  a.vcov <- A %*% vcov(api)[2:a.n, 2:a.n] %*% t(A)
  a.se <- sqrt(diag(a.vcov))
  a.up <- a.co2 + 2 * a.se  # 95% confidence interval
  a.lo <- a.co2 - 2 * a.se  # 95% confidence interval
  
  # Period effect SE
  P <- contr.sum(p.n, contrasts = TRUE)
  p.vcov <- P %*% vcov(api)[(a.n + 1):c.n, (a.n + 1):c.n] %*% t(P)
  p.se <- sqrt(diag(p.vcov))
  p.up <- p.co2 + 2 * p.se  # 95% confidence interval
  p.lo <- p.co2 - 2 * p.se  # 95% confidence interval
  
  # Combine APC mixed model information
  api2.coef <- c(co2[1], a.co2, p.co2)  # age, period, main effects
  api2.se <- c(se2[1], a.se, p.se)  # age, period, main effect se
  cbind(api2.coef, api2.se)
  api2.aci <- cbind(a.up, a.lo)  # confidence interval of age
  api2.pci <- cbind(p.up, p.lo)  # confidence interval of period
  api2.aci
  # For Table 2a. age period main effects
  # Get age, period, cohort effects ready
  age.c4 <- as.numeric(a.co2)
  period.c4 <- as.numeric(p.co2)
  age.se4 <- a.se
  period.se4 <- p.se
  
  # Intercept
  intercept <- as.numeric(api$coefficients[1])
  in.se <- summary(api)$coefficients[1, 2]
  in.p <- 2 * pnorm(-abs(intercept / in.se))  # p-values for constant
  in.sig <- NA
  if (in.p < 0.05 & in.p >= 0.01) {
    in.sig  <- "*"
  } else {
    if (in.p < 0.01 & in.p >= 0.001) {
      in.sig  <- "**"
    } else {
      if (in.p < 0.001) {
        in.sig  <- "***"
      } else {
        in.sig  <- ""
      }
    }
  }
  # Age effect
  a.z <- age.c4 / age.se4
  a.p <- 2 * pnorm(-abs(as.numeric(a.z)))  # p-values for age effects
  a.sig <- rep(NA, a.n)
  for (i in 1:a.n) {
    pp <- a.p[i]
    if (pp < 0.05 & pp >= 0.01) {
      a.sig[i]  <- "*"
    } else {
      if (pp < 0.01 & pp >= 0.001) {
        a.sig[i]  <- "**"
      } else {
        if (pp < 0.001) {
          a.sig[i]  <- "***"
        } else {
          a.sig[i]  <- ""
        }
      }
    }
  }
  
  # Period effect
  p.z <- period.c4 / period.se4
  p.p <- 2 * pnorm(-abs(as.numeric(p.z)))  # p-values for period effects
  p.sig <- rep(NA, p.n)
  for (i in 1:p.n) {
    pp <- p.p[i]
    if (pp < 0.05 & pp >= 0.01) {
      p.sig[i]  <- "*"
    } else {
      if (pp < 0.01 & pp >= 0.001) {
        p.sig[i]  <- "**"
      } else {
        if (pp < 0.001) {
          p.sig[i]  <- "***"
        } else {
          p.sig[i]  <- ""
        }
      }
    }
  }
  
  
  #combine table for age and period main effects: 
  table2a <- data.frame(effect = c(intercept,age.c4, period.c4), 
                        se = c(in.se,age.se4, period.se4), 
                        sig = c(in.sig,a.sig, p.sig))
  
  table2a$name <- c("Intercept", names(table(dt$age)), names(table(dt$period.f)))
  
  
  #-----------------------------------------------------------------------#
  
  #-----------------------------------------------------------------------#
  # Step 4. Cohort Interaction Terms####
  # Computing full interaction index#
  
  T <- array(rep(0, a.n * p.n * (a.n - 1) * (p.n-1)), dim = c(a.n * p.n, (a.n - 1) * (p.n - 1)))
  ind1 <- a.n * 1:(p.n - 1)
  ind2 <- (a.n*(p.n - 1) + 1):(a.n * p.n - 1)
  ind3 <- a.n * p.n
  ind <- c(ind1, ind2, ind3)       
  newind <- 1:(a.n * p.n)
  newind <- newind[-ind]
  T[newind,] <- diag((a.n - 1) * (p.n - 1))
  T[ind1,] <- -diag(p.n - 1)[, rep(1:(p.n - 1), each  = a.n - 1)]
  T[ind2,] <- -diag(a.n - 1)[, rep(1:(a.n - 1), p.n - 1)]
  T[ind3,] <- -rep(1, (a.n - 1) * (p.n - 1))
  
  # Computing full interaction contrast 
  api.c <- api$coefficients
  iatemp <- vcov(api)[(a.n + p.n):length(api.c),(a.n + p.n):length(api.c)]
  iavcov <- T%*%iatemp%*%t(T)
  iaesti <- as.vector(T%*%api.c[(a.n + p.n):length(api.c)]) #coefficient
  iase <- sqrt(diag(iavcov)) 
  z <- iaesti / iase
  pvalue <- 2 * pnorm(-abs(z))
  
  cohortindex <- as.vector(cindex)
  sig <- rep(NA, a.n * p.n)
  
  for (i in 1:(a.n * p.n)){
    
    pp <- pvalue[i]
    
    if (pp < 0.05 & pp >= 0.01){ sig[i] <- "*"} else{
      
      if (pp < 0.01 & pp >= 0.001){ sig[i] <- "**"} else{
        
        if (pp < 0.001){ sig[i] <- "***"} else{
          sig[i] <- "" }
      }
    }
  }
  
  ia <- data.frame(cohort = cohortindex, 
                   b = iaesti, 
                   se = iase, z, 
                   p = pvalue, 
                   sig = sig) 

  #--------------------------#
  #--------------------------#
  #--------------------------#
  #--------------------------#
  #--------------------------#
  # Creating table to store interaction terms
  table2b <- data.frame(age.group = c("15-19", "20-24", "25-29", "30-34", 
                                      "35-39", "40-44", "45-49", "50-54",
                                      "55-59", "60-64", "65-69", "70-74",
                                      "75-79", "80-84"),
                        c1 = rep(NA,a.n), p1 = rep(NA,a.n), c2 = rep(NA,a.n), p2 = rep(NA,a.n),
                        c3 = rep(NA,a.n), p3 = rep(NA,a.n), c4 = rep(NA,a.n), p4 = rep(NA,a.n),
                        c5 = rep(NA,a.n), p5 = rep(NA,a.n), c6 = rep(NA,a.n), p6 = rep(NA,a.n),
                        c7 = rep(NA,a.n), p7 = rep(NA,a.n), c8 = rep(NA,a.n), p8 = rep(NA,a.n),
                        c9 = rep(NA,a.n), p9 = rep(NA,a.n), c10 = rep(NA,a.n), p10 = rep(NA,a.n),
                        c11 = rep(NA,a.n), p11 = rep(NA,a.n), c12 = rep(NA,a.n), p12 = rep(NA,a.n),
                        c13 = rep(NA,a.n), p13 = rep(NA,a.n), c14 = rep(NA,a.n), p14 = rep(NA,a.n)
  )
  s <- seq(2, p.n * 2, by = 2)
  
  for (i in s){
    table2b[,i] <- iaesti[(1 + (i / 2 - 1) * a.n):(i / 2 * a.n)]
    table2b[,i + 1] <- ia$sig[(1 + (i / 2 - 1) * a.n):(i / 2 * a.n)]
  }
  
  #-------------------------------------------------------------------------------#
  #4.1.inter-cohort deviations
  cint <- rep(NA, c.n)
  cintse <- rep(NA, c.n)
  cintz <- rep(NA, c.n)
  cintp <- rep(NA, c.n)
  
  for (k in 1:c.n){
    O <- sum(cindex == k)
    k1 <- rep(1 / O, O)
    k2 <- rep(0, a.n * p.n)
    k2[cindex == k] <- k1
    
    contresti <- k2 %*% iaesti
    contrse <- sqrt(t(k2) %*% iavcov %*% k2)
    z <- contresti / contrse
    if(z>0){
      p <- 2*pnorm(z, lower.tail = F)
    }else{
      p <- 2*pnorm(z, lower.tail = T)
    }
    
    cint[k] <- contresti
    cintse[k] <- contrse
    cintz[k] <- z
    cintp[k] <- p
  }
  
  cgroup <- seq(1,c.n)
  sig <- rep(NA,c.n)
  for (i in 1:c.n){
    
    pp <- cintp[i]
    
    if (pp<0.05 & pp >= 0.01){sig[i] <- "*"} else{
      
      if (pp<0.01 & pp >= 0.001){sig[i] <- "**"} else{
        
        if (pp<0.001){sig[i] <- "***"} else{
          sig[i] <- ""}
      }
    }
  }
  
  #Table 3a Inter-cohort deviation
  inter.cohort <- data.frame(cohort = c(1:c.n), 
                             deviation = cint, 
                             s.e = cintse,
                             z = cintz, 
                             p = cintp, 
                             sig = sig) 
  
  #---------------------------------------------------------------------#
  #4.2. intra-cohort: cohort slope variation
  cslope <- rep(NA, c.n)
  cslopese <- rep(NA, c.n)
  cslopez <- rep(NA, c.n)
  cslopep <- rep(NA, c.n)
  
  poly <- 1
  
  for(k in (poly + 1):(c.n - poly)){
    
    O <- sum(cindex == k)
    k1 <- contr.poly(O)
    k2 <- rep(0, a.n * p.n)
    k2[cindex == k] <- k1[,poly]
    contresti <- k2 %*% iaesti
    contrse <- sqrt(t(k2) %*% iavcov %*% k2)
    z <- contresti / contrse
    
    if(z > 0){
      
      p <- 2*pnorm(z, lower.tail = F)
      
    } else{
      
      p <- 2*pnorm(z, lower.tail = T)
      
    }
    
    cslope[k] <- contresti
    cslopese[k] <- contrse
    cslopez[k] <- z
    cslopep[k] <- p
  }
  
  sig <- rep(NA, c.n)
  for (i in (poly + 1):(c.n - poly)){
    
    pp <- cslopep[i]
    
    if (pp < 0.05 & pp >= 0.01){sig[i] <- "*"} else{
      
      if (pp < 0.01 & pp >= 0.001){sig[i] <- "**"} else{
        
        if (pp < 0.001){sig[i] <- "***"} else{
          sig[i] <- ""}
      }
    }
  }
  #table 3b. Intra cohort slope table
  cgroup <- seq(1 , c.n)
  
  table3b <- data.frame(cohort = cgroup, 
                        slope = cslope, 
                        s.e = cslopese, 
                        z = cslopez, 
                        p = cslopep,
                        sig. = sig)
  
  #Figure 1 Age and Period main effects - Get data ready ####
  
  # Defining the age effects data frame
  a_df <- data.frame(
    effect= age.c4,
    age = age_age,
    ci_lo = age.c4 - 2 * age.se4,
    ci_up = age.c4 + 2 * age.se4
  )
  
  # Defining the period effects data frame
  p_df <- data.frame(
    effect = period.c4,
    year = period_years,
    ci_lo = period.c4 - 2 * period.se4,
    ci_up = period.c4 + 2 * period.se4
  )
  
  #cohort effects
  table3b$slope <- round(table3b$slope, digits = 3)
  cohort.sig <- NULL
  
  for (i in 1:c.n){
    cohort.sig[i] <- paste(table3b$slope[i],table3b$sig.[i], sep = "")
  }
  cohort.c4 <- as.numeric(cint)
  cohort.se4 <- as.numeric(cintse)
  
  cohort <- seq(1870, 2000, by = 5)
  
  # Cohort effects dataframe
  c_df<-data.frame(effect = cohort.c4,
                   cohort.n = cohort,
                   ci_lo = cohort.c4-2*cohort.se4,
                   ci_up = cohort.c4+2*cohort.se4,
                   intra = round(table3b$slope,digits = 3),
                   intra.sig = cohort.sig)

  
  #----------------------------------------------------------------#
  
  
  # Matrix for cohort interaction terms
  f2.d <- as.matrix(subset(table2b, select = c(c1, c2, c3, c4, 
                                               c5, c6, c7, c8,
                                               c9, c10, c11, c12,
                                               c13, c14)))
  
  
  # Variables for predicted rate plots #####
  
  # Initialize dt$pco and dt$aco with NA values
  dt$pco <- rep(NA, nrow(dt))
  dt$aco <- rep(NA, nrow(dt))
  
  # Looping over period_years and age_age
  for (i in 1:p.n) {
    j  <- period_years[i]
    dt$pco[dt$yr.n == j] <- p.co2[i]
  }
  p.co2
  for (i in 1:a.n) {
    j <- age_age[i]
    dt$aco[dt$age.n == j] <- a.co2[i]
  }
  
  # Arrange dt to properly assign interaction terms
  dt <- dt %>% 
    arrange(yr.n)
  
  # Assign cohort interaction terms to each observation
  dt$cco  <- c(f2.d[, 1], f2.d[, 2], f2.d[, 3], f2.d[, 4], 
               f2.d[, 5], f2.d[, 6], f2.d[, 7], f2.d[, 8],
               f2.d[, 9], f2.d[, 10], f2.d[, 11], f2.d[, 12],
               f2.d[, 13], f2.d[, 14])
  
  dt <- dt %>% 
    arrange(age.n)
  
  # Calculating predicted suicide rates and storing them in dt ####
  
  
  #### Plots are produced in separate R file ###
  
  
  #age only
  dt$p1 <- exp(intercept + dt$aco) / (1 + exp(intercept + dt$aco)) * 100000
  #with age+period
  dt$p2 <- exp(intercept + dt$aco + dt$pco) / (1 + exp(intercept + dt$aco + dt$pco)) * 100000
  #with all effects prediction
  dt$p3 <- exp(intercept + dt$aco + dt$pco + dt$cco) / (1 + exp(intercept + dt$aco + dt$pco + dt$cco)) * 100000
  # period only
  dt$p4 <- exp(intercept + dt$pco)/ (1 + exp(intercept + dt$pco)) * 100000
  
  
  ### ### ### ### ### ### ### ###
  
  # Age only model ***NOT USED***
  aonly <- glm(count ~ age.f + offset(log(pop)), 
               family = "poisson", data = dt)
  
  a.in <- coefficients(aonly)[1]
  co2 <- aonly$coefficients[1:a.n] # Coefficients
  se2 <- summary(api)$coef[,2] # Standard error
  a.co2 <- c(co2[2:a.n], -sum(co2[2:a.n])) # Age coefficients, all ages sum to 0
  
  # Assign age coefficients to each age
  for (i in 1:a.n){
    j <- age_age[i]
    dt$a.aco[dt$age.f == j] <- a.co2[i]
  }
  
  dt$a.p2 <- exp(a.in + dt$a.aco) / (1 + exp(a.in + dt$a.aco)) * 100000
  
  dt$ac.p2 <- exp(intercept + dt$aco + dt$cco) / (1 + exp(intercept + dt$aco + dt$cco)) * 100000
  
  ### ### ### ### ### ### ### ###
  
  #table 1 Suicide rates with cohorts ####
  # Round the values of tab1 to one decimal point
  round_numeric_columns <- function(table1) {
    table1[] <- lapply(table1, function(x) if (is.numeric(x)) round(x, 1) else x)
    return(table1)
  }
  
  # Apply the function and store the results in tab1
  tab1_data_structure <- round_numeric_columns(table1)
  
  # Figure 1 estimated age and period effects, holding period and cohort constant ####
  
  a_df1 <- as.data.frame(a_df) 
  
  age_plot <- ggplot(a_df1, aes(x = age_age, y = effect)) + 
    geom_errorbar(aes(ymin = ci_lo, ymax = ci_up), width = 1) + 
    scale_x_continuous(breaks = seq(17,82, by = 5),
                       labels = c("15-19", "20-24", "25-29", "30-34", 
                                  "35-39", "40-44", "45-49", "50-54",
                                  "55-59", "60-64", "65-69", "70-74",
                                  "75-79", "80-84")) +
    geom_point() +
    geom_line()+
    geom_hline(yintercept = 0)+
    xlab("Age") + ylab("Age Effect") +
    ggtitle("Estimated Age Effects with Period and Cohort Effects Held Constant") +
    theme_minimal()
  
  # Period main effects plot
  p1 <- data.frame(p_df)
  
  period_plot <- ggplot(p1, aes(x = year, y = effect)) +
    geom_point() +
    geom_line() +
    geom_hline(yintercept = 0) +
    geom_errorbar(aes(ymin = ci_lo, ymax = ci_up),width = 1) +
    scale_x_continuous(breaks = seq(1952, 2017, 5),
                       labels = c("1950-1954", "1955-1959", "1960-1964", "1965-1969", 
                                  "1970-1974", "1975-1979", "1980-1984", "1985-1989",
                                  "1990-1994", "1995-1999", "2000-2004", "2005-2009", 
                                  "2010-2014", "2015-2019")) +
    xlab("Period") + 
    ylab("Period Effect") +
    ggtitle("Estimated Period Effects with Age and Cohort Effects Held Constant")
  
  # Inter-cohort plot:
  
  c1 <- data.frame(c_df)
  
  cohort_plot <- ggplot(c1, aes(x = cohort.n, y = effect)) +
    scale_x_continuous(breaks = seq(1870, 2000, 5)) +
    geom_point() +
    geom_line()+
    geom_hline(yintercept = 0)+
    xlab("Cohort") + 
    ylab("Inter-Cohort Deviation") +
    ggtitle("Estimated Inter-Cohort Deviation Effects")
  
  # Storing other tables ####
  
  #Table 2a:Main effects
  table2a <- as.data.frame(table2a) %>% 
    relocate(name)
  
  #Table 2b: interaction terms
  table2b <- data.frame(table2b)
  
  #Table 3a: inter-cohort deviation
  table3a <- data.frame(inter.cohort) %>% 
    dplyr::mutate(
      cohort = seq(1870, 2000, by = 5)
    )
  
  
  #----------------------------#
  #----------------------------#
  #----------------------------#
  # Saving Section####
  
  # Save the results for all the output in the respective list and file_path
  results <- list(
    table1_list[[file_path]] <- tab1_data_structure, # rates with cohorts
    table2a_list[[file_path]] <- table2a, # main effects
    table2b_list[[file_path]] <- table2b, #  interaction terms
    table3a_list[[file_path]] <- table3a, # inter-cohort
    table3b_list[[file_path]] <- table3b, # intra-cohort
    cohort_plot_l[[file_path]] <- cohort_plot, # inter-cohort deviation plot
    age_plot_l[[file_path]] <- age_plot, # estimated age effects
    period_plot_l[[file_path]] <- period_plot, # estimated period effects
    final_dataset[[file_path]] <- dt, # final dataframe
    model_list[[file_path]] <- api_model, # poisson model
    age_main_with_confidence[[file_path]] <- a_df,
    period_main_with_confidence[[file_path]] <- p_df,
    cohort_with_confidence[[file_path]] <- c_df
  )
  
}

for (item_name in names(model_list)) {
  
  base_name <- str_remove(str_remove(item_name, "\\.xlsx$"), "Data/Analysis_data")
  
  save_selected_table <- model_list[[item_name]]
  
  file_name <- file.path(outDir, paste0("model_list", base_name, "_poisson_model.xlsx"))
  
  write_xlsx(save_selected_table, path = file_name)
}


for (item_name in names(table1_list)) {

  save_selected_table <- table1_list[[item_name]]

  file_name <- file.path(outDir, paste0("table1", item_name))

  write_csv(save_selected_table, file_name)
}

for (item_name in names(table2a_list)) {
  
  base_name <- str_remove(str_remove(item_name, "\\.xlsx$"), "Data/Analysis_data")
  
  save_selected_table <- table2a_list[[item_name]]
  
  file_name <- file.path(outDir, paste0("table2a", base_name, "_age_period_main_effects.xlsx"))
  
  write_xlsx(save_selected_table, path = file_name)
}


for (item_name in names(table2b_list)) {
  
  base_name <- str_remove(str_remove(item_name, "\\.xlsx$"), "Data/Analysis_data")
  
  save_selected_table <- table2b_list[[item_name]]
  
  file_name <- file.path(outDir, paste0("table2b", base_name, "_inter_coh_coef_by_age_period.xlsx"))
  
  write_xlsx(save_selected_table, path = file_name)
}

for (item_name in names(table3a_list)) {
  
  base_name <- str_remove(str_remove(item_name, "\\.xlsx$"), "Data/Analysis_data")
  
  save_selected_table <- table3a_list[[item_name]]
  
  file_name <- file.path(outDir, paste0("table3a", base_name, "_inter_cohort_deviation.xlsx"))
  
  write_xlsx(save_selected_table, path = file_name)
}

for (item_name in names(table3b_list)) {
  
  base_name <- str_remove(str_remove(item_name, "\\.xlsx$"), "Data/Analysis_data")
  
  save_selected_table <- table3b_list[[item_name]]
  
  file_name <- file.path(outDir, paste0("table3b", base_name, "_intra_cohort_slope.xlsx"))
  
  write_xlsx(save_selected_table, path = file_name)
}

for (item_name in names(age_main_with_confidence)) {
  
  base_name <- str_remove(str_remove(item_name, "\\.xlsx$"), "Data/Analysis_data")
  
  save_selected_table <- age_main_with_confidence[[item_name]]
  
  file_name <- file.path(outDir, paste0("age_CI", base_name, "_age_confidence_intervals.xlsx"))
  
  write_xlsx(save_selected_table, path = file_name)
}

for (item_name in names(cohort_with_confidence)) {
  
  base_name <- str_remove(str_remove(item_name, "\\.xlsx$"), "Data/Analysis_data")
  
  save_selected_table <- cohort_with_confidence[[item_name]]
  
  file_name <- file.path(outDir, paste0("cohort_CI", base_name, "_cohort_confidence_intervals.xlsx"))
  
  write_xlsx(save_selected_table, path = file_name)
}


for (item_name in names(final_dataset)) {
  
  df_figure3 <- final_dataset[[item_name]] %>% 
    dplyr::select(-count, -rate, -tot_p, -pop)
  
  base_name <- str_remove(str_remove(item_name, "\\.xlsx$"), "Data/Analysis_data")
  
  save_selected_table <- df_figure3
  
  file_name <- file.path(outDir, paste0("final_dataset", base_name, "_final_data.xlsx"))
  
  write_xlsx(save_selected_table, path = file_name)
}

for (item_name in names(age.main.coh.per.constant)) {
  
  base_name <- str_remove(str_remove(item_name, "\\.xlsx$"), "Data/Analysis_data")
  
  save_selected_plot <- age.main.coh.per.constant[[item_name]]
  
  file_name <- file.path(outDir, paste0("age.main.coh.per.constant", base_name, "_model_predicted_suicide_rates.png"))
  
  ggsave(file_name, plot = save_selected_plot, width = 11, height = 8.5)
}

for (item_name in names(age_plot_l)) {
  
  base_name <- str_remove(str_remove(item_name, "\\.xlsx$"), "Data/Analysis_data")
  
  save_selected_plot <- age_plot_l[[item_name]]
  
  file_name <- file.path(outDir, paste0("age_plot", base_name, "_estimated_age_effect.png"))
  
  ggsave(file_name, plot = save_selected_plot, width = 11, height = 8.5)
}

for (item_name in names(cohort_plot_l)) {
  
  base_name <- str_remove(str_remove(item_name, "\\.xlsx$"), "Data/Analysis_data")
  
  save_selected_plot <- cohort_plot_l[[item_name]]
  
  file_name <- file.path(outDir, paste0("cohort_plot_l", base_name, "_inter_cohort_deviation.png"))
  
  ggsave(file_name, plot = save_selected_plot, width = 11, height = 8.5)
}

for (item_name in names(period_plot_l)) {
  
  base_name <- str_remove(str_remove(item_name, "\\.xlsx$"), "Data/Analysis_data")
  
  save_selected_plot <- period_plot_l[[item_name]]
  
  file_name <- file.path(outDir, paste0("period_plot_l", base_name, "_estimated_period_effects.png"))
  
  ggsave(file_name, plot = save_selected_plot, width = 11, height = 8.5)
}


