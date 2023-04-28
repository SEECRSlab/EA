### INITIAL PROCESSING OF EA DATA ###

# Script by Miriam Gieske
# Jan. 3, 2020

# This script will extract the columns you need from the raw EA data, 
#   put them in the correct order,
#   and extract the data for the standards

# Each time you use this script to process a new batch of data,
#   you will need to change the following:
# The file path telling R which folder the data are in
# The names of the weigh sheet and EA data files
# The name of the Excel file R saves at the end

# When switching from one batch of data to another, you should also clear the old
#   data by clicking the broom in the "Environment" window and clear the old code
#   by clicking in the "Console" window and hitting Ctrl+L.

#### LOAD PACKAGES ####

library(readxl) # Reads in .xls and .xlsx files
library(openxlsx) # Writes .xlsx files

#### READ IN THE DATA ####

# Tell R where to find the data (i.e. the file path)
# Note: If you copy and paste the file path from the top of the Windows
#     file explorer window, you will need to change the backslashes ("\")
#     to forward slashes ("/").
setwd("/Volumes/GoogleDrive/My Drive/Tillage/Soil Data")

# Read in the weigh sheet
weigh_sheet <- read_excel("Tillage 2.0 Post Soil CN Weigh Sheet.xlsx", skip = 5,
                          col_names = c("Analysis order", "Tray position", 
                                        "Sample ID", "description", 
                                        "Target wt (mg)", "Actual wt (mg)"))
             
# Remove any rows where there is no weight data
weigh_sheet <- weigh_sheet[!is.na(weigh_sheet$`Actual wt (mg)`), ]

# Read in the raw data from the VarioPYRO Cube
raw_data <- read_excel("Tillage 2.0 Post Soil CN.xlsx")

# Look in the Environment window (upper right).  You should see the two files that were read in.
# Check to make sure they have the same number of "observations" (rows)

#### EXTRACT THE DATA YOU NEED ####

# Combine the two datasets while dropping columns you don't need
raw1 <- cbind(weigh_sheet[ , c(1, 3, 4, 6)], 
              raw_data[ , c(1, 2, 4, 5, 7, 8, 10)])  

# Check whether the sample weights and descriptions line up
# If all the values in the selected columns match up, it will say "TRUE" in the window below.
# "FALSE" means that one or more values do not match up.  You should look and decide whether
# the mismatch is a real problem.  (Hint: click on the file name in the environment
# window to open it so you can scroll through it.)
identical(raw1$`Actual wt (mg)`, raw1$`Weight (mg)`)
identical(raw1$description, raw1$Name)  #MF 6/18/2020: Description column doesn't seem to have a corresponding column to match; always returns FALSE
identical(raw1$`Sample ID`, raw1$Name)


# Identify entries where weights don't match (added by MF 6/18/2020)
raw1[raw1$`Actual wt (mg)`!= raw1$`Weight (mg)`, ]
# Identify entries where sample names don't match (added by MF 6/18/2020)
raw1[raw1$`Sample ID`!= raw1$Name, ]


# Get the data you want for the "Processed data" sheet in your Excel workbook
#   and put the columns in the correct order
raw2 <- raw1[ , c(1:3, 5, 7:11)]

#### EXTRACT THE STANDARDS ####

acet <- raw2[grepl("ACETANILIDE", raw2$`Sample ID`, fixed=TRUE)==TRUE, ]
peach <- raw2[grepl("Peach leaves", raw2$`Sample ID`, fixed=TRUE)==TRUE, ]
rsmt <- raw2[grepl("Rosemount soil", raw2$`Sample ID`, fixed=TRUE)==TRUE, ]

# Put the three sets of standards together
stds <- rbind(acet, peach, rsmt)

#### EXTRACT THE REPLICATES ####

# Pull out the replicates
reps <- raw2[raw2$description=="unknown_replicate", ]
rep_ids <- unlist(strsplit(reps$`Sample ID`, "_"))
reps2 <- raw2[raw2$`Sample ID` %in% rep_ids, ]
reps3 <- rbind(reps, reps2)
#reps3$`Sample ID 2` <- sub("_.*", "", reps3$`Sample ID`)
#reps3 <- reps3[ , c(1:2, 10, 3:9)]
#reps3 <- reps3[order(reps3$`Sample ID`)]

#### GET R2, STD DEV, SLOPE, MEAN, EXPECTED VALUE, CALIBRATION VALUE #### (S. Pey, 3.24.2020) 

# Create a matrix to store the expected standard values, mean standard values and calculated
# calibration values for each of the four standards (Peach leaves, rosemount soil, acentanilide,
# and USGS 40)

qc <- matrix(NA, 18, 2)
row.names(qc) <- c("Acetanilide (r squared)", "Acetanilide (slope)", "Acetanilide (std dev)", 
                   "Acetanilide (mean)", "Aetenalide (expected)","Acetanilide (calibration)",
                   "Peach leaves (r squared)", "Peach leaves (slope)", "Peach leaves (std dev)", 
                   "Peach leaves (mean)", "Peach leaves (expected)","Peach leaves (calibration)",
                   "Rosemount Soil (r squared)", "Rosemount Soil (slope)", "Rosemount Soil (std dev)", 
                   "Rosemount Soil (mean)", "Rosemount Soil (expected)","Rosemount Soil (calibration)")
colnames(qc) <- c("N [%]","C [%]") 

## [ACETANILIDE] Calculate the r-squared, standard deviation, slope and mean standard values 
## for %N anc %C. Add the expected %N and %C values to the matrix.
## Calculate the calibration value and add it to the matrix.

qc[1, 1] <- summary(lm(`N [%]` ~ `Analysis order`, data = acet))$r.squared
qc[1, 2] <- summary(lm(`C [%]` ~ `Analysis order`, data = acet))$r.squared

qc[2, 1] <- summary(lm(`N [%]` ~ `Analysis order`, data = acet))$coefficients[2, 1]
qc[2, 2] <- summary(lm(`C [%]` ~ `Analysis order`, data = acet))$coefficients[2, 1]

qc[3, 1] <- sd(acet$`N [%]`)
qc[3, 2] <- sd(acet$`C [%]`)

qc[4, 1] <- mean(acet$`N [%]`)
qc[4, 2] <- mean(acet$`C [%]`)

qc[5,1] <- 10.36 # Expected %N, acetanilide
qc[5,2] <- 71.09 # Expected %C acetanilide

qc[6, 1] <- qc[4,1] / qc[5,1] # %N calibration, acetanilide
qc[6, 2] <- qc[4,2] / qc[5,2] # %C calibration, acetanilide

## [PEACH LEAVES] Calculate the r-squared, standard deviation, slope and mean standard values 
## for %N and %C. Add the expected %N and %C values to the matrix.
## Calculate the calibration value and add it to the matrix.
qc[7, 1] <- summary(lm(`N [%]` ~ `Analysis order`, data = peach))$r.squared
qc[7, 2] <- summary(lm(`C [%]` ~ `Analysis order`, data = peach))$r.squared

qc[8, 1] <- summary(lm(`N [%]` ~ `Analysis order`, data = peach))$coefficients[2, 1]
qc[8, 2] <- summary(lm(`C [%]` ~ `Analysis order`, data = peach))$coefficients[2, 1]

qc[9, 1] <- sd(peach$`N [%]`)
qc[9, 2] <- sd(peach$`C [%]`)

qc[10, 1] <- mean(peach$`N [%]`)
qc[10, 2] <- mean(peach$`C [%]`)

qc[11,1] <- 2.28   # Expected %N, peach leaves
qc[11,2] <- 50.40  # Expected %C, peach leaves

qc[12, 1] <- qc[10, 1] / qc[11,1]  # %N calibration, peach leaves
qc[12, 2] <- qc[10, 2] / qc[11,2]  # %C calibration, peach leaves

## [ROSEMOUNT SOIL] Calculate the r-squared, standard deviation, slope and mean standard values 
## for %N and %C. Add the expected %N and %C values to the matrix.
## Calculate the calibration value and add it to the matrix.

qc[13, 1] <- summary(lm(`N [%]` ~ `Analysis order`, data = rsmt))$r.squared
qc[13, 2] <- summary(lm(`C [%]` ~ `Analysis order`, data = rsmt))$r.squared

qc[14, 1] <- summary(lm(`N [%]` ~ `Analysis order`, data = rsmt))$coefficients[2, 1]
qc[14, 2] <- summary(lm(`C [%]` ~ `Analysis order`, data = rsmt))$coefficients[2, 1]

qc[15, 1] <- sd(rsmt$`N [%]`)
qc[15, 2] <- sd(rsmt$`C [%]`)

qc[16, 1] <- mean(rsmt$`N [%]`)
qc[16, 2] <- mean(rsmt$`C [%]`)

qc[17,1] <- 0.21   # Expected %N, Rosemount soil
qc[17,2] <- 2.22   # Expected %C, Rosemount soil

qc[18, 1] <- qc[16,1] / qc[17,1] # %N calibration, Rosemount soil
qc[18, 2] <- qc[16,2] / qc[17,2] # %C calibration, Rosemount soil


# Round all values to two decimal places
#qc <- round(qc, 2)
# Change matrix to data frame
qc_df <- data.frame(rownames(qc), qc, row.names = NULL)  # Add rownames form the matrix as first column (MF 6/19/2020)
# data.frame() changes colname format. Rename columns (MF 6/2020)
names(qc_df) <- c("QC Value", "N [%]", "C [%]")  

#### CORRECTING DATA (standard calibration) #### (S. Pey, 3/24/2020)

## Below, comment out (add a "#" symbol in front of the line) the standards that 
## you would NOT like to use for calibration (e.g.the standard you wish to use for 
## calibration should appear in black text). 

# %N calibration
#N.cal <- c(qc[6,1]) # [ACETANILIDE]
#N.cal <- c(qc[12,1]) # [PEACH LEAVES]
N.cal <- c(qc[18,1]) # [ROSEMOUNT SOIL]

# %C calibration
#C.cal <- c(qc[6,2]) # [ACETANILIDE]
#C.cal <- c(qc[12,2]) # [PEACH LEAVES]
C.cal <- c(qc[18,2]) # [ROSEMOUNT SOIL]

## Add the calibration values to the raw2 dataframe. Double check that the correct
## calibration values appear in the table. If the calibration value is incorrect,
## make sure that you have selected the correct standard to use for calibration in 
## the above section of code. 

raw2$N.cal <- N.cal
raw2$C.cal <- C.cal

## Apply calibration standards to the data and add the corrected data values into 
## the raw2 dataframe.
raw2$corrected.N <- (raw2$`N [%]` / raw2$N.cal)
raw2$corrected.C <- (raw2$`C [%]` / raw2$C.cal)

## Calculate and add C/N ratio values into the raw2 dataframe. ### (C. Loopstra, 6/26/2020)
corrected.CNratio <- (raw2$corrected.C/raw2$corrected.N)
raw2$corrected.CNratio <- corrected.CNratio

## Change raw2 column names to their final column names
colnames(raw2) = c("Analysis order", "Sample ID", 
                   "Description", "Weight (mg)",
                   "N Area", "C Area", "N [%]", "C [%]", "C/N ratio",
                   "N calibration value", "C calibration value", 
                   "Final Corrected %N", "Final Corrected %C","Final Corrected CN ratio")

#### SAVE THE DATA ####

sheet_list <- list("Raw data" = raw_data,
                   "Processed data" = raw2,
                   "Standards" = stds,
                   "QC" = qc_df,  # MF 6/18/2020: changed to the tibble version of qc data
                   "Replicates" = reps3)

write.xlsx(sheet_list, file = "Analyzed.Tillage2.0.PostSoilCN.rosemount.xlsx")  
