#Libraries
library(tidyverse)
library(readr)
library(readxl)
library(purrr)
library(lubridate)

#this directory needs to be changed when you start a new quarter
base = "C:/Users/kkelderman/Documents/Scripts/Benzene/Refineries/"
new = "12_Downloads_Sep-18-2023/"
ftc = "FTC Files/2023 Q2/"
setwd(str_c(base, new))

#Column names for each tab in an excel spreadsheet
facility_cols <- c('Facility_Record', 'Facility', 'Address', 'Address_2', 'City',
                   'County', 'State', 'Zip', 'Agency_ID', 'Year', 'Quarter', 'Extra',
                   'Extra_2')
period_cols <- c('Facility_Record','Sampling_Period_ID', 'Sampling_Start_Date',
                'Sampling_End_Date', 'Sample_Concentration_Change',
                'Annual_Avg_Concentration_Change', 'Comments')
sampler_cols <- c('Facility_Record','Sampler Name', 'Latitude',
                 'Longitude', 'Passive Sampler Type', 'Comments')
results_cols <- c('Facility_Record','Sampling_Period_ID', 'Sampler Name', 'Passive Sampler Type',
                 'Sampling Benzene Concentration', 'Corrected Benzene Concentration',
                 'Below MDL?', 'Lab Reported Benzene Concentration', 'Outlier?', 'Skipped?',
                 'Other Data Flags', 'Explanation')

#### Reading in the data from all the spreadsheets into separate lists ####
refinery_files <- dir(pattern = "\\.xlsx") #skipping the first 12 rows of each sheet

facility_list <- list()
period_list <- list()
sampler_list <- list()
results_list <- list()


for (i in 1:length(refinery_files)){
  #reading in facility tab data
  facility_list[[i]] <- read_xlsx(refinery_files[i], sheet = 3,
                                  col_names = facility_cols, skip = 12)
  facility_list[[i]]$Count <- i
  
  #reading in period tab data
  period_list[[i]] <- read_xlsx(refinery_files[i], sheet = 4,
                                col_names = period_cols, skip = 12,
                                col_types = c("text", "text", "text", "text", "numeric", "numeric", "text"))
  period_list[[i]]$Count <- i
  
  #reading in sampler tab data
  sampler_list[[i]] <- read_xlsx(refinery_files[i], sheet = 5,
                                 col_names = sampler_cols, skip = 12)
  sampler_list[[i]]$Count <- i
  
  #reading in result tab data
  results_list[[i]] <- read_xlsx(refinery_files[i], sheet = 6,
                                 col_names = results_cols, skip = 12)
  results_list[[i]]$Count <- i
  
}

#### Cleaning the data (filtering out example rows and blanks) ####
rows_remove = c('XML Tag:', 'e.g.: 1', 'e.g.: ER01',
                'RecordId')

filter_data_tabs <- function(list, rows_remove) {
  filtered_results <- lapply(list, function(x) {
    filter(x,
           !is.na(Facility_Record),
           !Facility_Record %in% rows_remove)
  })
  return(filtered_results)
}

facilities <- filter_data_tabs(facility_list, rows_remove)
periods <- filter_data_tabs(period_list, rows_remove)
samples <- filter_data_tabs(sampler_list, rows_remove)
results <- filter_data_tabs(results_list, rows_remove)

#### Creating raw files for FTC ####
#facility join file in the format for FTC column order
facility_df <- facilities %>%
  reduce(rbind) %>%
  unique()

#checking for incorrect state abbreviations (Wa, New Jersey, Texas)
print(unique(facility_df$State))

facility_df$State[facility_df$State == "Texas"] <- "TX"
facility_df$State[facility_df$State == "New Jersey"] <- "NJ"
facility_df$State[facility_df$State == "Wa"] <- "WA"

fac_join <- facility_df %>%
  select(Facility_Record, Facility, City, State, Year, Quarter, Count)

#raw facility file (no name corrections)
ftc_fac <- facility_df %>%
  select(-Count) %>%
  unique()
write_excel_csv(ftc_fac, str_c(base, ftc, "00_raw_facilities.csv"),na="")

#period file
ftc_per <- periods %>%
  reduce(rbind) %>%
  left_join(fac_join, by = c("Facility_Record", "Count"))

#this is to deal with the tricky date issue I was having in 2023 Quarter 1/2 download
period_date_1 <- ftc_per %>%
  filter(str_detect(Sampling_End_Date, "\\/"))
period_date_1$Sampling_End_Date <- str_replace_all(period_date_1$Sampling_End_Date, "\\/23", "\\/2023")
period_date_1$Sampling_Start_Date <- str_replace_all(period_date_1$Sampling_Start_Date, "\\/23", "\\/2023")
period_date_1$Sampling_End_Date <- as.Date(period_date_1$Sampling_End_Date, "%m/%d/%Y")
period_date_1$Sampling_Start_Date <- as.Date(period_date_1$Sampling_Start_Date, "%m/%d/%Y")

period_date_2 <- ftc_per %>%
  filter(!str_detect(Sampling_End_Date, "\\/"))
period_date_2$Sampling_End_Date <- as.Date(as.numeric(period_date_2$Sampling_End_Date),
                                           origin = "1899-12-30")
period_date_2$Sampling_Start_Date <- as.Date(as.numeric(period_date_2$Sampling_Start_Date),
                                             origin = "1899-12-30")

ftc_per <- rbind(period_date_1, period_date_2) %>%
  select(9:13,1:7)

write_excel_csv(ftc_per, str_c(base, ftc, "00_raw_period.csv"),na="")

#sampler file
sampler_raw <- samples %>% 
  reduce(rbind) %>%
  left_join(fac_join, by = c("Facility_Record", "Count")) %>%
  select(8:10,2:4,6) %>%
  filter(Latitude != '--') %>% #removing '--' from coordinate column (for a trip blank)
  unique()

#getting rid of degree symbol and extra characters in coordinate columns
sampler_raw$Longitude <- gsub("\u00b0 W", "", paste(sampler_raw$Longitude))
sampler_raw$Latitude <- gsub("\u00b0 N", "", paste(sampler_raw$Latitude))

#making all values in longitude column negative (which they should all be)
#before doing this, make sure all values in the Latitude column are positive
#if there are negative values, it's possible the lat and long were switched
#in which case you need to fix that manually
sampler_raw$Longitude <- as.numeric(as.character(sampler_raw$Longitude))
sampler_raw$Longitude <- with(sampler_raw, ifelse(Longitude > 0, Longitude*-1, Longitude))

write_csv(sampler_raw, str_c(base, ftc,"00_sampler_info.csv"),na="")

#results file
results_raw <- results %>% 
  reduce(rbind) %>%
  left_join(fac_join, by = c("Facility_Record", "Count")) %>%
  select(-Count)

#checking for characters in benzene concentration data
char_check <- results_raw %>%
  mutate(check_conc = str_detect(`Sampling Benzene Concentration`, "[^0-9.]"),
         check_corr = str_detect(`Corrected Benzene Concentration`, "[^0-9.]")) %>%
  filter(check_conc == TRUE | check_corr == TRUE)

#writing raw data
write_excel_csv(results_raw, str_c(base, ftc,"00_raw_results.csv"), na='')
