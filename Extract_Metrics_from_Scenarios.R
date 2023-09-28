#############################
# Run this fifth (5)
# This script is used to extract 
# the metrics from the detailed scenarios
#
# The input is the detailed scenarios 
# stored in the Excel file defined by data_in_detailed_scenarios
# The metrics are stored in an excel file defined by metrics_in_detailed_scenarios
#############################
# load the libraries
library(readxl)
library(tidyr)
library(dplyr)
library(stringr)
library(writexl)

# initialize the variables
data_list_scenarios <- "./Data/Cleaned_List_Scenarios.xlsx"
data_queries_scenarios <- "./Data/Queries_Scenarios.xlsx"
data_in_detailed_scenarios <- "./Data/Detailed_Evaluation_Scenarios.xlsx"
# data_in_scenarios = ""
metrics_out  <- "./Data/Detailed_Metrics_Scenarios.xlsx"

# Query_cols = ""
# Query_rows = ""
# Query = ""
# detailed_scenarios = ""
# regex_patterns <- paste0(prompts, " ([0-9.]+)")

# Read the Excel file
df <- read_excel(data_in_detailed_scenarios)
data_scenarios <- read_excel(data_in_detailed_scenarios)
df_queries <- read_excel(data_queries_scenarios)
df_list <- read_excel(data_list_scenarios)

# put all the data in lower case
df$detailed_scenarios <- str_to_lower(df$detailed_scenarios)

# Replace all dots in decimal numbers with commas
# df$detailed_scenarios <- str_replace_all(df$detailed_scenarios, "(\\d)\\.(\\d)", "\\1,\\2")

# Remove all commas used as thousand separators for money
# df$detailed_scenarios <- str_replace_all(df$detailed_scenarios, "(?<=\\d),(?=\\d{3}(?![\\d,]))", "")
# Remove all commas used as thousand separators, but preserve decimal commas
# df$detailed_scenarios <- str_replace_all(df$detailed_scenarios, "(?<=\\d),(?=\\d{1,2}(\\D|$))", "")
# Remove all commas from the detailed_scenarios column
df$detailed_scenarios <- str_replace_all(df$detailed_scenarios, ",", "")
# Remove all $ from the detailed_scenarios column
df$detailed_scenarios <- str_replace_all(df$detailed_scenarios, "\\$", "")
# To remove all occurrences of the pattern "(0-1)"
df$detailed_scenarios <- str_replace_all(df$detailed_scenarios, "\\(0-1\\)", "")
# To remove all occurrences of the pattern "usd"
df$detailed_scenarios <- str_replace_all(df$detailed_scenarios, "usd", "")
# remove any triple spaces
df$detailed_scenarios <- str_replace_all(df$detailed_scenarios, "\\s{3,}", " ")
# remove any double spaces
df$detailed_scenarios <- str_replace_all(df$detailed_scenarios, "\\s{2,}", " ")

# Define the patterns to be replaced and their replacement
replacement_patterns <- c(

  "threat probability" = "threat estimate",
  "probability that the threat will be present" = "threat estimate",
  "probability of threat presence" = "threat estimate",
  "probability of presence" = "threat estimate",
  "probability of threat" = "threat estimate",
  "threat estimate occurrence" = "threat estimate",
  "threat estimate exploitation" = "probability of exploitation",
  
  "probability of exploitation" = "exploitation estimate",
  "probability of vulnerability exploitation" = "exploitation estimate",
  
  "estimated expected damages" = "expected damages",
  "estimated impact" = "expected damages",
  "maximal impact" = "maximal damages",
  "maximal damages: 1" = "maximal damages: 0.95",
  "maximal damages:1" = "maximal damages: 0.95",
  
  "level of organizational resilience" = "organizational resilience",

  "impact reduction for proposed mitigation measures" = "impact reduction",
  "impact reduction of proposed mitigation measures" = "impact reduction",
  "impact reduction for proposed mitigation cost" = "impact reduction",
  "impact reduction of proposed mitigation cost" = "impact reduction",
  "probability reduction for proposed mitigation measures" = "probability reduction",
  "probability reduction of proposed mitigation measures" = "probability reduction",
  "probability reduction of proposed mitigation cost" = "probability reduction",
  "probability reduction for proposed mitigation cost" = "probability reduction",
  
  "estimated cvss score" = "cvss_score",
  "cvss score" = "cvss_score",
  "base score" = "cvss_score",
  "overall score" = "cvss_score",
  
  "mitigations" = "mitigation",
  "implementing mitigation cost is " = "",
  "total cost" = "mitigations_costs",
  "budgetary estimate for mitigation measures" = "mitigations_costs",
  "budgetary estimate" = "mitigations_costs",
  "mitigation measures" = "mitigations_costs",
  "mitigation estimate" = "mitigations_costs",
  "mitigation costs estimate" = "mitigations_costs",
  "mitigation cost for implementation" = "mitigations_costs",
  "mitigation costs is" = "mitigations_costs",
  "mitigation costs" = "mitigations_costs",
  "mitigation cost" = "mitigations_costs",
  "proposed mitigations_costs: 1. " = "",
  "proposed mitigations_costs " = "",
  "mitigations_costs: 1. " = ""
  )

# Apply the replacements
df$detailed_scenarios <- str_replace_all(df$detailed_scenarios, replacement_patterns)

# Extract metrics using regex and store in separate columns
df <- df %>%
  mutate(
    Pb_Threat = str_extract(detailed_scenarios, "(?<=threat estimate: )\\d+(\\.\\d{1,2})?"),
    Pb_Exploitation = str_extract(detailed_scenarios, "(?<=exploitation estimate: )\\d+(\\.\\d{1,2})?"),
    Exp_Damages = str_extract(detailed_scenarios, "(?<=expected damages: )\\d+(\\.\\d{1,2})?"),
    Max_Damages = str_extract(detailed_scenarios, "(?<=maximal damages: )\\d+(\\.\\d{1,2})?"),
    Resilience = str_extract(detailed_scenarios, "(?<=organizational resilience: )\\d+(\\.\\d{1,2})?"),
    Exp_Utility = str_extract(detailed_scenarios, "(?<=expected utility: )\\d+(\\.\\d{1,2})?"),
    CVSS_Score = str_extract(detailed_scenarios, "(?<=cvss_score: )\\d+(\\.\\d{1,2})?"),
    Mitigation_Cost = str_extract(detailed_scenarios, "(?<=mitigations_costs: )\\d+(\\.\\d{1,2})?"),
    Imp_Reduction = str_extract(detailed_scenarios, "(?<=impact reduction: )\\d+(\\.\\d{1,2})?"),
    Pb_Reduction = str_extract(detailed_scenarios, "(?<=probability reduction: )\\d+(\\.\\d{1,2})?")
  ) 

# Convert the extracted columns to numeric data type for calculations
cols_to_convert <- c("Pb_Threat",
                     "Pb_Exploitation", 
                     "Exp_Damages", 
                     "Max_Damages", 
                     "Resilience", 
                     "Exp_Utility", 
                     "CVSS_Score", 
                     "Mitigation_Cost", 
                     "Imp_Reduction", 
                     "Pb_Reduction"
                     )

df[cols_to_convert] <- lapply(df[cols_to_convert], as.numeric)


# Perform a full join
# Create an index column
df_list$id <- seq_len(nrow(df_list))
df_queries$id <- seq_len(nrow(df_queries))
data_scenarios$id <- seq_len(nrow(data_scenarios))
df$id <- seq_len(nrow(df))

# To remove a column named 'unwanted_column':
#dataframe_name <- select(dataframe_name, -unwanted_column)


#combined_dataframe_3 <- full_join(df_queries, df_list, by = "id" )
#combined_dataframe_2 <- full_join(df_list, data_scenarios, by = "id" )
combined_dataframe <- full_join(df_list, df, by = "id" )

# do a final cleaning of the data in 'combined_dataframe'
combined_dataframe <- rename(combined_dataframe, "No" = "id")
combined_dataframe <- combined_dataframe %>%
  filter(
    !is.na(Pb_Threat),
    !is.na(Pb_Exploitation), 
    !is.na(Exp_Damages), 
    !is.na(Max_Damages), 
    !is.na(Resilience), 
    !is.na(Exp_Utility), 
    !is.na(CVSS_Score), 
    !is.na(Mitigation_Cost), 
    !is.na(Imp_Reduction), 
    !is.na(Pb_Reduction))

# Save the processed dataframe
write_xlsx(combined_dataframe, metrics_out)

