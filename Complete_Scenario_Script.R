#############################
# Run this first (1)
# This script generates summary scenarios
# saved as a text file in the data_list_scenarios (variable)
# query_rows is used to dtermine the number of times to run the query
# query_message is used to determine the nimber and details of the scenarios
#############################

# load the libraries
library(openai)
library(readxl) 
library(writexl)
library(arsenal)
library(httr)
library(jsonlite)
library(stringr)
library(readxl)
library(tidyr)
library(dplyr)
library(tidyverse)
library(officer)
library(lubridate)

# initialize the variables
my_API <- "xxxxxxxxxxxxx"

data_list_scenarios <- "./Data/List_Scenario.txt"
data_in_scenarios <- "./Data/List_Scenario.txt"
data_out_scenarios <- "./Data/Cleaned_List_Scenarios.xlsx"
data_in_detailed_scenarios <- "./Data/Detailed_Evaluation_Scenarios.xlsx"
data_metrics_out  <- "./Data/Detailed_Metrics_Scenarios.xlsx"

data_scenarios = ""
model = ""
messages=""
chat=""
query = ""
query_role = ""
query_message = ""
query_content = ""
query_cols = ""
query_rows = ""
detailed_scenarios = ""
summary_scenarios = ""
scenarios_list = ""
scenarios_list2 = ""
data_scenarios = ""
scenario.df = data.frame()
summary_scenarios.df = data.frame()

context_role = "user"
context_message = "You are a cybersecurity consultant, working in Canada, hired by a customer to perform a scenario-based cynersecurity risk assessment."
query_rows <- 25 # for testing
query_message <- "create 20 cybersecurity risk scenarios for the Montreal Regional Hospital, a large 500 bed hospital located in Montreal. It has a IT infrastructure that includes a financial management and accounting system, a CRM, a website and multiple healthcare information systems. In the answer provide a scenario name in the format NAME: and a short scenario description in the format DESCRIPTION:"

# Asking Questions to ChatGPT, Saving and Cleaning Answer
hey_chatGPT <- function(answer_my_question) {
  chat_GPT_answer <- POST(
    url = "https://api.openai.com/v1/chat/completions",
    add_headers(Authorization = paste("Bearer", my_API)),
    content_type_json(),
    encode = "json",
    body = list(
      model = "gpt-3.5-turbo",
      messages = list(
        list(
          role = context_role,
          content = context_message
        ),
        list(
          role = "user",
          content = answer_my_question
        )
      )
    )
  )
  str_trim(content(chat_GPT_answer)$choices[[1]]$message$content)
}

# generate the detailed scenarios with ChatGPT
for (i in 1:query_rows) {
  print(i)
  random_time = floor(runif(1, min=3, max=9))
  summary_scenarios[i] <- hey_chatGPT(query_message)
  Sys.sleep(random_time)
  write(summary_scenarios[i], file = data_list_scenarios, append = TRUE, sep = "\n\n")
}

#############################
# Part (2)
# This script cleans the summary scenarios
# to prepare them for the next step
# by removing the empty lines
# and using NAME and DESCRIPTION
# to prepare the list for the next step 
#############################

# Read the file content
file_content <- readLines(data_in_scenarios, warn = FALSE)

# Filter out blank lines
temp_file_content <- file_content[!grepl("^\\s*$", file_content)]

# Filter lines that start with "NAME:" or "DESCRIPTION:"
filtered_content <- temp_file_content[grepl("^(NAME:|DESCRIPTION:)", temp_file_content)]
# cat(filtered_content)

# Modify lines that start with "1. NAME:" or similar patterns
modified_content <- sub("^\\d+\\.\\s*(NAME:|DESCRIPTION:)", "\\1", file_content)

# Extract content using regular expressions
name_pattern <- "(?<=NAME: )[^\\n]*"
description_pattern <- "(?<=DESCRIPTION: )[^\\n]*"

names <- str_extract(modified_content, name_pattern) %>% na.omit()
descriptions <- str_extract(modified_content, description_pattern) %>% na.omit()

# Ensure both vectors are of the same length
if (length(names) != length(descriptions)) {
  stop("Mismatch in the number of names and descriptions.")
}

# Convert to dataframe
summary_scenarios.df <- data.frame(
  NAME = names,
  DESCRIPTION = descriptions,
  stringsAsFactors = FALSE
)

# sauvegarde du fichier
write_xlsx(summary_scenarios.df,data_out_scenarios)

#############################
# Part (3)
# This script is used to prepare the queries
# from the list of scenarios them for the next step
#############################

# Read the Excel File
Scenarios.df <- read_excel(data_in_scenarios) 

# The Excel file has a column named 'NAME' and 'DESCRIPTION'
Scenario_Name <- Scenarios.df$NAME
Scenario_Description <- Scenarios.df$DESCRIPTION

# Assemble the query
prefix <- "Expand the cybersecurity risk scenario "
midfix_1  <- " for the organization "
organization_name <- organization_define
midfix_2 <- ". For the scenario described as "
# suffix <- " Provide a detailed cybersecurity risk scenario. Develop a List of the stakeholders involved. Present a bullet list of Event sequence, a description of the Consequences, some Historical data and proposed Mitigation measures. On a scale of 0 to 1, indicate the Probability that the threat will be present, the Probability of exploitation, the Estimated expected damages, the Maximal damages, the Level of organizational resilience, and the Expected utility. All of these need to be on a scale of 0 to 1. Calculate the CVSS version 3.1 score of the vulnerabilities and provide the details of the calculation. Include a budgetary estimate for the costs of implementing the proposed Mitigation measures. Indicate the impact reduction on a scale of 0 to 1 for the proposed Mitigation measures.Indicate the probability reduction on a scale of 0 to 1 for the proposed Mitigation measures."
suffix <- " Using this information, create a detailed cybersecurity risk scenario. Identify the stakeholders involved. Provide an event sequence leading to this scenario. Provide a description of the consequences. Provide historical data for a similar scenario. On a scale of 0 to 1, provide the probability that the threat will be present, presented as threat probability. On a scale of 0 to 1, provide the probability of exploitation of the vulnerability by the threat, presented as probability of exploitation. On a scale of 0 to 1, provide the estimated expected damages, presented as expected damages. On a scale of 0 to 1, provide the maximal damages, presented as maximal damages. On a scale of 0 to 1, provide the level of organizational resilience, presented as organizational resilience. On a scale of 0 to 1, provide the expected utility of the assets at risk in this scenario, presented as expected utility. Calculate the CVSS version 3.1 score for the vulnerabilities, presented as Estimated CVSS Score. Provide proposed mitigation measures for this scenario. Calculate the costs of implementing the proposed mitigation measures for this scenario, calculate the cost of the mitigation measures, presented the costs as mitigation costs. On a scale of 0 to 1, provide the impact reduction effect for the proposed mitigation measures, presented as impact reduction. On a scale of 0 to 1, provide the probability reduction effect for the proposed mitigation measures, presented as probability reduction."

# Assemble the query
Scenarios.df$query <- paste0(prefix, Scenario_Name, midfix_1, organization_name, midfix_2, Scenario_Description, suffix)

# Save the results of this
write_xlsx(Scenarios.df, data_out_queries)

#############################
# Part (4)
# This script generates detailed scenarios
# by enhancing the summary scenarios
#############################

# Read the spreadsheet into R
data_scenarios <- read_excel(data_in_scenarios)

# Asking Questions to ChatGPT, Saving and Cleaning Answer
hey_chatGPT <- function(answer_my_question) {
  chat_GPT_answer <- POST(
    url = "https://api.openai.com/v1/chat/completions",
    add_headers(Authorization = paste("Bearer", my_API)),
    content_type_json(),
    encode = "json",
    body = list(
      model = "gpt-3.5-turbo-0301",
      messages = list(
        list(
          role = "user",
          content = answer_my_question
        )
      )
    )
  )
  str_trim(content(chat_GPT_answer)$choices[[1]]$message$content)
}

# Read the spreadsheet into R
data_scenarios <- read_excel(data_in_scenarios)

#size the query
query_cols <- ncol(data_scenarios)
query_rows <- nrow(data_scenarios)
query <- data_scenarios$query
# query_rows <- 5 # use for testing we used a smaller value to save processing 

# generate the detailed scenarios with ChatGPT
for (i in 1:query_rows) {
  print(i)
  Random_time = floor(runif(1, min=3, max=10))
  detailed_scenarios[i] <- hey_chatGPT(data_scenarios$query[i])
  
  # save the detailled scenario to a Word document to use later
  captured_text <- capture.output(cat(detailed_scenarios[i], sep = "\n"))
  
  # Initialize a Word document object
  doc <- read_docx()
  
  # Add captured text to the Word document
  for (line in captured_text) {
    doc <- body_add_par(doc, line, style = "Normal")
  }
  
  # Create file name with variable name + timestamp
  file_name_variable <- data_scenarios$NAME[i]
  file_name_variable <- gsub("/", "-", file_name_variable)
  timestamp <- format(now(), "%Y%m%d_%H%M%S")
  final_file_name <- paste0("./Data/Scenarios/", file_name_variable, "_", timestamp, ".docx")
  
  # Save the Word document
  print(doc, target = final_file_name)
  
  # wait for a few random seconds so openai won't think you are abusing
  Sys.sleep(Random_time)
}

# sauvegarde du fichier
data_out_scenarios <- as.data.frame(detailed_scenarios)
write_xlsx(data_out_scenarios,data_out_detailed_scenarios)

#############################
# Part (5)
# This script is used to extract 
# the metrics from the detailed scenarios
#############################

# Read the Excel file
df <- read_excel(data_in_detailed_scenarios)

# put all the data in lower case
df$detailed_scenarios <- str_to_lower(df$detailed_scenarios)

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

# Save the processed dataframe
write_xlsx(df, data_metrics_out)
