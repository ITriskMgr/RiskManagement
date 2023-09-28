#############################
# Run this third (3)
# This script is used to prepare the queries
# from the list of scenarios them for the next step
# 
# The input comed from data_in_scenarios
# The outpout is an Excel file
# with the queries ready to be sent to ChatGPT
# as determined by data_out_queries 
#############################

# Load the required libraries
library(readxl) # for reading excel files
library(writexl) # used to write the files
library(arsenal)
library(openai)
library(httr)
library(jsonlite)
library(stringr)
library(tidyverse)
library(dplyr)

# initialize the variables
data_in_scenarios <- "./Data/Cleaned_List_Scenarios.xlsx"
data_out_queries <- "./Data/Queries_Scenarios.xlsx"
organization_define <- "Montreal Regional Hospital"

# Read the Excel File
Scenarios.df <- read_excel(data_in_scenarios) 

# Replacing all occurrences of '/' with an empty string in the column 'NAME' of 'Scenarios.df'
Scenarios.df$NAME <- gsub("/", "", Scenarios.df$NAME)

# The Excel file has a column named 'NAME' and 'DESCRIPTION'
Scenario_Name <- Scenarios.df$NAME
Scenario_Description <- Scenarios.df$DESCRIPTION

# Assemble the Query
prefix <- "Expand the cybersecurity risk scenario "
midfix_1  <- " for the organization "
organization_name <- organization_define
midfix_2 <- ". For the scenario described as "
# suffix <- " Provide a detailed cybersecurity risk scenario. Develop a List of the stakeholders involved. Present a bullet list of Event sequence, a description of the Consequences, some Historical data and proposed Mitigation measures. On a scale of 0 to 1, indicate the Probability that the threat will be present, the Probability of exploitation, the Estimated expected damages, the Maximal damages, the Level of organizational resilience, and the Expected utility. All of these need to be on a scale of 0 to 1. Calculate the CVSS version 3.1 score of the vulnerabilities and provide the details of the calculation. Include a budgetary estimate for the costs of implementing the proposed Mitigation measures. Indicate the impact reduction on a scale of 0 to 1 for the proposed Mitigation measures.Indicate the probability reduction on a scale of 0 to 1 for the proposed Mitigation measures."
suffix <- " Using this information, create a detailed cybersecurity risk scenario. Identify the stakeholders involved. Provide an event sequence leading to this scenario. Provide a description of the consequences. Provide historical data for a similar scenario. On a scale of 0 to 1, provide the probability that the threat will be present, presented as threat probability. On a scale of 0 to 1, provide the probability of exploitation of the vulnerability by the threat, presented as probability of exploitation. On a scale of 0 to 1, provide the estimated expected damages, presented as expected damages. On a scale of 0 to 1, provide the maximal damages, presented as maximal damages. On a scale of 0 to 1, provide the level of organizational resilience, presented as organizational resilience. On a scale of 0 to 1, provide the expected utility of the assets at risk in this scenario, presented as expected utility. Calculate the CVSS version 3.1 score for the vulnerabilities, presented as Estimated CVSS Score. Provide proposed mitigation measures for this scenario. Calculate the costs of implementing the proposed mitigation measures for this scenario, calculate the cost of the mitigation measures, presented the costs as mitigation costs. On a scale of 0 to 1, provide the impact reduction effect for the proposed mitigation measures, presented as impact reduction. On a scale of 0 to 1, provide the probability reduction effect for the proposed mitigation measures, presented as probability reduction."

# Assemble the query
Scenarios.df$Query <- paste0(prefix, Scenario_Name, midfix_1, organization_name, midfix_2, Scenario_Description, suffix)

# Save the results of this
write_xlsx(Scenarios.df, data_out_queries)

