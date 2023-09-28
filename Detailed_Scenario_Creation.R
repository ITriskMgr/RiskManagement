#############################
# Run this fourth (4)
# This script generates detailed scenarios
# by enhancing the summary scenarios
#
# The input is determined by data_in_scenarios
# The output is the detailed scenarios 
# stored in the Excel file defined by data_out_detailed_scenarios
# a copy of the detailed scenarios are saved in ./Data/Scenarios
#############################

# Load the required libraries
library(readxl) # for reading excel files
library(writexl) # used to write the files
library(arsenal)
library(openai)
library(httr)
library(jsonlite)
library(stringr)
library(officer)
library(lubridate)
library(stringr)

# initialize the variables
# data_in_organization <- "./Data/Describe_Target.xlsx"
data_in_scenarios <- "./Data/Queries_Scenarios.xlsx"
data_out_detailed_scenarios <- "./Data/Detailed_Evaluation_Scenarios.xlsx"
data_out_scenarios = ""
my_API <- "xxxxxxxxxxxxxxx"
model = ""
messages=""
chat=""
query_message = ""
context_message = "Act as a senior cybersecurity consultant hired by a customer to help perform a scenario-based cybersecurity risk assessment."
role = "user"
content = ""
Query_cols = ""
Query_rows = ""
Query = ""
detailed_scenarios = ""

# Read the spreadsheet into R
data_scenarios <- read_excel(data_in_scenarios)

# remove problematic characters in NAME
# Replacing all occurrences of '/' with an empty string in the column 'NAME' of 'Scenarios.df'
data_scenarios$NAME <- gsub("/", "", data_scenarios$NAME)
# Replacing all occurrences of ',' with an empty string in the column 'NAME' of 'Scenarios.df'
data_scenarios$NAME <- gsub(",", "", data_scenarios$NAME)
# Replacing all occurrences of '-' with an empty string in the column 'NAME' of 'Scenarios.df'
data_scenarios$NAME <- gsub("-", "", data_scenarios$NAME)
# Replacing all occurrences of ' with an empty string in the column 'NAME' of 'Scenarios.df'
data_scenarios$NAME <- gsub("'", "", data_scenarios$NAME)
# Replacing all occurrences of : with an empty string in the column 'NAME' of 'Scenarios.df'
data_scenarios$NAME <- gsub(":", "", data_scenarios$NAME)
# Replacing all occurrences of ( with an empty string in the column 'NAME' of 'Scenarios.df'
data_scenarios$NAME <- gsub("\\(", "", data_scenarios$NAME)
# Replacing all occurrences of ) with an empty string in the column 'NAME' of 'Scenarios.df'
data_scenarios$NAME <- gsub("\\)", "", data_scenarios$NAME)
# Removing all occurrences of space from the 'NAME' column in the 'data_scenarios' data frame
data_scenarios$NAME <- gsub(" ", "_", data_scenarios$NAME)
# Limiting the length of each string in the column 'NAME' to 20 characters
data_scenarios$NAME <- substr(data_scenarios$NAME, 1, 20)
# Adding row number at the end of each string in the 'NAME' column, preceded by an underscore
data_scenarios$NAME <- paste(data_scenarios$NAME, 1:nrow(data_scenarios), sep="_")

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
# data_scenarios <- read_excel(data_in_scenarios)

#size the query
Query_cols <- ncol(data_scenarios)
Query_rows <- nrow(data_scenarios)
Query <- data_scenarios$Query
# Query_rows <- 5 # use for testing we used a smaller value to save processing 


# generate the detailed scenarios with ChatGPT
for (i in 1:Query_rows) {
  print(i)
  Random_time = floor(runif(1, min=20, max=70))
  
  
  # detailed_scenarios[i] <- hey_chatGPT(data_scenarios$Query[i])
  
  if(!is.null(hey_chatGPT(data_scenarios$Query[i]))) {
    detailed_scenarios[i] <- hey_chatGPT(data_scenarios$Query[i])
  
  
    # save the detailled scenario to a Word document to use later
    captured_text <- capture.output(cat(detailed_scenarios[i], sep = "\n"))
  
    # Initialize a Word document object
    doc <- read_docx()
  
    # Add captured text to the Word document
    for (line in captured_text) {
      doc <- body_add_par(doc, line, style = "Normal")
    }
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

