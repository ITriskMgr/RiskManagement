#############################
# Run this first (1)
# This script generates summary scenarios
# saved as a text file in the list_scenarios (variable)
# Query_rows is used to dtermine the number of times to run the query
# Query_message is used to determine the nimber and details of the scenarios
# The input is determined by the variables: 
#      Context_role
#      Context_message 
#      Query_rows
#      Query_message
# The output is a txt file defined by the variable list_scenarios
#############################

# load the libraries
library(openai)
library(readxl) # for reading excel files
library(writexl) # used to write the files
library(arsenal)
library(httr)
library(jsonlite)
library(stringr)

# initialize the variables
# data_in_organization <- "./Data/Describe_Target.xlsx"
data_out_scenarios <- "./Data/List_Scenarios.xlsx"
list_scenarios <- "./Data/List_Scenario.txt"
data_scenarios = ""
model = ""
messages=""
chat=""
Query = ""
Query_role = ""
Query_message = ""
Query_content = ""
Query_cols = ""
Query_rows = ""
Detailed_scenarios = ""
Summary_scenarios = ""
Scenario.df = data.frame()
scenarios_list = ""
scenarios_list2 = ""

##############################
# OpenAI API 
my_API <- "xxxxxxxxxxxx"
# define the context
Context_role = "user"
Context_message = "You are a cybersecurity consultant, working in Canada, hired by a customer to perform a scenario-based cybersecurity risk assessment."
# define the query
Query_rows <- 50 # for testing use a lower number
Query_message <- "create 20 cybersecurity risk scenarios for the Montreal Regional Hospital, a large 500 bed hospital located in Montreal. It has a IT infrastructure that includes a financial management and accounting system, a CRM, a website and multiple healthcare information systems. In the answer provide a scenario name in the format NAME: and a short scenario description in the format DESCRIPTION:"
##############################

# Asking Questions to ChatGPT, Saving and Cleaning Answer
hey_chatGPT <- function(answer_my_question) {
  chat_GPT_answer <- POST(
    url = "https://api.openai.com/v1/chat/completions",
    add_headers(Authorization = paste("Bearer", my_API)),
    content_type_json(),
    encode = "json",
    body = list(
      model = "gpt-3.5-turbo",
      #model = "gpt-4",
      messages = list(
        list(
          role = Context_role,
          content = Context_message
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
for (i in 1:Query_rows) {
  # print("Generating lists of scenario: ", i)
  print(i)
  Random_time = floor(runif(1, min=1, max=5))
  # print(Random_time)
  Summary_scenarios[i] <- hey_chatGPT(Query_message)
  Sys.sleep(Random_time)
  write(Summary_scenarios[i], file = list_scenarios, append = TRUE, sep = "\n\n")
}
