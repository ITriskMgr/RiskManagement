#############################
# Run this sixth (6)
# This script is run the simulation 
#
# The input is the detailed metrics scenarios 
# stored in the Excel file defined by Detailed_Metrics_Scenarios.xlsx
# The results are stored in an excel file defined by Scenarios_Simulations.xlsx
#############################
# load the libraries
library(readxl)
library(tidyr)
library(dplyr)
library(stringr)
library(writexl)

risk_tolerance = 0.4
Facteur_Organisationel = 1000

data_metrics = "./Data/Detailed_Metrics_Scenarios.xlsx"
results_metrics = "./Data/Scenarios_Simulations.xlsx"

df <- read_excel(data_metrics)

# Convert a specific column to numeric in case it is not
#df$Exp_Damages <- as.numeric(df$Exp_Damages)
#df$Max_Damages <- as.numeric(df$Max_Damages)

# Initialize a new column in the dataframe to store the simulation results
#df$Simulation_Result <- NA


# Assuming df is your dataframe
# n_trials is the number of trials per row

simulate_monte_carlo <- function(exp_damage, max_damage, n_trials=10000) {
  results <- rnorm(n_trials, mean=exp_damage, sd=(max_damage - exp_damage)*3)
  mean_result <- mean(results)
  return(mean_result)
}

monte_carlo_simulation_2 <- function(value_Imp_Reduction, value_Pb_Reduction, n = 10000) {
  sim_Imp_Reduction <- rnorm(n, mean = value_Imp_Reduction, sd = 0.2)
  sim_Pb_Reduction <- rnorm(n, mean = value_Pb_Reduction, sd = 0.2)
  
  result <- mean(sim_Imp_Reduction * sim_Pb_Reduction)
  return(result)
}


# Vectorized approach for running the simulation for each row
n_trials <- 1000000

# Create a new column to store the Monte Carlo simulation results for the damages
df$MonteCarlo_Damages <- mapply(simulate_monte_carlo, df$Exp_Damages, df$Max_Damages, MoreArgs=list(n_trials=n_trials))
# Round off the values in the 'MonteCarlo_Result' column to two decimal places

df <- df %>%
  rowwise() %>%
  mutate(MonteCarlo_Mitigated = monte_carlo_simulation_2(Imp_Reduction, Pb_Reduction, n = n_trials)) %>%
  ungroup()

# Calculate the KRI
df$KRI_Estimate <- ((df$Pb_Threat * df$Pb_Exploitation * df$MonteCarlo_Damages * df$Exp_Utility * (df$CVSS_Score / 10)) / df$Resilience) * Facteur_Organisationel
df$KRI_Max <- ((df$Pb_Threat * df$Pb_Exploitation * df$Max_Damages * df$Exp_Utility * (df$CVSS_Score / 10)) / df$Resilience) * Facteur_Organisationel
df$KRI_Tolerate <- (((1-risk_tolerance) * df$Pb_Threat * df$Pb_Exploitation * df$Exp_Damages * df$Exp_Utility * (df$CVSS_Score / 10)) / df$Resilience) * Facteur_Organisationel
df$KRI_Mitigated <- df$MonteCarlo_Mitigated * df$KRI_Estimate
df$KRI_Residual <- df$KRI_Estimate - df$KRI_Mitigated
df$KRI_Ratio <- (df$KRI_Mitigated / df$Mitigation_Cost) * Facteur_Organisationel

# Round the KRI to 2 digits
df$MonteCarlo_Damages <- round(df$MonteCarlo_Damages, 3)
df$MonteCarlo_Mitigated <- round(df$MonteCarlo_Mitigated, 3)
df$KRI_Estimate <- round(df$KRI_Estimate, 2)
df$KRI_Max <- round(df$KRI_Tolerate, 2)
df$KRI_Tolerate <- round(df$KRI_Tolerate, 2)
df$KRI_Mitigated <- round(df$KRI_Mitigated, 2)
df$KRI_Residual <- round(df$KRI_Residual, 2)
df$KRI_Ratio <- round(df$KRI_Ratio, 4)

# Save the processed dataframe
write_xlsx(df, results_metrics)
