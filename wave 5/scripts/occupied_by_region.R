library(haven)
library(writexl)
library(dplyr)
library(tidyr)

# Read the SAV file
data <- read_sav("901849.sav", encoding = "cp1251")

# View the first few rows
head(data)

# Create the summary table with total weights and shares
occupied_summary <- data %>%
  # Filter out missing values for key variables
  filter(!is.na(Z1), !is.na(Q6), !is.na(Weight)) %>%
  # Group by the oblast name (using the labels)
  group_by(Oblast = as_factor(Z1)) %>%
  # Calculate required metrics
  summarise(
    # 1. Total Weight of people from the given oblast
    Total_Weight_Oblast = sum(Weight, na.rm = TRUE),
    
    # 2. Total weight of people from the oblast whose hometown is occupied
    Weight_Occupied = sum(Weight[Q6 == 1], na.rm = TRUE),
    
    # 3. Share of respondents whose hometown is occupied
    Share_Occupied_Pct = (Weight_Occupied / Total_Weight_Oblast) * 100,
    
    # Optional: raw count of respondents for context
    Sample_Size = n()
  ) %>%
  # Sort by the share in descending order
  arrange(desc(Share_Occupied_Pct))

# View the result
print(occupied_summary)
