library(haven)
library(writexl)
library(dplyr)
library(tidyr)

input_file <- "901849.sav"

data <- read_sav(input_file, encoding = "cp1251")

# 1. Create a copy of the variable Z1
data$Z1_corrected <- data$Z1

# 2. Identify rows where Region is 'Kyiv Oblast' (10) AND City is 'Kyiv' (21)
# We use which() to safely ignore NA values
rows_to_fix <- which(data$Z1 == 10 & data$Z3.1 == 21)

# 3. Update those specific rows to 'Kyiv City' (11)
data$Z1_corrected[rows_to_fix] <- 11

# --- Optional: Verify the change ---

# Check counts before (in original)
print("Original Z1 count for Kyiv Oblast (10) and Kyiv City (11):")
table(data$Z1[data$Z1 %in% c(10, 11)])

# Check counts after (in corrected)
print("Corrected Z1 count for Kyiv Oblast (10) and Kyiv City (11):")
table(data$Z1_corrected[data$Z1_corrected %in% c(10, 11)])

# 3. Calculate the sum of weights and shares
region_summary <- data %>%
    group_by(Region = as_factor(Z1_corrected)) %>%
    summarise(
        Total_Weight = sum(Weight_2, na.rm = TRUE)
    ) %>%
    mutate(
        Share_Percent = (Total_Weight / sum(Total_Weight)) * 100
    ) %>%
    arrange(desc(Total_Weight))

# 4. View the result
print(region_summary)

write_xlsx(region_summary, "origin_region_cleanedKyiv.xlsx")
