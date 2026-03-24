library(haven)
library(dplyr)
library(tidyr)
library(writexl)

# 1. Read data
input_file <- "901849.sav"
data <- read_sav(input_file, encoding = "cp1251")

# --- STEP 1: DEFINE SCENARIO PROBABILITIES ---
# Based on your table.
# Note: J2=5 is "Hard to say", J2=3 is "Rather not", J2=4 is "Def not"
prob_table <- tibble(
    J2 = c(1, 2, 5, 3, 4),
    Label = c("Def Yes", "Rather Yes", "Hard to Say", "Rather No", "Def No"),
    Prob_Optimistic = c(1.00, 0.75, 0.50, 0.25, 0.00),
    Prob_Middle = c(0.80, 0.60, 0.30, 0.10, 0.00),
    Prob_Pessimistic = c(0.60, 0.50, 0.20, 0.10, 0.00)
)

# --- STEP 2: CALCULATE ADULTS (RESPONDENTS) ---

adults_calc <- data %>%
    filter(!is.na(J2)) %>%
    left_join(prob_table, by = "J2") %>%
    summarise(
        Group = "Adults",
        Total_Weight = sum(Weight_2, na.rm = TRUE),

        # Calculate Weighted Counts for each scenario
        Val_Optimistic = sum(Weight_2 * Prob_Optimistic, na.rm = TRUE),
        Val_Middle = sum(Weight_2 * Prob_Middle, na.rm = TRUE),
        Val_Pessimistic = sum(Weight_2 * Prob_Pessimistic, na.rm = TRUE)
    ) %>%
    mutate(
        Share_Optimistic  = Val_Optimistic / Total_Weight,
        Share_Middle      = Val_Middle / Total_Weight,
        Share_Pessimistic = Val_Pessimistic / Total_Weight
    )

# --- STEP 3: CALCULATE CHILDREN ---

children_calc <- data %>%
    filter(!is.na(J2)) %>% # Parent must have an answer
    select(Weight_2, J2, A2.1_6, B1, starts_with("B2_")) %>%
    # Reshape to long format (one row per child)
    pivot_longer(
        cols = matches("^B2_\\d+$"),
        names_to = "child_index",
        values_to = "child_age"
    ) %>%
    # Extract child index number
    mutate(child_num = as.numeric(sub("B2_", "", child_index))) %>%
    # Filter: Keep only actual children (child_num <= Total Children B1)
    filter(child_num <= B1) %>%
    filter(!is.na(child_age)) %>%
    # Determine Weight Adjustment
    mutate(
        # If A2.1_6 == 1 (Has partner), divide weight by 2
        has_partner = if_else(!is.na(A2.1_6) & A2.1_6 == 1, 1, 0),
        child_weight = if_else(has_partner == 1, Weight_2 / 2, Weight_2)
    ) %>%
    # Attach Probabilities (Child inherits Parent's J2)
    left_join(prob_table, by = "J2") %>%
    summarise(
        Group = "Children",
        Total_Weight = sum(child_weight, na.rm = TRUE),

        # Calculate Weighted Counts
        Val_Optimistic = sum(child_weight * Prob_Optimistic, na.rm = TRUE),
        Val_Middle = sum(child_weight * Prob_Middle, na.rm = TRUE),
        Val_Pessimistic = sum(child_weight * Prob_Pessimistic, na.rm = TRUE)
    ) %>%
    mutate(
        Share_Optimistic  = Val_Optimistic / Total_Weight,
        Share_Middle      = Val_Middle / Total_Weight,
        Share_Pessimistic = Val_Pessimistic / Total_Weight
    )

# --- STEP 4: COMBINE RESULTS ---

final_result <- bind_rows(adults_calc, children_calc) %>%
    select(Group, Total_Weight, Share_Optimistic, Share_Middle, Share_Pessimistic)

# Print to console
print(final_result)

# Export to Excel
write_xlsx(final_result, "Return_Shares_Adults_Children.xlsx")
