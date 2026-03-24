# IMPORTANT! DOUBLE CHECK THE CODING OF THE GENDERS! 
# Previously, in wave 4, the labels were reversed so this script assumes that and reverses them. 
# Check how the genders of adults and children are coded. 1 = Male, 2 = Female, or vice versa.

install.packages(c("haven", "writexl", "dplyr", "tidyr", "readr"))

library(haven)
library(writexl)
library(dplyr)
library(tidyr)
library(readr)

input_file <- "901849.sav"
output_file <- "gender_age_pyramid.xlsx"

data <- read_sav(input_file, encoding = "cp1251")

# "З чоловіком\\дружиною чи партнером" used to establish if there are two parents per one child.
attr(data$A2.1_6, "label") # read the label of the answer category. Question A2.1, category 6.

# Each child inherits the Weight of the respondent (parent)
# but it should be divided by two, if there are two parents per one child.

# Select relevant columns
data_selected <- data %>%
    select(
        Weight_2, # Statistical Weight
        S1, # Respondent gender
        S2, # Respondent age
        A2.1_6, # Partnership status
        B1, # Number of children
        starts_with("B2_"), # Children's ages
        starts_with("B3_") # Children's genders
    )

# Rename columns for better understanding
data_selected <- data_selected %>%
    rename(
        respondent_gender = S1,
        respondent_age = S2,
        two_parents = A2.1_6,
        num_children = B1
    )

# Reshape children's ages and genders to long format
children_long <- data_selected %>%
    pivot_longer(
        cols = matches("^B[23]_\\d+$"), # Matches B2_1 to B2_5 (children's ages) and B3_1 to B3_5 (children's genders)
        names_to = c(".value", "child_num"),
        names_pattern = "(B[23])_(\\d+)" # Capture 'B2' or 'B3' and the child number
    ) %>%
    mutate(
        child_num = as.numeric(child_num), # Convert child_num to numeric
        child_gender = case_when(
            # IMPORTANT! DOUBLE CHECK THE LABELS!
            B3 == 1 ~ 2, # Recode: 1 (Female) -> 2 (Female) IMPORTANT! REVERSE ORDER OF GENDERS!
            B3 == 2 ~ 1, # Recode: 2 (Male) -> 1 (Male) IMPORTANT! REVERSE ORDER OF GENDERS!
            TRUE ~ NA_real_
        )
    ) %>%
    filter(child_num <= num_children) %>% # Only include existing children
    filter(!is.na(B2) & !is.na(B3)) %>% # Remove rows with missing age or gender
    rename(
        child_age = B2
    ) %>%
    select(Weight_2, child_age, child_gender, two_parents) %>%
    mutate(
        age = child_age,
        gender = child_gender, # Use recoded gender
        child = 1 # Indicator for children
    ) %>%
    select(Weight_2, age, gender, child, two_parents)

# Prepare respondent data
respondents <- data_selected %>%
    select(Weight_2, respondent_age, respondent_gender, two_parents) %>%
    mutate(
        age = respondent_age,
        gender = respondent_gender, # Gender: 1 = Male, 2 = Female
        child = 0 # Indicator for respondents
    ) %>%
    select(Weight_2, age, gender, child, two_parents)

# Convert 'gender' columns to numeric and remove labels to avoid conflicts
respondents <- respondents %>%
    mutate(gender = as.numeric(gender))

children_long <- children_long %>%
    mutate(gender = as.numeric(gender))

# Combine respondents and children into final dataset
final_data <- bind_rows(
    respondents,
    children_long
)

# Convert two_parents to binary (1 if has partner, 0 otherwise)
final_data <- final_data %>%
    mutate(
        two_parents = ifelse(two_parents == 1, 1, 0)
    )

# Convert gender codes to labels
final_data <- final_data %>%
    mutate(
        gender = case_when(
            gender == 1 ~ "Male",
            gender == 2 ~ "Female",
            TRUE ~ NA_character_
        )
    )

# Calculate adjusted Weight: for children with two parents, the Weight is divided by two.
# For building the population pyramid, USE ADJUSTED Weight! 
# Otherwise, the children share will be overrepresented.
final_data <- final_data %>%
  mutate(
    adjusted_Weight = if_else(child == 1 & two_parents == 1,
                              Weight_2 / 2,
                              Weight_2)
  ) %>%
  select(age, gender, child, two_parents, adjusted_Weight) 

# --- 1. CALCULATE CURRENT PYRAMID (2025) ---

# Grouping into brackets as per image provided
current_pyramid_raw <- final_data %>%
  mutate(age_group = case_when(
    age >= 0 & age <= 5   ~ "0-5 years",
    age >= 6 & age <= 9   ~ "6-9",
    age >= 10 & age <= 13 ~ "10-13",
    age >= 14 & age <= 17 ~ "14-17",
    age >= 18 & age <= 24 ~ "18-24",
    age >= 25 & age <= 34 ~ "25-34",
    age >= 35 & age <= 44 ~ "35-44",
    age >= 45 & age <= 54 ~ "45-54",
    age >= 55 & age <= 64 ~ "55-64",
    age >= 65 & age <= 74 ~ "65-74",
    age >= 75             ~ "75+ years",
    TRUE ~ "Unknown"
  )) %>%
  # Ensure the age groups appear in the correct order
  mutate(age_group = factor(age_group, levels = c(
    "0-5 years", "6-9", "10-13", "14-17", "18-24", "25-34", 
    "35-44", "45-54", "55-64", "65-74", "75+ years"
  ))) %>%
  filter(!is.na(gender) & !is.na(age_group)) %>%
  group_by(age_group, gender) %>%
  summarise(total_adj_weight = sum(adjusted_Weight, na.rm = TRUE), .groups = "drop") 

# Format for the "genderage_pyramid" sheet
pyramid_summary <- current_pyramid_raw %>%
  pivot_wider(names_from = gender, values_from = total_adj_weight, values_fill = 0) %>%
  rename(Males = Male, Females = Female) %>%
  relocate(Males, .before = Females)

# --- 2. DEFINE PREVIOUS PERIOD DATA (2024) ---

prev_data_df <- tibble(
  age_group = c("0-5 years", "6-9", "10-13", "14-17", "18-24", "25-34", 
                "35-44", "45-54", "55-64", "65-74", "75+ years"),
  Prev_Males_Val = c(45.18, 44.49, 55.11, 56.27, 51.15, 74.22, 102.47, 74.89, 43.17, 32.73, 5.04),
  Prev_Females_Val = c(40.46, 47.33, 61.20, 50.26, 89.09, 106.70, 186.96, 116.31, 73.18, 39.31, 4.78)
)

# --- 3. CALCULATE DIFFERENCE TABLE & P-VALUES ---

# Calculate Grand Totals (Sum of all weights in the sample)
prev_total_weight <- sum(prev_data_df$Prev_Males_Val) + sum(prev_data_df$Prev_Females_Val)
curr_total_weight <- sum(current_pyramid_raw$total_adj_weight)

# Helper function for Two-Sample Z-Test for Proportions
# Inputs: x1 (count/weight group 1), n1 (total 1), x2 (count/weight group 2), n2 (total 2)
calculate_z_pval <- function(x1, n1, x2, n2) {
  p1 <- x1 / n1
  p2 <- x2 / n2
  p_pool <- (x1 + x2) / (n1 + n2)
  
  # Standard Error
  se <- sqrt(p_pool * (1 - p_pool) * ((1/n1) + (1/n2)))
  
  # Z-score
  z <- (p1 - p2) / se
  
  # Two-tailed P-value
  pval <- 2 * pnorm(-abs(z))
  return(pval)
}

# Prepare Current Data with Shares
current_shares <- current_pyramid_raw %>%
  pivot_wider(names_from = gender, values_from = total_adj_weight, values_fill = 0) %>%
  rename(Curr_Males_Val = Male, Curr_Females_Val = Female) %>%
  mutate(
    Curr_Male_Share = Curr_Males_Val / curr_total_weight,
    Curr_Female_Share = Curr_Females_Val / curr_total_weight
  )

# Prepare Previous Data with Shares
prev_shares <- prev_data_df %>%
  mutate(
    Prev_Male_Share = Prev_Males_Val / prev_total_weight,
    Prev_Female_Share = Prev_Females_Val / prev_total_weight
  )

# Merge, Calculate Logic for Chart, and P-Values
diff_analysis <- full_join(prev_shares, current_shares, by = "age_group") %>%
  mutate(
    # --- CHART LOGIC ---
    # WOMEN
    fem_base = pmin(Prev_Female_Share, Curr_Female_Share),
    Women_increase_pct = if_else(Curr_Female_Share > Prev_Female_Share, 
                                 Curr_Female_Share - Prev_Female_Share, 0),
    Women_decrease_pct = if_else(Prev_Female_Share > Curr_Female_Share, 
                                 Prev_Female_Share - Curr_Female_Share, 0),
    
    # MEN (Negative Side)
    male_base = pmin(Prev_Male_Share, Curr_Male_Share) * -1,
    Men_increase_pct = if_else(Curr_Male_Share > Prev_Male_Share, 
                               (Curr_Male_Share - Prev_Male_Share) * -1, 0),
    Men_decrease_pct = if_else(Prev_Male_Share > Curr_Male_Share, 
                               (Prev_Male_Share - Curr_Male_Share) * -1, 0),
    
    # --- NET SHARES (No *100, kept as decimal) ---
    net_male = Curr_Male_Share,
    net_female = Curr_Female_Share,
    
    # --- STATISTICAL SIGNIFICANCE (P-VALUES) ---
    # Comparing 2024 vs 2025 proportions for Males
    pval_male = mapply(calculate_z_pval, 
                       Prev_Males_Val, prev_total_weight, 
                       Curr_Males_Val, curr_total_weight),
    
    # Comparing 2024 vs 2025 proportions for Females
    pval_female = mapply(calculate_z_pval, 
                         Prev_Females_Val, prev_total_weight, 
                         Curr_Females_Val, curr_total_weight),
    
    # --- FORMATTING SEPARATORS ---
    sep1 = NA,
    sep2 = NA,
    sep3 = NA
    
  ) %>%
  select(
    age_group,
    `Men, decrease %` = Men_decrease_pct,
    `Men, increase %` = Men_increase_pct,
    male_base,
    fem_base,
    `Women, increase %` = Women_increase_pct,
    `Women, decrease %` = Women_decrease_pct,
    net_male,
    net_female,
    sep1,                 # Empty Col 1
    `2024 Male%` = Prev_Male_Share,
    `2024 Female%` = Prev_Female_Share,
    sep2,                 # Empty Col 2
    `2025 Male%` = Curr_Male_Share,
    `2025 Female%` = Curr_Female_Share,
    sep3,                 # Empty Col 3
    `P-value Male` = pval_male,
    `P-value Female` = pval_female
  )

# --- SAVE TO EXCEL ---

sheets_to_save <- list(
  "raw_data" = final_data,
  "genderage_pyramid" = pyramid_summary,
  "diff_graph_data" = diff_analysis
)

write_xlsx(sheets_to_save, output_file)
