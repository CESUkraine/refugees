library(haven)
library(dplyr)
library(tidyr)
library(writexl)

# 1. Read data
input_file <- "901849.sav"
data <- read_sav(input_file, encoding = "cp1251")

# --- STEP 1: DEFINE SCENARIO PROBABILITIES ---
# Mapping J2 answer to probabilities of returning
prob_table <- tibble(
    J2 = c(1, 2, 5, 3, 4),
    Prob_Return_Opt = c(1.00, 0.75, 0.50, 0.25, 0.00),
    Prob_Return_Mid = c(0.80, 0.60, 0.30, 0.10, 0.00),
    Prob_Return_Pes = c(0.60, 0.50, 0.20, 0.10, 0.00)
)

# --- HELPER FUNCTION ---
# Calculates Return vs Stay shares for all 3 scenarios
calc_scenario_shares <- function(df, group_var_name) {
    df %>%
        group_by(Group = .data[[group_var_name]]) %>%
        summarise(
            N_Weighted = sum(final_weight, na.rm = TRUE),

            # 1. OPTIMISTIC SCENARIO
            Opt_Return_Pct = sum(final_weight * Prob_Return_Opt, na.rm = TRUE) / sum(final_weight, na.rm = TRUE),
            Opt_Stay_Pct = 1 - Opt_Return_Pct,

            # 2. MIDDLE SCENARIO
            Mid_Return_Pct = sum(final_weight * Prob_Return_Mid, na.rm = TRUE) / sum(final_weight, na.rm = TRUE),
            Mid_Stay_Pct = 1 - Mid_Return_Pct,

            # 3. PESSIMISTIC SCENARIO
            Pes_Return_Pct = sum(final_weight * Prob_Return_Pes, na.rm = TRUE) / sum(final_weight, na.rm = TRUE),
            Pes_Stay_Pct = 1 - Pes_Return_Pct
        ) %>%
        mutate(across(ends_with("Pct"), ~ round(., 4)))
}

# --- PREPARE BASE DATA ---

# Filter respondents with valid J2 and join probabilities.
# base_data retains all original columns (including E2_xx)
base_data <- data %>%
    filter(!is.na(J2)) %>%
    left_join(prob_table, by = "J2")

# --- TABLE 1: POPULATION SEGMENTS (Adults, Children 0-13, Children 14-17) ---

# 1A. Adults
t1_adults <- base_data %>%
    select(Weight_2, starts_with("Prob_Return")) %>%
    mutate(
        Segment = "Adults (Respondents)",
        final_weight = Weight_2
    )

# 1B. Children
t1_children <- base_data %>%
    select(Weight_2, A2.1_6, B1, starts_with("B2_"), starts_with("Prob_Return")) %>%
    pivot_longer(
        cols = matches("^B2_\\d+$"),
        names_to = "child_index",
        values_to = "child_age"
    ) %>%
    mutate(child_num = as.numeric(sub("B2_", "", child_index))) %>%
    filter(child_num <= B1 & !is.na(child_age)) %>%
    mutate(
        # Weight Adjustment (Halve if partner exists)
        has_partner = if_else(!is.na(A2.1_6) & A2.1_6 == 1, 1, 0),
        final_weight = if_else(has_partner == 1, Weight_2 / 2, Weight_2),

        # Define Age Segments
        Segment = case_when(
            child_age >= 0 & child_age <= 13 ~ "Children (0-13)",
            child_age >= 14 & child_age <= 17 ~ "Children (14-17)",
            TRUE ~ "Children (Other)"
        )
    )

# 1C. Combine and Calculate
df_segments <- bind_rows(t1_adults, t1_children)
table_1_segments <- calc_scenario_shares(df_segments, "Segment")

# --- TABLE 2: ADULTS BY AGE (S2a) ---

df_age <- base_data %>%
    mutate(
        Age_Group = as_factor(S2a),
        final_weight = Weight_2
    ) %>%
    filter(!is.na(S2a))

table_2_age <- calc_scenario_shares(df_age, "Age_Group")

# --- TABLE 3: ADULTS BY EDUCATION (Z4) ---

df_edu <- base_data %>%
    mutate(
        Education = case_when(
            Z4 %in% c(1, 2) ~ "Середня (Secondary)",
            Z4 == 3 ~ "Середня спеціальна (Vocational)",
            Z4 == 4 ~ "Незакінчена вища (Unfinished higher)",
            Z4 %in% c(5, 6, 7) ~ "Вища (Higher)",
            TRUE ~ NA_character_
        ),
        final_weight = Weight_2
    ) %>%
    filter(!is.na(Education))

table_3_edu <- calc_scenario_shares(df_edu, "Education")

# --- TABLE 4: ADULTS BY EMPLOYMENT STATUS ---
# Working if E2_01, E2_02, E2_04, E2_05, E2_10, or E2_11 == 1

# List of columns to check
work_cols <- c("E2_01", "E2_02", "E2_04", "E2_05", "E2_10", "E2_11")

# Ensure we only select columns that actually exist in the dataframe
# (prevents errors if one code was never selected by anyone and missing from sav)
existing_work_cols <- intersect(work_cols, names(base_data))

df_work <- base_data %>%
    mutate(
        # Check if ANY of the columns equal 1.
        # usage of `if_any` iterates over columns. `!is.na(.)` ensures safety against missing values.
        is_working = if_any(all_of(existing_work_cols), ~ !is.na(.) & . == 1),
        Employment_Status = if_else(is_working, "Working", "Non-Working"),
        final_weight = Weight_2
    )

table_4_work <- calc_scenario_shares(df_work, "Employment_Status")

# --- EXPORT RESULTS ---

# Display in Console
print("--- Table 1: Population Segments ---")
print(table_1_segments)
print("--- Table 2: Adults by Age ---")
print(table_2_age)
print("--- Table 3: Adults by Education ---")
print(table_3_edu)
print("--- Table 4: Adults by Employment ---")
print(table_4_work)

# Save to Excel
write_xlsx(
    list(
        "Segments" = table_1_segments,
        "Adults_Age" = table_2_age,
        "Adults_Education" = table_3_edu,
        "Adults_Employment" = table_4_work
    ),
    "Return_Scenarios_Detailed.xlsx"
)
