library(haven)
library(dplyr)
library(tidyr)
library(writexl)

# 1. Read data
input_file <- "901849.sav"
data <- read_sav(input_file, encoding = "cp1251")

# --- STEP 1: LOGIC DEFINITION & CLEANING ---

# Helper to handle NA in multiple choice (convert NA to 0)
clean_mc <- function(x) {
    if_else(is.na(x), 0, x)
}

# Process Adult Data
processed_adults <- data %>%
    mutate(
        # 1. Clean K1 (Documents)
        K1_1_clean = clean_mc(K1_1),

        # 2. Clean K2 (Plans if TPS cancelled)
        # Note: K2 is usually NA if K1_1 == 1. We treat NA as 0 for logic checks.
        K2_1_clean = clean_mc(K2_1), # Search permit elsewhere
        K2_2_clean = clean_mc(K2_2), # Stay illegal
        K2_3_clean = clean_mc(K2_3), # Move illegal
        K2_4_clean = clean_mc(K2_4), # Return Ukraine
        K2_5_clean = clean_mc(K2_5), # Temp Return
        K2_6_clean = clean_mc(K2_6), # Other
        K2_7_clean = clean_mc(K2_7), # Hard to say

        # 3. DEFINE CATEGORIES

        # A. SECURE: Already have other documents
        IS_SECURE = if_else(K1_1_clean == 1, 1, 0),

        # B. VULNERABLE: Do not have documents
        IS_VULNERABLE = 1 - IS_SECURE,

        # C. BROAD RETURN: Vulnerable AND (Return OR Temp Return)
        WILL_RETURN_BROAD = if_else(
            IS_VULNERABLE == 1 & (K2_4_clean == 1 | K2_5_clean == 1),
            1, 0
        ),

        # D. STRICT RETURN: Vulnerable AND (Return Only)
        # Sum of all "Non-Return" K2 options (Search, Stay, Move, Temp, Other, IDK)
        k2_others_sum = K2_1_clean + K2_2_clean + K2_3_clean + K2_5_clean + K2_6_clean + K2_7_clean,
        WILL_RETURN_STRICT = if_else(
            IS_VULNERABLE == 1 & K2_4_clean == 1 & k2_others_sum == 0,
            1, 0
        ),

        # E. STAY/SEARCH/MOVE (Vulnerable but NOT returning)
        WILL_STAY_OR_OTHER = if_else(
            IS_VULNERABLE == 1 & WILL_RETURN_BROAD == 0,
            1, 0
        ),

        # 4. DEMOGRAPHICS (Adults)

        # Age Group
        Demo_Age_Group = as_factor(S2a),

        # Education (Grouped)
        Demo_Education = case_when(
            Z4 %in% c(1, 2) ~ "Secondary (Середня)",
            Z4 == 3 ~ "Vocational (Середня спец.)",
            Z4 == 4 ~ "Incomplete Higher (Незакінчена вища)",
            Z4 %in% c(5, 6, 7) ~ "Higher (Вища)",
            TRUE ~ NA_character_
        ),

        # Employment (Check if any E2 variable is 1)
        work_check = if_any(
            any_of(c("E2_01", "E2_02", "E2_04", "E2_05", "E2_10", "E2_11")),
            ~ !is.na(.) & . == 1
        ),
        Demo_Employment = if_else(work_check, "Working", "Non-Working")
    )

# --- STEP 2: PREPARE DATASETS ---

# 1. ADULTS
df_adults <- processed_adults %>%
    select(
        Weight_2, Demo_Age_Group, Demo_Education, Demo_Employment,
        IS_SECURE, IS_VULNERABLE, WILL_RETURN_BROAD, WILL_RETURN_STRICT, WILL_STAY_OR_OTHER
    ) %>%
    mutate(
        Segment = "Adults (Respondents)",
        final_weight = Weight_2
    )

# 2. CHILDREN (Inherit Parent's K1/K2 status)
df_children <- processed_adults %>%
    select(
        Weight_2, A2.1_6, B1, starts_with("B2_"),
        IS_SECURE, IS_VULNERABLE, WILL_RETURN_BROAD, WILL_RETURN_STRICT, WILL_STAY_OR_OTHER
    ) %>%
    pivot_longer(
        cols = matches("^B2_\\d+$"),
        names_to = "child_index",
        values_to = "child_age"
    ) %>%
    mutate(child_num = as.numeric(sub("B2_", "", child_index))) %>%
    filter(child_num <= B1 & !is.na(child_age)) %>%
    mutate(
        # Weight Adjustment: Halve if partner exists
        has_partner = if_else(!is.na(A2.1_6) & A2.1_6 == 1, 1, 0),
        final_weight = if_else(has_partner == 1, Weight_2 / 2, Weight_2),

        # Create Child Segments
        Segment = case_when(
            child_age >= 0 & child_age <= 13 ~ "Children (0-13)",
            child_age >= 14 & child_age <= 17 ~ "Children (14-17)",
            TRUE ~ "Children (Unknown Age)"
        ),

        # Placeholders for Adult Demographics (Empty for children)
        Demo_Age_Group = NA,
        Demo_Education = NA,
        Demo_Employment = NA
    )

# --- STEP 3: AGGREGATION FUNCTION ---

# This function sums the weights for every category
calculate_weighted_sums <- function(df, group_col) {
    df %>%
        group_by(Group = .data[[group_col]]) %>%
        summarise(
            # 1. Total Population
            N_Respondents = n(),
            Weighted_Pop_Total = sum(final_weight, na.rm = TRUE),

            # 2. Documents Status
            Weighted_Secure_Docs = sum(final_weight * IS_SECURE, na.rm = TRUE),
            Weighted_Vulnerable = sum(final_weight * IS_VULNERABLE, na.rm = TRUE),

            # 3. Return Intentions (Weighted)
            Weighted_Return_Broad = sum(final_weight * WILL_RETURN_BROAD, na.rm = TRUE),
            Weighted_Return_Strict = sum(final_weight * WILL_RETURN_STRICT, na.rm = TRUE),

            # 4. Stay/Other Intentions (Weighted)
            Weighted_Stay_Other = sum(final_weight * WILL_STAY_OR_OTHER, na.rm = TRUE)
        ) %>%
        # Calculate % of Total Population for quick reference
        mutate(
            Pct_Secure = Weighted_Secure_Docs / Weighted_Pop_Total,
            Pct_Return_Strict = Weighted_Return_Strict / Weighted_Pop_Total
        )
}

# --- STEP 4: GENERATE TABLES ---

# A. General Segmentation (Adults + Children Groups)
# Combine just for this table
df_all_segments <- bind_rows(
    df_adults %>% select(Segment, final_weight, starts_with("IS_"), starts_with("WILL_")),
    df_children %>% select(Segment, final_weight, starts_with("IS_"), starts_with("WILL_"))
)
table_segments <- calculate_weighted_sums(df_all_segments, "Segment")

# B. Adult Demographics
table_adults_age <- calculate_weighted_sums(df_adults %>% filter(!is.na(Demo_Age_Group)), "Demo_Age_Group")
table_adults_edu <- calculate_weighted_sums(df_adults %>% filter(!is.na(Demo_Education)), "Demo_Education")
table_adults_emp <- calculate_weighted_sums(df_adults %>% filter(!is.na(Demo_Employment)), "Demo_Employment")

# --- STEP 5: EXPORT ---

print("--- TPS Cancellation Analysis Generated ---")
print(table_segments)

write_xlsx(
    list(
        "1_Segments_Overview" = table_segments,
        "2_Adults_Age"        = table_adults_age,
        "3_Adults_Education"  = table_adults_edu,
        "4_Adults_Employment" = table_adults_emp
    ),
    "TPS_Cancellation_Impact_Analysis.xlsx"
)
