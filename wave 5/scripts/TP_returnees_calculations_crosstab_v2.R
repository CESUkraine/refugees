library(haven)
library(dplyr)
library(tidyr)
library(writexl)

# 1. Read data
input_file <- "901849.sav"
data <- read_sav(input_file, encoding = "cp1251")

# --- STEP 1: LOGIC DEFINITION & CLEANING ---

clean_mc <- function(x) {
    if_else(is.na(x), 0, x)
}

processed_adults <- data %>%
    mutate(
        # 1. Clean K1
        K1_1_clean = clean_mc(K1_1),

        # 2. Clean K2
        K2_1_clean = clean_mc(K2_1), # Search permit
        K2_2_clean = clean_mc(K2_2), # Stay illegal
        K2_3_clean = clean_mc(K2_3), # Move illegal
        K2_4_clean = clean_mc(K2_4), # Return Ukraine
        K2_5_clean = clean_mc(K2_5), # Temp Return
        K2_6_clean = clean_mc(K2_6), # Other
        K2_7_clean = clean_mc(K2_7), # Hard to say

        # 3. DEFINE CATEGORIES

        # A. SECURE
        IS_SECURE = if_else(K1_1_clean == 1, 1, 0),

        # B. VULNERABLE
        IS_VULNERABLE = 1 - IS_SECURE,

        # C. BROAD RETURN (Any mention of Return or Temp Return)
        WILL_RETURN_BROAD = if_else(
            IS_VULNERABLE == 1 & (K2_4_clean == 1 | K2_5_clean == 1),
            1, 0
        ),

        # D. STRICT RETURN (Return Only - purely for sub-analysis)
        k2_others_sum = K2_1_clean + K2_2_clean + K2_3_clean + K2_5_clean + K2_6_clean + K2_7_clean,
        WILL_RETURN_STRICT = if_else(
            IS_VULNERABLE == 1 & K2_4_clean == 1 & k2_others_sum == 0,
            1, 0
        ),

        # --- MODIFICATION START ---

        # E. UNCERTAIN
        # Logic: Vulnerable AND K2_7 (Hard to say) is 1 AND All other K2 options are 0
        k2_substantive_sum = K2_1_clean + K2_2_clean + K2_3_clean + K2_4_clean + K2_5_clean + K2_6_clean,
        WILL_UNCERTAIN = if_else(
            IS_VULNERABLE == 1 & K2_7_clean == 1 & k2_substantive_sum == 0,
            1, 0
        ),

        # F. WILL STAY (The remainder)
        # Logic: Vulnerable AND Not Returning AND Not Uncertain
        # This captures: Search Permit, Stay Illegal, Move to 3rd Country, Other
        WILL_STAY = if_else(
            IS_VULNERABLE == 1 & WILL_RETURN_BROAD == 0 & WILL_UNCERTAIN == 0,
            1, 0
        ),

        # --- MODIFICATION END ---

        # 4. DEMOGRAPHICS
        Demo_Age_Group = as_factor(S2a),
        Demo_Education = case_when(
            Z4 %in% c(1, 2) ~ "Secondary (Середня)",
            Z4 == 3 ~ "Vocational (Середня спец.)",
            Z4 == 4 ~ "Incomplete Higher (Незакінчена вища)",
            Z4 %in% c(5, 6, 7) ~ "Higher (Вища)",
            TRUE ~ NA_character_
        ),
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
        IS_SECURE, IS_VULNERABLE,
        WILL_RETURN_BROAD, WILL_RETURN_STRICT,
        WILL_STAY, WILL_UNCERTAIN
    ) %>%
    mutate(
        Segment = "Adults (Respondents)",
        final_weight = Weight_2
    )

# 2. CHILDREN
df_children <- processed_adults %>%
    select(
        Weight_2, A2.1_6, B1, starts_with("B2_"),
        IS_SECURE, IS_VULNERABLE,
        WILL_RETURN_BROAD, WILL_RETURN_STRICT,
        WILL_STAY, WILL_UNCERTAIN
    ) %>%
    pivot_longer(
        cols = matches("^B2_\\d+$"),
        names_to = "child_index",
        values_to = "child_age"
    ) %>%
    mutate(child_num = as.numeric(sub("B2_", "", child_index))) %>%
    filter(child_num <= B1 & !is.na(child_age)) %>%
    mutate(
        has_partner = if_else(!is.na(A2.1_6) & A2.1_6 == 1, 1, 0),
        final_weight = if_else(has_partner == 1, Weight_2 / 2, Weight_2),
        Segment = case_when(
            child_age >= 0 & child_age <= 13 ~ "Children (0-13)",
            child_age >= 14 & child_age <= 17 ~ "Children (14-17)",
            TRUE ~ "Children (Unknown Age)"
        ),
        Demo_Age_Group = NA, Demo_Education = NA, Demo_Employment = NA
    )

# --- STEP 3: AGGREGATION FUNCTION ---

calculate_weighted_sums <- function(df, group_col) {
    df %>%
        group_by(Group = .data[[group_col]]) %>%
        summarise(
            N_Respondents = n(),
            Weighted_Pop_Total = sum(final_weight, na.rm = TRUE),

            # Documents
            Weighted_Secure_Docs = sum(final_weight * IS_SECURE, na.rm = TRUE),
            Weighted_Vulnerable = sum(final_weight * IS_VULNERABLE, na.rm = TRUE),

            # Return
            Weighted_Return_Broad = sum(final_weight * WILL_RETURN_BROAD, na.rm = TRUE),
            # Note: Return Strict is a subset of Broad, not a separate main bucket
            Weighted_Return_Strict = sum(final_weight * WILL_RETURN_STRICT, na.rm = TRUE),

            # New Categories
            Weighted_Will_Stay = sum(final_weight * WILL_STAY, na.rm = TRUE),
            Weighted_Uncertain = sum(final_weight * WILL_UNCERTAIN, na.rm = TRUE)
        ) %>%
        mutate(
            Pct_Secure = Weighted_Secure_Docs / Weighted_Pop_Total,
            Pct_Will_Stay = Weighted_Will_Stay / Weighted_Pop_Total
        )
}

# --- STEP 4: GENERATE TABLES ---

# A. General Segmentation
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

print("--- TPS Cancellation Analysis Generated (New Logic) ---")
print(table_segments)

write_xlsx(
    list(
        "1_Segments_Overview" = table_segments,
        "2_Adults_Age"        = table_adults_age,
        "3_Adults_Education"  = table_adults_edu,
        "4_Adults_Employment" = table_adults_emp
    ),
    "TPS_Cancellation_Impact_Analysis_v2.xlsx"
)
