library(haven)
library(dplyr)
library(tidyr)
library(writexl)

# 1. Read data
input_file <- "901849.sav"
data <- read_sav(input_file, encoding = "cp1251")

# --- STEP 1: LOGIC & CLEANING ---

clean_mc <- function(x) {
    if_else(is.na(x), 0, x)
}

processed_data <- data %>%
    filter(!is.na(J2)) %>% # Ensure we have the general return intention
    mutate(
        # 1. J2 Labeling
        J2_Label = case_when(
            J2 == 1 ~ "1. Definitely Plan to Return",
            J2 == 2 ~ "2. Rather Plan to Return",
            J2 == 3 ~ "3. Rather NOT Plan to Return",
            J2 == 4 ~ "4. Definitely NOT Plan to Return",
            J2 == 5 ~ "5. Hard to Say",
            TRUE ~ "Unknown"
        ),

        # 2. TPS Cancellation Logic (K1/K2)
        K1_1_clean = clean_mc(K1_1),
        K2_4_clean = clean_mc(K2_4), # Return
        K2_5_clean = clean_mc(K2_5), # Temp Return

        # Sum of non-return options in K2
        k2_others_sum = clean_mc(K2_1) + clean_mc(K2_2) + clean_mc(K2_3) +
            clean_mc(K2_5) + clean_mc(K2_6) + clean_mc(K2_7),

        # A. SECURE
        IS_SECURE = if_else(K1_1_clean == 1, 1, 0),

        # B. VULNERABLE
        IS_VULNERABLE = 1 - IS_SECURE,

        # C. BROAD RETURN (Vulnerable + (Return OR Temp Return))
        WILL_RETURN_BROAD = if_else(
            IS_VULNERABLE == 1 & (K2_4_clean == 1 | K2_5_clean == 1),
            1, 0
        ),

        # D. STRICT RETURN (Vulnerable + Return ONLY)
        WILL_RETURN_STRICT = if_else(
            IS_VULNERABLE == 1 & K2_4_clean == 1 & k2_others_sum == 0,
            1, 0
        ),

        # E. STAY/OTHER (Vulnerable + NOT Returning)
        WILL_STAY_OR_OTHER = if_else(
            IS_VULNERABLE == 1 & WILL_RETURN_BROAD == 0,
            1, 0
        )
    )

# --- STEP 2: PREPARE ADULTS & CHILDREN ---

# 1. ADULTS
df_adults <- processed_data %>%
    select(
        Weight_2, J2_Label, IS_SECURE, IS_VULNERABLE,
        WILL_RETURN_BROAD, WILL_RETURN_STRICT, WILL_STAY_OR_OTHER
    ) %>%
    mutate(
        Segment = "Adults (Respondents)",
        final_weight = Weight_2
    )

# 2. CHILDREN (0-17 Pooled)
df_children <- processed_data %>%
    select(
        Weight_2, A2.1_6, B1, starts_with("B2_"), J2_Label,
        IS_SECURE, IS_VULNERABLE, WILL_RETURN_BROAD, WILL_RETURN_STRICT, WILL_STAY_OR_OTHER
    ) %>%
    pivot_longer(
        cols = matches("^B2_\\d+$"),
        names_to = "child_index",
        values_to = "child_age"
    ) %>%
    mutate(child_num = as.numeric(sub("B2_", "", child_index))) %>%
    filter(child_num <= B1 & !is.na(child_age)) %>% # Keep valid children
    mutate(
        # Weight Adjustment
        has_partner = if_else(!is.na(A2.1_6) & A2.1_6 == 1, 1, 0),
        final_weight = if_else(has_partner == 1, Weight_2 / 2, Weight_2),
        Segment = "Children (0-17)"
    ) %>%
    select(
        Segment, final_weight, J2_Label, IS_SECURE, IS_VULNERABLE,
        WILL_RETURN_BROAD, WILL_RETURN_STRICT, WILL_STAY_OR_OTHER
    )

# --- STEP 3: AGGREGATE CROSS-TABS ---

# Combine datasets
df_combined <- bind_rows(df_adults, df_children)

# Group by Segment AND J2 Answer
crosstab_results <- df_combined %>%
    group_by(Segment, J2_Label) %>%
    summarise(
        # 1. Total Count for this J2 answer
        J2_Group_Weighted_N = sum(final_weight, na.rm = TRUE),

        # 2. Breakdown of TPS Status within this J2 answer
        Weighted_Secure_Docs = sum(final_weight * IS_SECURE, na.rm = TRUE),
        Weighted_Vulnerable = sum(final_weight * IS_VULNERABLE, na.rm = TRUE),
        Weighted_Return_Broad = sum(final_weight * WILL_RETURN_BROAD, na.rm = TRUE),
        Weighted_Return_Strict = sum(final_weight * WILL_RETURN_STRICT, na.rm = TRUE),
        Weighted_Stay_Other = sum(final_weight * WILL_STAY_OR_OTHER, na.rm = TRUE),
        .groups = "drop"
    ) %>%
    arrange(Segment, J2_Label)

# --- STEP 4: EXPORT ---

print(crosstab_results)

write_xlsx(
    list("J2_vs_TPS_Crosstab" = crosstab_results),
    "Return_Intentions_CrossTab.xlsx"
)
