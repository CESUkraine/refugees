library(haven)
library(writexl)
library(dplyr)
library(tidyr)
library(stringr)

# 1. Read data
input_file <- "901849.sav"
data <- read_sav(input_file, encoding = "cp1251")

# --- CHECK FOR MULTIPLE ANSWERS (Optional - for your info) ---
multi_check <- data %>%
  mutate(
    K1_count = rowSums(across(starts_with("K1_")), na.rm = TRUE),
    K2_count = rowSums(across(starts_with("K2_")), na.rm = TRUE)
  )

cat("Respondents with multiple answers in K1:", sum(multi_check$K1_count > 1))
cat("\nRespondents with multiple answers in K2:", sum(multi_check$K2_count > 1), "\n")


# --- 2. PROCESS K1 (Mutually Exclusive Priority) ---
clean_data <- data %>%
  mutate(
    K1_Priority_Label = case_when(
      K1_1 == 1 ~ "Вже є інші документи",
      K1_2 == 1 ~ "Планую оформити інші",
      K1_3 == 1 ~ "Залишуся на тимч. захисті",
      K1_4 == 1 ~ "Мало знаю, почекаю",
      # MERGED HERE: Both 5 and 6 become "Інше/ВВ"
      K1_5 == 1 ~ "Інше/ВВ",
      K1_6 == 1 ~ "Інше/ВВ",
      TRUE ~ NA_character_
    )
  ) %>%
  filter(!is.na(K1_Priority_Label))

# --- 3. PROCESS K2 (Keep Multiple Answers) ---
k2_long <- clean_data %>%
  select(ID, K1_Priority_Label, Weight_2, starts_with("K2_")) %>%
  pivot_longer(
    cols = starts_with("K2_"),
    names_to = "K2_Code_Raw",
    values_to = "selected"
  ) %>%
  filter(selected == 1) %>%
  mutate(
    K2_Label = recode(K2_Code_Raw,
      "K2_1" = "Шукатиму дозвіл в іншій країні",
      "K2_2" = "Залишуся тут без дозволу",
      "K2_3" = "Переїду в іншу без дозволу",
      "K2_4" = "Повернуся в Україну",
      "K2_5" = "Тимчасово в Україну, потім виїду",
      # MERGED HERE: Both 6 and 7 become "Інше/ВВ"
      "K2_6" = "Інше/ВВ",
      "K2_7" = "Інше/ВВ"
    )
  ) %>%
  # CRITICAL STEP: Deduplicate
  # If someone selected BOTH "Other" and "Hard to say", they now have two rows
  # with the label "Інше/ВВ". We must keep only one row per person per label
  # to avoid double-counting the weight.
  distinct(ID, K2_Label, .keep_all = TRUE)

# --- 4. HANDLE LOGIC FOR SKIPPED K2 ---
final_long <- clean_data %>%
  select(ID, K1_Priority_Label, Weight_2) %>%
  left_join(k2_long %>% select(ID, K2_Label), by = "ID") %>%
  mutate(
    # If K2 is NA (meaning they skipped it because of logic in K1),
    # we assign them to "Вже є інші документи" (or whatever logic fits your survey flow)
    K2_Label = ifelse(is.na(K2_Label), "Вже є інші документи", K2_Label)
  )

# --- 5. CALCULATE STATISTICS ---

# A. Denominators (Total Weight per K1 Group)
k1_totals <- clean_data %>%
  group_by(K1_Priority_Label) %>%
  summarise(Total_K1_Weight = sum(Weight_2, na.rm = TRUE))

# B. Numerators (Weight per Intersection) & Calculation
k1_k2_cross <- final_long %>%
  group_by(K1_Priority_Label, K2_Label) %>%
  summarise(
    sum_Weight_2 = sum(Weight_2, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  left_join(k1_totals, by = "K1_Priority_Label") %>%
  mutate(
    Percentage = (sum_Weight_2 / Total_K1_Weight) * 100
  )

# View
print(k1_k2_cross)

# Export
write_xlsx(k1_k2_cross, "K1_K2_CrossSection_v2.xlsx")
