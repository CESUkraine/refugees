library(haven)
library(writexl)
library(dplyr)
library(tidyr)

# 1. Read your .sav file
input_file <- "901849_separate_subsample_weights.sav"
# input_file <- "901737_wave4.sav"
data <- read_sav(input_file, encoding = "cp1251")

# 2. Convert J11_ and B3_ columns to human-readable factors
data_converted <- data %>%
  mutate(
    across(starts_with("J11_"), haven::as_factor), # does your child want to return to UA or stay abroad?
    across(starts_with("B3_"), haven::as_factor) # child's gender
  )

# 3. Reshape from wide to long, including ID and Weight in the selection
df_kids <- data_converted %>%
  select(
    ID, 
    Weight, 
    matches("^(B2|B3|J11)_\\d+$")  # picks up B2_1..B2_10 and J11_1..J11_10
  ) %>%
  pivot_longer(
    cols = matches("^(B2|B3|J11)_\\d+$"),
    names_to = c(".value", "child_id"),
    names_pattern = "(B2|B3|J11)_(\\d+)"
  ) %>%
  # Rename pivoted columns
  rename(
    age = B2,
    gender = B3,
    status = J11
  ) %>%
  # Filter out rows without a valid age
  filter(!is.na(age))

# 4. Select the final columns you want in your output
df_kids_final <- df_kids %>%
  select(ID, Weight, age, gender, status)

# 5. Write to xlsx
write_xlsx(df_kids_final, "children_return_age_gender.xlsx")

# for (name in colnames(data)) {
#     print(name)
#     print(attr(data[[name]], "label"))
#     print("- - - - - - - - - - -")
# }

attributes(data$Z4)
attr(data$B5_1_1, "label")
attr(data$B5_1_7, "label")
