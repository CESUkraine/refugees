library(haven)
library(writexl)
library(dplyr)
library(tidyr)

# 1. Read your .sav file
input_file <- "901737.sav"
data <- read_sav(input_file)

# 2. Convert Q17_ and B3_ columns to human-readable factors
data_converted <- data %>%
  mutate(
    across(starts_with("Q17_"), haven::as_factor),
    across(starts_with("B3_"), haven::as_factor)
  )

# 3. Reshape from wide to long, including ID and weight in the selection
df_kids <- data_converted %>%
  select(
    ID, 
    weight, 
    matches("^(B2|B3|Q17)_\\d+$")  # picks up B2_1..B2_10 and Q17_1..Q17_10
  ) %>%
  pivot_longer(
    cols = matches("^(B2|B3|Q17)_\\d+$"),
    names_to = c(".value", "child_id"),
    names_pattern = "(B2|B3|Q17)_(\\d+)"
  ) %>%
  # Rename pivoted columns
  rename(
    age = B2,
    gender = B3,
    status = Q17
  ) %>%
  # Filter out rows without a valid age
  filter(!is.na(age))

# 4. Select the final columns you want in your output
df_kids_final <- df_kids %>%
  select(ID, weight, age, gender, status)

# 5. Write to xlsx
write_xlsx(df_kids_final, "children_return_age_gender.xlsx")

for (name in colnames(data)) {
    print(name)
    print(attr(data[[name]], "label"))
    print("- - - - - - - - - - -")
}

attributes(data$Z4)
attr(data$B5_1_1, "label")
attr(data$B5_1_7, "label")
attr(data$)
