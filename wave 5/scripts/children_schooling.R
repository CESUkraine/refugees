# Needed packages
library(haven)
library(writexl)
library(dplyr)
library(tidyr)
library(stringr)

# 1. Read data
input_file <- "901849.sav"
data <- read_sav(input_file, encoding = "cp1251")

# 2. Convert J11 columns to human-readable factors
data_converted <- data %>%
    mutate(across(starts_with("J11_"), haven::as_factor))

#### PART A: Pivot J11 => Get the list of "Real" Children
df_children_status <- data_converted %>%
    select(ID, Weight_2, starts_with("J11_")) %>%
    pivot_longer(
        cols         = starts_with("J11_"),
        names_to     = "child_id",
        names_prefix = "J11_",
        values_to    = "decision_status"
    ) %>%
    mutate(child_id = as.integer(child_id)) %>%
    filter(!is.na(decision_status))

#### PART B: Process the B5 Matrix (Child_Option)
df_b5_long <- data_converted %>%
    select(ID, Weight_2, matches("^B5_\\d+_\\d+$")) %>%
    pivot_longer(
        cols = matches("^B5_\\d+_\\d+$"),
        names_to = c("child_id", "option_no"),
        names_pattern = "B5_(\\d+)_(\\d+)",
        values_to = "value"
    ) %>%
    mutate(
        child_id = as.integer(child_id),
        option_no = as.integer(option_no)
    )

#### PART C: Join Status & Schooling Data
# Ensures we only keep B5 data for children that actually exist in J11
df_joined <- df_children_status %>%
    inner_join(df_b5_long, by = c("ID", "Weight_2", "child_id"))

#### PART D: Pivot Back to Wide (Columns B5_1...B5_7)
# Creates one row per child
df_wide <- df_joined %>%
    pivot_wider(
        names_from = option_no,
        values_from = value,
        names_prefix = "B5_"
    ) %>%
    drop_na(starts_with("B5_")) %>%
    select(
        ID,
        Weight_2,
        child_id,
        decision_status,
        num_range("B5_", 1:7)
    )

#### PART E: Create Textual Combination Column
# Creates strings like "B5_1" or "B5_1 & B5_3"
b5_cols <- paste0("B5_", 1:7)

# We use apply to go row-by-row
df_final <- df_wide
df_final$education_combination <- apply(df_final[, b5_cols], 1, function(x) {
    # Find which columns have a 1
    active_cols <- names(x)[which(x == 1)]

    if (length(active_cols) == 0) {
        return("None")
    } else {
        return(paste(active_cols, collapse = " & "))
    }
})

#### PART F: Create the Weighted Pivot Table
# Instead of count(), we sum the weights.
df_pivot_table <- df_final %>%
    group_by(education_combination, decision_status) %>%
    summarise(
        weighted_count = sum(Weight_2, na.rm = TRUE), # THE FIX: Summing weights
        .groups = "drop"
    ) %>%
    pivot_wider(
        names_from = decision_status,
        values_from = weighted_count,
        values_fill = 0
    )

#### PART G: Write Both Sheets to XLSX
write_xlsx(
    list(
        "Data_Per_Child" = df_final,
        "Weighted_Summary" = df_pivot_table
    ),
    "children_schooling_return.xlsx"
)
