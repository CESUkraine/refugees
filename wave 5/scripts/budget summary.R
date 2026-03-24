library(survey)
library(tidyverse)
library(writexl)

# 1. Identify the percentage columns (E4.2 set)
e42_cols <- names(data)[grepl("^E4.2_", names(data))]

# 2. Clean Data
# We replace NAs with 0 (if they don't have that budget item, it's 0%)
data_budget <- data %>%
  mutate(across(all_of(e42_cols), ~ replace_na(as.numeric(.), 0)))

# 3. Define Survey Design
s_design <- svydesign(ids = ~1, weights = ~Weight_2, data = data_budget)

# 4. Calculate Weighted Mean for each Budget Component
# This gives the "Average % of budget" across the population
budget_means <- svymean(as.formula(paste0("~", paste(e42_cols, collapse = "+"))), s_design)

# 5. Format the Summary Table
budget_summary <- as.data.frame(budget_means) %>%
  rownames_to_column(var = "Budget_Item_Code") %>%
  rename(Mean_Contribution = mean) %>%
  mutate(
    # Clean up names
    Budget_Item_Code = str_remove(Budget_Item_Code, "mean\\."),
    # Convert to 0-100 scale
    Weighted_Percentage_of_Total_Budget = round(Mean_Contribution, 1)
  )

# 6. Re-scaling to 100% 
# Sometimes due to rounding or survey "noise", the total is 99% or 101%.
# We normalize it so the parts equal exactly 100% of the "Average Budget"
total_sum <- sum(budget_summary$Weighted_Percentage_of_Total_Budget)
budget_summary <- budget_summary %>%
  mutate(Normalized_Percentage = round((Weighted_Percentage_of_Total_Budget / total_sum) * 100, 1))

# 7. Add Labels (Pulling from metadata if available)
budget_summary$Label <- map_chr(budget_summary$Budget_Item_Code, function(col) {
  lbl <- attr(data[[col]], "label")
  if (is.null(lbl)) return(col) else return(lbl)
})

# 8. Export to Excel
write_xlsx(list(Budget_Composition = budget_summary), "Weighted_Budget_Analysis.xlsx")

# --- Print Results ---
print(budget_summary %>% select(Label, Normalized_Percentage))