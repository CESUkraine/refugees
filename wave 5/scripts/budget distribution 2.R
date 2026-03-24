library(survey)
library(tidyverse)
library(writexl)

# 1. Define the sets
e41_cols <- names(data)[grepl("^E4.1_", names(data))] # Income Sources (Width)
e42_cols <- names(data)[grepl("^E4.2_", names(data))] # Budget Percentages (Height)

# 2. Prepare Data
data_clean <- data %>%
  mutate(across(all_of(c(e41_cols, e42_cols)), ~ replace_na(as.numeric(.), 0)))

# 3. Define Survey Design
s_design <- svydesign(ids = ~1, weights = ~Weight_2, data = data_clean)

# 4. Calculate Mekko Data for each E4.1 Income Group
mekko_results <- map(e41_cols, function(income_src) {
  
  # Filter design for people who have THIS income source (E4.1_x == 1)
  sub_design <- subset(s_design, get(income_src) == 1)
  
  # Calculate Total Weighted Population for this group (X-axis Width)
  pop_weight <- sum(weights(sub_design, "sampling"))
  
  # Calculate Weighted Mean for each budget item (Y-axis Heights)
  budget_means <- svymean(as.formula(paste0("~", paste(e42_cols, collapse = "+"))), sub_design)
  
  # Combine into a single row
  lbl <- attr(data[[income_src]], "label")
  if (is.null(lbl)) lbl <- income_src
  
  as.data.frame(t(as.vector(budget_means))) %>%
    set_names(e42_cols) %>%
    mutate(
      Income_Source = lbl,
      Population_Weight = pop_weight
    ) %>%
    select(Income_Source, Population_Weight, everything())
}) %>% bind_rows()

# 5. Export
write_xlsx(list(Budget_Mekko_Data = mekko_results), "Income_Source_Budget_Mekko.xlsx")

# --- Print ---
print(mekko_results)