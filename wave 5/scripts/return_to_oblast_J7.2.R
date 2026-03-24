library(haven)
library(dplyr)
library(tidyr)
library(writexl)

# 1. Read data
input_file <- "901849.sav"
data <- read_sav(input_file, encoding = "cp1251")

data$J7.2_01 # тільки додому
data$J7.2_02 # вінницька обл
data$J7.2_26 # важко сказати

library(haven)
library(dplyr)
library(stringr)

# --- 1. Preparation (Same as before) ---
# Ensure weights are numeric
data$Weight_final <- as.numeric(data$Weight_2)
data$Weight_final[is.na(data$Weight_final)] <- 0

# Ensure "Only Home" is 0/1
data$only_home <- as.numeric(data$J7.2_01)
data$only_home[is.na(data$only_home)] <- 0

# Map J7.2 columns
j7_map <- data.frame(col_name = character(), col_label = character(), stringsAsFactors = FALSE)
for (i in 1:26) {
  col_name <- sprintf("J7.2_%02d", i)
  if (col_name %in% names(data)) {
    lbl <- attr(data[[col_name]], "label")
    if (is.null(lbl)) lbl <- ""
    j7_map <- rbind(j7_map, data.frame(col_name = col_name, col_label = trimws(lbl)))
  }
}

# --- 2. Master List of Regions ---
z1_labels <- as_factor(data$Z1) %>% levels()
z1_codes  <- 1:length(z1_labels)

results <- list()

# --- 3. The Modified Loop ---
for (i in seq_along(z1_codes)) {
  
  region_code <- z1_codes[i]
  region_name <- z1_labels[i]
  
  # --- MERGE LOGIC START ---
  
  # SKIP Code 11 (Kyiv City) entirely, as we will handle it inside Code 10
  if (region_code == 11) next 
  
  # If we are at Code 10 (Kyiv Oblast), we merge 10 and 11
  if (region_code == 10) {
    
    # Update Name
    region_name <- "Київ та Київська область"
    
    # 1. Direct Choice: Uses J7.2_10 (Kyiv Oblast)
    # We find the column mapped to the original "Київська область" label
    target_col_row <- j7_map %>% filter(str_detect(col_label, "Київська"))
    
    direct_choice_vector <- rep(0, nrow(data))
    if (nrow(target_col_row) > 0) {
      vals <- as.numeric(data[[target_col_row$col_name[1]]])
      vals[is.na(vals)] <- 0
      direct_choice_vector <- vals
    }
    
    # 2. Home Choice: Z1 is 10 OR 11, AND Only Home is clicked
    # We allow Z1 to be 10 or 11 here
    home_choice_vector <- ifelse(data$only_home == 1 & data$Z1 %in% c(10, 11), 1, 0)
    
    mapped_col_name <- paste(target_col_row$col_name[1], "(Merged)")
    
  } else {
    # --- STANDARD LOGIC (For all other regions) ---
    
    # Match by label
    match_row <- j7_map %>% filter(col_label == trimws(region_name))
    
    direct_choice_vector <- rep(0, nrow(data))
    if (nrow(match_row) > 0) {
      target_col <- match_row$col_name[1]
      vals <- as.numeric(data[[target_col]])
      vals[is.na(vals)] <- 0
      direct_choice_vector <- vals
      mapped_col_name <- target_col
    } else {
      mapped_col_name <- "None (Only via Z1)"
    }
    
    # Home Choice: Strict matching (Z1 == region_code)
    home_choice_vector <- ifelse(data$only_home == 1 & as.numeric(data$Z1) == region_code, 1, 0)
  }
  
  # --- 4. Calculate Share ---
  final_indicator <- ifelse(direct_choice_vector == 1 | home_choice_vector == 1, 1, 0)
  
  weighted_n <- sum(final_indicator * data$Weight_final, na.rm = TRUE)
  total_n    <- sum(data$Weight_final, na.rm = TRUE)
  share      <- weighted_n / total_n
  
  results[[length(results) + 1]] <- data.frame(
    Z1_Code = region_code,
    Region_Name = region_name,
    Mapped_Column = mapped_col_name,
    Share = share
  )
}

# --- 5. Final Table ---
final_table <- do.call(rbind, results)
final_table$Percentage <- round(final_table$Share * 100, 2)

print(final_table)

write_xlsx(final_table, "return_to_oblast_J7.2.xlsx")
