# ==============================================================================
# 0. SETUP & LIBRARIES
# ==============================================================================
# install.packages(c("tidyverse", "factoextra", "klaR", "haven", "openxlsx", "labelled"))

library(tidyverse)
library(factoextra)
library(klaR)       # For K-Modes
library(haven)      # For saving .sav files
library(openxlsx)   # For Excel reporting
library(labelled)   # For handling survey labels

# Fix for the 'select' conflict (MASS vs dplyr)
select <- dplyr::select 

# ==============================================================================
# 1. ROBUST DATA PREPARATION
# ==============================================================================

# CHECK: Ensure required objects exist
if (!exists("refugees") || !exists("refugees_clustering_ready")) {
  stop("CRITICAL ERROR: 'refugees' (raw) or 'refugees_clustering_ready' (cleaned) objects are missing. Run your cleaning script first.")
}

# 1. Safety Check: Do rows match initially?
if (nrow(refugees) != nrow(refugees_clustering_ready)) {
  warning("Row counts between raw 'refugees' and 'refugees_clustering_ready' differ. Make sure they align before proceeding.")
}

# 2. Combine ID with the Cleaning Data TEMPORARILY
temp_combined <- refugees_clustering_ready %>%
  mutate(TEMP_ID_TRACKER = refugees$ID) # Add ID from master file

# 3. Final Cleaning for Algorithm
df_clean_combined <- temp_combined %>%
  select(-any_of(c("Weight", "weight", "ID"))) %>% 
  mutate(across(-TEMP_ID_TRACKER, as.factor)) %>% 
  drop_na()

# 4. SPLIT BACK APART & CONVERT TO DATA.FRAME
df_ids <- df_clean_combined %>% select(TEMP_ID_TRACKER) %>% rename(ID = TEMP_ID_TRACKER)
df_algo <- df_clean_combined %>% select(-TEMP_ID_TRACKER) %>% as.data.frame()

print(paste("Data ready. Original rows:", nrow(refugees), "| Final rows:", nrow(df_algo)))

# ==============================================================================
# 2. DETERMINE OPTIMAL CLUSTERS (Custom Elbow for K-Modes)
# ==============================================================================

set.seed(123)
wss <- numeric()
k_range <- 1:10 

print("Calculating Elbow Plot (this may take a moment)...")

for (k in k_range) {
  k_model <- kmodes(df_algo, modes = k, iter.max = 10, weighted = FALSE)
  wss[k] <- sum(k_model$withindiff)
}

# Plot the Elbow
elbow_data <- data.frame(k = k_range, wss = wss)

print(
  ggplot(elbow_data, aes(x = k, y = wss)) +
    geom_line() +
    geom_point() +
    scale_x_continuous(breaks = k_range) +
    labs(title = "Elbow Method for K-Modes",
         subtitle = "Look for the 'bend'",
         x = "Number of Clusters (k)",
         y = "Total Within-Cluster Dissimilarity") +
    theme_minimal()
)

# ==============================================================================
# 3. RUN CLUSTERING
# ==============================================================================

# Update this number based on the plot generated above.
final_k <- 4 

set.seed(333)
print(paste0("Running final K-Modes with ", final_k, " clusters..."))
km.res <- kmodes(df_algo, modes = final_k, iter.max = 20, weighted = FALSE)

# Attach cluster labels to a temp dataframe
df_clustered_ids <- df_ids %>%
  mutate(cluster = as.factor(km.res$cluster))

print("Cluster Sizes (Unweighted):")
print(table(df_clustered_ids$cluster))

# ==============================================================================
# 4. MERGE & EXPORT
# ==============================================================================

# Merge the clusters back into the ORIGINAL full dataset
refugees_final <- refugees %>%
  left_join(df_clustered_ids, by = "ID")

# Save SPSS file
write_sav(refugees_final, "refugees_wave5_with_clusters.sav")
print("SPSS file saved: refugees_wave5_with_clusters.sav")

# ==============================================================================
# 5. GENERATE EXCEL REPORT (TRAFFIC LIGHT SIGNIFICANCE: GREEN/RED)
# ==============================================================================

# 1. DEFINE EXCEL STYLES
# ------------------------------------------------------------------------------

# Main Title Style
titleStyle <- createStyle(
  fontSize = 12, 
  fontColour = "#000000",
  textDecoration = "bold",
  fgFill = "#E7E6E6" 
)

# Super Headers
superHeaderStyle <- createStyle(
  fontSize = 11,
  fontColour = "#FFFFFF", 
  fgFill = "#44546A",      
  textDecoration = "bold",
  halign = "center",
  border = "TopBottomLeftRight"
)

# Column Headers
colHeaderStyle <- createStyle(
  fontSize = 10,
  fontColour = "#000000",
  fgFill = "#D9D9D9",      
  textDecoration = "bold",
  halign = "center",
  border = "Bottom"
)

# Data Formats
numStyle <- createStyle(numFmt = "#,##0")
pctStyle <- createStyle(numFmt = "0.0") 

# --- NEW: Traffic Light Styles ---
# Significantly HIGHER (Green)
posStyle <- createStyle(
  fontColour = "#006100",       # Dark Green Text (Readable)
  fgFill = "#C6EFCE",           # Light Green Background
  numFmt = "0.0" 
)

# Significantly LOWER (Red)
negStyle <- createStyle(
  fontColour = "#9C0006",       # Dark Red Text (Readable)
  fgFill = "#FFC7CE",           # Light Red Background
  numFmt = "0.0" 
)

# 2. SETUP WEIGHTS & STATS
# ------------------------------------------------------------------------------
weight_variable <- if("Weight_2" %in% names(refugees_final)) "Weight_2" else "weight"
if(!weight_variable %in% names(refugees_final)) {
  warning("Weight variable not found! Defaulting to unweighted counts.")
  refugees_final$temp_weight_col <- 1
  weight_variable <- "temp_weight_col"
}

# HELPER: Get Z-Score (Cluster vs Rest)
# Returns actual Z value. If sample too small, returns 0.
get_z_score <- function(n_cluster, pct_cluster, n_total, pct_total) {
  x_cluster <- n_cluster * (pct_cluster / 100)
  x_total   <- n_total * (pct_total / 100)
  n1 <- n_cluster; x1 <- x_cluster
  n2 <- n_total - n_cluster; x2 <- x_total - x_cluster
  
  if (n1 < 5 || n2 < 5 || n2 < 0) return(0)
  
  p_pool <- (x1 + x2) / (n1 + n2)
  # Check for 0 or 1 variance
  if(p_pool == 0 || p_pool == 1) return(0)
  
  se <- sqrt(p_pool * (1 - p_pool) * ((1 / n1) + (1 / n2)))
  
  if (se == 0) return(0)
  z <- ((x1 / n1) - (x2 / n2)) / se
  return(z)
}

# 3. DATA PROCESSING FUNCTION
# ------------------------------------------------------------------------------
process_cluster_data <- function(.data, var, cluster_var, wt_var) {
  
  # A. Basic Clean & Aggregation
  if (all(is.na(.data[[var]]))) return(NULL)
  
  data_clean <- .data %>%
    filter(!is.na(.data[[cluster_var]])) %>%
    rename(target = !!sym(var), cl = !!sym(cluster_var), wt = !!sym(wt_var)) %>%
    mutate(
      cl_val = as.character(cl),
      raw = as.character(target),
      lab = tryCatch(as.character(to_factor(target)), error = function(e) as.character(target)),
      Category = ifelse(raw == lab, lab, paste0(lab, " [", raw, "]"))
    ) %>%
    filter(!is.na(target))
  
  if(nrow(data_clean) == 0) return(NULL)
  
  # B. Calculate Stats
  # 1. Cluster Stats
  cl_stats <- data_clean %>%
    group_by(cl_val, Category) %>%
    summarise(n = sum(wt, na.rm = TRUE), .groups = "drop") %>% 
    group_by(cl_val) %>%
    mutate(pct = n / sum(n) * 100) %>% 
    ungroup()
  
  # 2. Total Stats
  tot_stats <- data_clean %>%
    group_by(Category) %>%
    summarise(n_Total = sum(wt, na.rm = TRUE), .groups = "drop") %>% 
    mutate(pct_Total = n_Total / sum(n_Total) * 100)
  
  # 3. Base Ns
  cl_Ns <- data_clean %>% group_by(cl_val) %>% summarise(N = sum(wt, na.rm = TRUE))
  grand_N <- sum(data_clean$wt, na.rm = TRUE)
  
  # C. Reshape (Silent Mode Logic)
  clusters <- sort(unique(data_clean$cl_val))
  
  # Counts
  wide_n <- cl_stats %>%
    select(Category, cl_val, n) %>%
    pivot_wider(names_from = cl_val, values_from = n, values_fill = 0, names_prefix = "N_")
  
  expected_n_cols <- paste0("N_", clusters)
  missing_n <- setdiff(expected_n_cols, names(wide_n))
  if(length(missing_n) > 0) wide_n[missing_n] <- 0
  wide_n <- wide_n %>% select(Category, all_of(expected_n_cols))
  
  df_n <- tot_stats %>% select(Category, n_Total) %>% left_join(wide_n, by = "Category")
  
  # Percentages
  wide_pct <- cl_stats %>%
    select(Category, cl_val, pct) %>%
    pivot_wider(names_from = cl_val, values_from = pct, values_fill = 0, names_prefix = "P_")
  
  expected_p_cols <- paste0("P_", clusters)
  missing_p <- setdiff(expected_p_cols, names(wide_pct))
  if(length(missing_p) > 0) wide_pct[missing_p] <- 0
  wide_pct <- wide_pct %>% select(Category, all_of(expected_p_cols))
  
  df_pct <- tot_stats %>% select(Category, pct_Total) %>% left_join(wide_pct, by = "Category")
  
  # D. Final DF
  df_gap <- data.frame(GAP = rep("", nrow(df_n)))
  final_df <- bind_cols(df_n, df_gap, df_pct %>% select(-Category))
  
  # E. Significance Mapping (Green/Red)
  sig_coords <- list()
  n_cl <- length(clusters)
  start_pct_col <- 5 + n_cl 
  
  for(i in 1:nrow(final_df)) {
    t_n <- grand_N
    t_pct <- df_pct$pct_Total[i]
    
    for(j in 1:n_cl) {
      cl_name <- clusters[j]
      
      # Get Cluster Data
      c_n <- cl_Ns$N[cl_Ns$cl_val == cl_name]
      col_name <- paste0("P_", cl_name)
      c_pct <- df_pct[[col_name]][i]
      
      # Calculate Z-Score
      z_score <- get_z_score(c_n, c_pct, t_n, t_pct)
      
      # Determine Color
      sig_type <- "none"
      if(z_score > 1.96) {
        sig_type <- "pos" # Higher (Green)
      } else if (z_score < -1.96) {
        sig_type <- "neg" # Lower (Red)
      }
      
      if(sig_type != "none") {
        target_col_idx <- start_pct_col + (j - 1)
        sig_coords[[length(sig_coords) + 1]] <- list(row = i, col = target_col_idx, type = sig_type)
      }
    }
  }
  
  return(list(
    data = final_df, 
    sigs = sig_coords, 
    clusters = clusters,
    n_clusters = n_cl
  ))
}

# 4. EXECUTION LOOP
# ------------------------------------------------------------------------------
wb <- createWorkbook()
addWorksheet(wb, "Cluster_Profiles")
showGridLines(wb, "Cluster_Profiles", showGridLines = FALSE) 

vars_to_skip <- c("ID", "Weight", "weight", "Weight_2", "cluster", "TEMP_ID_TRACKER", "temp_weight_col")
vars_to_analyze <- setdiff(names(refugees_final), vars_to_skip)

curr_row <- 1

print("Generating Traffic-Light Report...")

for (var in vars_to_analyze) {
  
  lbl <- var_label(refugees_final[[var]])
  if (is.null(lbl)) lbl <- "" 
  header_text <- paste0(var, " - ", lbl)
  
  tryCatch({
    res <- process_cluster_data(refugees_final, var, "cluster", weight_variable)
    
    if (!is.null(res)) {
      df <- res$data
      sigs <- res$sigs
      n_cl <- res$n_clusters
      clusters <- res$clusters
      
      # 1. Header
      writeData(wb, "Cluster_Profiles", header_text, startRow = curr_row, startCol = 1)
      addStyle(wb, "Cluster_Profiles", titleStyle, rows = curr_row, cols = 1:(n_cl*2 + 4), stack = TRUE)
      
      # 2. Super Headers
      sh_row <- curr_row + 1
      
      writeData(wb, "Cluster_Profiles", "COUNTS (N)", startRow = sh_row, startCol = 2)
      mergeCells(wb, "Cluster_Profiles", cols = 2:(2 + n_cl), rows = sh_row)
      addStyle(wb, "Cluster_Profiles", superHeaderStyle, rows = sh_row, cols = 2:(2 + n_cl), stack = TRUE)
      
      pct_start_col <- 4 + n_cl
      pct_end_col   <- pct_start_col + n_cl
      writeData(wb, "Cluster_Profiles", "PERCENTAGES (%)", startRow = sh_row, startCol = pct_start_col)
      mergeCells(wb, "Cluster_Profiles", cols = pct_start_col:pct_end_col, rows = sh_row)
      addStyle(wb, "Cluster_Profiles", superHeaderStyle, rows = sh_row, cols = pct_start_col:pct_end_col, stack = TRUE)
      
      # 3. Col Headers
      ch_row <- curr_row + 2
      headers <- c("Category", "Total", clusters, "", "Total", clusters)
      writeData(wb, "Cluster_Profiles", t(headers), startRow = ch_row, startCol = 1, colNames = FALSE)
      
      style_cols <- c(1:(2+n_cl), pct_start_col:pct_end_col) 
      addStyle(wb, "Cluster_Profiles", colHeaderStyle, rows = ch_row, cols = style_cols, stack = TRUE)
      
      # 4. Data
      d_row <- curr_row + 3
      writeData(wb, "Cluster_Profiles", df, startRow = d_row, startCol = 1, colNames = FALSE)
      
      # 5. Styling
      n_rows <- nrow(df)
      rows_range <- d_row:(d_row + n_rows - 1)
      
      # Base Numbers
      addStyle(wb, "Cluster_Profiles", numStyle, rows = rows_range, cols = 2:(2+n_cl), gridExpand = TRUE)
      # Base Percentages
      addStyle(wb, "Cluster_Profiles", pctStyle, rows = rows_range, cols = pct_start_col:pct_end_col, gridExpand = TRUE)
      
      # Highlights (Traffic Light)
      if(length(sigs) > 0) {
        for(s in sigs) {
          excel_r <- (d_row - 1) + s$row
          excel_c <- s$col
          
          # Choose Style based on Type
          chosen_style <- if(s$type == "pos") posStyle else negStyle
          
          addStyle(wb, "Cluster_Profiles", chosen_style, rows = excel_r, cols = excel_c, stack = TRUE)
        }
      }
      
      # Widths
      setColWidths(wb, "Cluster_Profiles", cols = 1, widths = 40)
      setColWidths(wb, "Cluster_Profiles", cols = c(2:(pct_end_col)), widths = 10)
      setColWidths(wb, "Cluster_Profiles", cols = 3 + n_cl, widths = 2) 
      
      curr_row <- curr_row + nrow(df) + 5
    }
  }, error = function(e) {
    message(paste("Error processing variable:", var))
  })
}

saveWorkbook(wb, "Cluster_Profiles_Wave5.xlsx", overwrite = TRUE)
print("SUCCESS")