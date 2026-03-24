# ==============================================================================
# STABLE VARIABLE SELECTION FOR UNSUPERVISED CLUSTERING (K-MODES + BORUTA STABILITY)
# Logic:
# 1) run 2–3 (or more) baseline clusterings (different k / seeds / optional feature subsets)
# 2) run Boruta with cluster-as-label for each run
# 3) keep variables that appear as Confirmed often (stability frequency)
# 4) re-cluster on stable vars + export SPSS + Excel profile(s)
# ==============================================================================

# ==============================================================================
# 0. LIBRARIES
# ==============================================================================
suppressPackageStartupMessages({
  library(tidyverse)
  library(haven)
  library(naniar)
  library(factoextra)
  library(klaR)
  library(openxlsx)
  library(labelled)
  library(Boruta)
})

select <- dplyr::select

# ==============================================================================
# 1. READ DATA
# ==============================================================================
input_file <- "C:/Users/vever/Downloads/901849.sav"
stopifnot(file.exists(input_file))
refugees <- read_sav(input_file, encoding = "CP1251")

# ==============================================================================
# 2. CLEAN NAMES + NORMALIZE ID (CRITICAL)
# ==============================================================================
fix_names <- function(df) {
  n <- names(df)
  repl <- c("А"="A","В"="B","С"="C","Е"="E","Н"="H","К"="K","М"="M","О"="O","Р"="P","Т"="T","Х"="X")
  n <- stringr::str_replace_all(n, repl)
  n <- stringr::str_replace_all(n, "\\s+", "")
  names(df) <- n
  df
}
refugees <- fix_names(refugees)

if ("ID" %in% names(refugees)) {
  refugees <- refugees %>% mutate(ID = as.character(ID) %>% stringr::str_trim())
} else {
  refugees <- refugees %>% mutate(ID = as.character(seq_len(nrow(refugees))))
}

# ==============================================================================
# 3. SAFE SELECT
# ==============================================================================
refugees_raw <- refugees %>%
  select(
    any_of(c("ID", "Weight_2")),
    matches("^A2\\.1_\\d+$"),
    matches("^A3_\\d+$"),
    any_of(c("V3","V5")),
    matches("^V7_\\d+$"),
    matches("^G1_\\d+$"),
    any_of("G2"),
    any_of("D1"),
    matches("^E2_\\d+$"),
    any_of("E3"),
    any_of("J2"),
    matches("^J4\\.0_\\d+$"),
    matches("^J4_\\d+$"),
    matches("^J12_\\d+$"),
    matches("^K1_\\d+$")
  )

getv <- function(df, nm) if (nm %in% names(df)) df[[nm]] else rep(NA, nrow(df))
is1  <- function(x) ifelse(is.na(x), 0L, ifelse(x == 1, 1L, 0L))

# ==============================================================================
# 4. FEATURE ENGINEERING
# ==============================================================================
refugees_clean <- refugees_raw %>%
  mutate(
    living_situation = case_when(
      getv(., "A2.1_1") == 1 ~ "Alone",
      getv(., "A2.1_2") == 1 | getv(., "A2.1_3") == 1 | getv(., "A2.1_4") == 1 |
        getv(., "A2.1_5") == 1 | getv(., "A2.1_6") == 1 ~ "Family",
      getv(., "A2.1_7") == 1 | getv(., "A2.1_8") == 1 | getv(., "A2.1_9") == 1 ~ "Friends/Other",
      TRUE ~ "Family"
    ),
    
    push_factor_destruction = ifelse(is1(getv(., "A3_04")) == 1 | is1(getv(., "A3_05")) == 1, 1L, 0L),
    push_factor_safety      = ifelse(is1(getv(., "A3_01")) == 1 | is1(getv(., "A3_02")) == 1 |
                                       is1(getv(., "A3_03")) == 1 | is1(getv(., "A3_08")) == 1, 1L, 0L),
    push_factor_soft        = ifelse(is1(getv(., "A3_06")) == 1 | is1(getv(., "A3_09")) == 1 |
                                       is1(getv(., "A3_10")) == 1 | is1(getv(., "A3_14")) == 1, 1L, 0L),
    
    lang_knowledge = case_when(
      getv(., "V3") <= 2 ~ 1L,
      getv(., "V3") > 2 & getv(., "V3") <= 4 ~ 2L,
      getv(., "V3") >= 5 ~ 3L,
      is.na(getv(., "V3")) ~ 1L
    ),
    
    local_attitude = case_when(
      getv(., "V5") <= 2 ~ 1L,
      getv(., "V5") == 3 ~ 2L,
      getv(., "V5") >= 4 ~ 3L,
      is.na(getv(., "V5")) ~ 2L
    ),
    
    has_local_friends = ifelse(is1(getv(., "V7_1")) == 1 | is1(getv(., "V7_2")) == 1, 1L, 0L),
    
    housing_type = case_when(
      getv(., "D1") == 1 ~ 1L,
      getv(., "D1") == 2 | getv(., "D1") == 4 ~ 2L,
      getv(., "D1") >= 3 ~ 3L,
      TRUE ~ 3L
    ),
    
    emp_status = case_when(
      is1(getv(., "E2_04")) == 1 | is1(getv(., "E2_05")) == 1 | is1(getv(., "E2_11")) == 1 ~ 1L,
      is1(getv(., "E2_01")) == 1 | is1(getv(., "E2_02")) == 1 | is1(getv(., "E2_10")) == 1 ~ 2L,
      is1(getv(., "E2_06")) == 1 ~ 3L,
      is1(getv(., "E2_07")) == 1 | is1(getv(., "E2_08")) == 1 | is1(getv(., "E2_09")) == 1 ~ 4L,
      is1(getv(., "E2_12")) == 1 | is1(getv(., "E2_13")) == 1 | is1(getv(., "E2_14")) == 1 ~ 5L,
      TRUE ~ 3L
    ),
    
    econ_status = case_when(
      getv(., "E3") <= 2 ~ 1L,
      getv(., "E3") %in% c(3, 4) ~ 2L,
      getv(., "E3") >= 5 ~ 3L,
      TRUE ~ 2L
    ),
    
    receives_housing_help = ifelse(is1(getv(., "G1_01")) == 1, 1L, 0L),
    receives_cash_help    = ifelse(is1(getv(., "G1_04")) == 1, 1L, 0L),
    
    benefit_dynamic = case_when(
      getv(., "G2") == 1 ~ 1L,
      getv(., "G2") == 2 ~ 2L,
      getv(., "G2") == 3 ~ 3L,
      getv(., "G2") %in% c(4, 5) ~ 4L,
      is.na(getv(., "G2")) ~ 0L
    ),
    
    plan_return = case_when(
      getv(., "J2") %in% c(1, 2) ~ 1L,
      getv(., "J2") %in% c(3, 4) ~ 2L,
      TRUE ~ 3L
    ),
    
    barrier_security     = ifelse(is1(getv(., "J4.0_01")) == 1 | is1(getv(., "J4.0_02")) == 1 | is1(getv(., "J4.0_03")) == 1, 1L, 0L),
    barrier_integration  = ifelse(is1(getv(., "J4.0_05")) == 1 | is1(getv(., "J4.0_10")) == 1, 1L, 0L),
    barrier_quality_life = ifelse(is1(getv(., "J4.0_06")) == 1 | is1(getv(., "J4.0_07")) == 1 | is1(getv(., "J4.0_08")) == 1, 1L, 0L),
    
    diff_material = ifelse(is1(getv(., "J12_02")) == 1 | is1(getv(., "J12_03")) == 1 | is1(getv(., "J12_04")) == 1, 1L, 0L),
    diff_psych    = ifelse(is1(getv(., "J12_08")) == 1 | is1(getv(., "J12_09")) == 1 | is1(getv(., "J12_10")) == 1, 1L, 0L),
    diff_social   = ifelse(is1(getv(., "J12_01")) == 1 | is1(getv(., "J12_05")) == 1 | is1(getv(., "J12_12")) == 1, 1L, 0L),
    
    status_secure = case_when(
      is1(getv(., "K1_1")) == 1 ~ 1L,
      is1(getv(., "K1_2")) == 1 | is1(getv(., "K1_3")) == 1 ~ 2L,
      TRUE ~ 3L
    )
  )

# ==============================================================================
# 5. FINAL DATA FOR CLUSTERING (ALL FEATURES)
# ==============================================================================
all_features <- c(
  "living_situation",
  "push_factor_destruction","push_factor_safety","push_factor_soft",
  "lang_knowledge","local_attitude","has_local_friends",
  "housing_type","emp_status","econ_status",
  "receives_housing_help","receives_cash_help","benefit_dynamic",
  "plan_return",
  "barrier_security","barrier_integration","barrier_quality_life",
  "diff_material","diff_psych","diff_social",
  "status_secure"
)

refugees_clustering_ready <- refugees_clean %>%
  select(all_of(all_features)) %>%
  mutate(across(everything(), as.factor))

message("NAs in clustering features: ", sum(is.na(refugees_clustering_ready)))

# ==============================================================================
# 6. PREPARE DATA FOR K-MODES (KEEP IDS ALIGNED)
# ==============================================================================
id_vec <- refugees_raw$ID %>% as.character() %>% stringr::str_trim()
stopifnot(length(id_vec) == nrow(refugees_clustering_ready))

temp_combined <- refugees_clustering_ready %>% mutate(TEMP_ID_TRACKER = id_vec)

df_clean_combined <- temp_combined %>%
  drop_na() %>%
  mutate(across(-TEMP_ID_TRACKER, as.factor))

df_ids  <- df_clean_combined %>% transmute(ID = as.character(TEMP_ID_TRACKER) %>% stringr::str_trim())
df_algo <- df_clean_combined %>% select(-TEMP_ID_TRACKER) %>% as.data.frame()

message("Data ready. Original rows: ", nrow(refugees),
        " | Final rows used for clustering: ", nrow(df_algo))

# ==============================================================================
# 7. DEFINE EXPERIMENT GRID (2–3+ VARIANTS)
#    You can add more rows here. This is the whole "stability" idea.
# ==============================================================================
# Optionally: try different feature subsets (all vs no-push-factors vs no-diffs, etc.)
feature_sets <- list(
  full = names(df_algo),
  no_push = setdiff(names(df_algo), c("push_factor_destruction","push_factor_safety","push_factor_soft")),
  no_diff = setdiff(names(df_algo), c("diff_material","diff_psych","diff_social"))
)

# Choose 2–3 ks and some seeds (you can expand to 10+ runs if you want)
experiment_grid <- tidyr::crossing(
  k = c(3, 4, 5),
  seed = c(101, 202, 303),
  feature_set = names(feature_sets)
) %>%
  mutate(run_id = paste0("k", k, "s", seed, "", feature_set))

print(experiment_grid)

# ==============================================================================
# 8. FUNCTIONS: K-MODES RUN + BORUTA (cluster-as-label)
# ==============================================================================
run_kmodes <- function(df_algo, k, seed) {
  set.seed(seed)
  km <- klaR::kmodes(df_algo, modes = k, iter.max = 25, weighted = FALSE)
  km
}

run_boruta_cluster_label <- function(df_x, cluster_vec, seed, maxRuns = 200) {
  set.seed(seed)
  b <- Boruta::Boruta(x = as.data.frame(df_x), y = as.factor(cluster_vec), doTrace = 0, maxRuns = maxRuns)
  b <- Boruta::TentativeRoughFix(b)
  confirmed <- Boruta::getSelectedAttributes(b, withTentative = FALSE)
  tentative <- Boruta::getSelectedAttributes(b, withTentative = TRUE)
  list(confirmed = confirmed, tentative = tentative, boruta_obj = b)
}

# ==============================================================================
# 9. RUN EXPERIMENTS + COLLECT SELECTED VARIABLES
# ==============================================================================
results_list <- vector("list", nrow(experiment_grid))

for (i in seq_len(nrow(experiment_grid))) {
  k_i <- experiment_grid$k[i]
  s_i <- experiment_grid$seed[i]
  fs  <- experiment_grid$feature_set[i]
  rid <- experiment_grid$run_id[i]
  
  message("\n[RUN] ", rid)
  
  cols <- feature_sets[[fs]]
  df_x <- df_algo[, cols, drop = FALSE]
  
  km <- run_kmodes(df_x, k = k_i, seed = s_i)
  
  # Boruta seed: change a bit so it isn't identical to kmodes seed
  b <- run_boruta_cluster_label(df_x, km$cluster, seed = s_i + 999, maxRuns = 200)
  
  results_list[[i]] <- list(
    run_id = rid,
    k = k_i,
    seed = s_i,
    feature_set = fs,
    cluster = km$cluster,
    confirmed = b$confirmed,
    tentative = b$tentative
  )
}

# ==============================================================================
# 10. STABILITY TABLE (VARIANT C): "STABLE IN EACH FEATURE_SET"
#   Idea: variable must be Confirmed with share >= threshold WITHIN EVERY feature_set
#   This is stricter than overall frequency and usually yields fewer vars.
# ==============================================================================

stopifnot(exists("results_list"), exists("experiment_grid"))

suppressPackageStartupMessages({
  library(dplyr)
  library(tidyr)
  library(purrr)
  library(tibble)
})

# --- 10.0 Long table: run_id x variable for CONFIRMED ---
confirmed_long <- purrr::map_dfr(results_list, function(x) {
  tibble(
    run_id   = x$run_id,
    variable = x$confirmed
  )
})

# Optional: also build tentative (for diagnostics only)
tentative_long <- purrr::map_dfr(results_list, function(x) {
  tibble(
    run_id   = x$run_id,
    variable = x$tentative
  )
})

# --- 10.1 Attach run metadata (k, seed, feature_set) ---
run_meta <- experiment_grid %>%
  select(run_id, k, seed, feature_set) %>%
  distinct()

confirmed_long <- confirmed_long %>% left_join(run_meta, by = "run_id")
tentative_long <- tentative_long %>% left_join(run_meta, by = "run_id")

# --- 10.2 Per-feature_set shares for each variable ---
# count "in how many runs within feature_set var was confirmed"
fs_share <- confirmed_long %>%
  distinct(run_id, feature_set, variable) %>%
  count(feature_set, variable, name = "confirmed_n_fs") %>%
  left_join(run_meta %>% count(feature_set, name = "runs_fs"), by = "feature_set") %>%
  mutate(confirmed_share_fs = confirmed_n_fs / runs_fs)

# (Diagnostics) overall frequency table, just to print top ones
freq_df <- confirmed_long %>%
  distinct(run_id, variable) %>%
  count(variable, name = "confirmed_n") %>%
  mutate(confirmed_share = confirmed_n / nrow(run_meta)) %>%
  arrange(desc(confirmed_share), desc(confirmed_n), variable)

message("\n[INFO] Top variables by overall Confirmed frequency (diagnostic):")
print(freq_df, n = min(30, nrow(freq_df)))

# --- 10.3 FINAL selection rule: must meet threshold in EACH feature_set ---
stability_threshold <- 0.70  # <- зроби 0.8 якщо хочеш ще жорсткіше

stable_vars <- fs_share %>%
  group_by(variable) %>%
  summarise(
    min_share_across_fs = min(confirmed_share_fs),
    .groups = "drop"
  ) %>%
  filter(min_share_across_fs >= stability_threshold) %>%
  arrange(desc(min_share_across_fs), variable) %>%
  pull(variable)

message("\n[SELECTED] Stable variables (must be stable in EACH feature_set):")
print(stable_vars)

# --- 10.4 If too few: OPTIONAL controlled fallback (tight, not bloating) ---
# Якщо не хочеш fallback — постав min_vars_needed <- 0
min_vars_needed <- 8
if (length(stable_vars) < min_vars_needed) {
  message("[WARN] Too few vars under strict rule (n=", length(stable_vars), "). ",
          "Fallback: take top ", min_vars_needed, " by MIN share across feature_set.")
  
  stable_vars <- fs_share %>%
    group_by(variable) %>%
    summarise(min_share_across_fs = min(confirmed_share_fs), .groups = "drop") %>%
    arrange(desc(min_share_across_fs), variable) %>%
    slice_head(n = min(min_vars_needed, n())) %>%
    pull(variable)
}

message("\n[FINAL] stable_vars:")
print(stable_vars)

# (Optional) save fs_share diagnostic table
fs_diag <- fs_share %>%
  arrange(feature_set, desc(confirmed_share_fs), variable)

# write.csv(fs_diag,
#           file = file.path(getwd(), paste0("boruta_stability_by_featureset_", format(Sys.Date(), "%Y%m%d"), ".csv")),
#           row.names = FALSE)

# ==============================================================================
# 11. FINAL CLUSTERING ON STABLE VARS (choose ONE final k and seed)
# ==============================================================================
final_k <- 4
final_seed <- 444

df_algo_stable <- df_algo %>% select(all_of(stable_vars))

set.seed(final_seed)
km.final <- kmodes(df_algo_stable, modes = final_k, iter.max = 30, weighted = FALSE)

df_clustered_ids <- df_ids %>%
  mutate(
    ID = as.character(ID) %>% stringr::str_trim(),
    cluster_stable = as.factor(km.final$cluster)
  )

message("\n[INFO] Final cluster sizes (stable vars):")
print(table(df_clustered_ids$cluster_stable))

# Merge back
refugees_final <- refugees %>%
  mutate(ID = as.character(ID) %>% stringr::str_trim()) %>%
  left_join(df_clustered_ids, by = "ID")

# Export SPSS
out_sav_stable <- "refugees_with_clusters_stable.sav"
write_sav(refugees_final, out_sav_stable)
message("💾 Saved SPSS file: ", out_sav_stable)


cat("\n--- DIAG ---\n")
cat("stable_vars length:", length(stable_vars), "\n")
cat("stable_vars in refugees_final:", sum(stable_vars %in% names(refugees_final)), "\n")
print(head(stable_vars, 20))

cat("\ncluster_stable NA share:\n")
print(mean(is.na(refugees_final$cluster_stable)))

cat("\nWeight_2 class:\n")
if ("Weight_2" %in% names(refugees_final)) {
  print(class(refugees_final$Weight_2))
  print(head(refugees_final$Weight_2, 10))
  cat("Weight_2 numeric NA share after as.numeric:\n")
  print(mean(is.na(suppressWarnings(as.numeric(refugees_final$Weight_2)))))
} else {
  cat("Weight_2 NOT FOUND\n")
}
cat("--- END DIAG ---\n")
# ==============================================================================
# 12. EXCEL REPORT: PROFILE ALL VARIABLES (refugees_final) BY cluster_stable
# ==============================================================================

suppressPackageStartupMessages({
  library(dplyr)
  library(tidyr)
  library(openxlsx)
  library(labelled)
  library(stringr)
})

# ------------------------------------------------------------------------------
# 12.0 Preconditions + profiling source
# ------------------------------------------------------------------------------
stopifnot(exists("refugees_final"))
stopifnot("cluster_stable" %in% names(refugees_final))

profile_df  <- refugees_final
cluster_var <- "cluster_stable"

message("[INFO] Profiling source: refugees_final (RAW dataset + cluster_stable)")
message("[INFO] Rows: ", nrow(profile_df), " | Cols: ", ncol(profile_df))

# ------------------------------------------------------------------------------
# 12.1 Styles
# ------------------------------------------------------------------------------
titleStyle <- createStyle(fontSize = 12, textDecoration = "bold", fgFill = "#E7E6E6")
superHeaderStyle <- createStyle(fontSize = 11, fontColour = "#FFFFFF", fgFill = "#44546A",
                                textDecoration = "bold", halign = "center",
                                border = "TopBottomLeftRight")
colHeaderStyle <- createStyle(fontSize = 10, fgFill = "#D9D9D9", textDecoration = "bold",
                              halign = "center", border = "Bottom")
numStyle <- createStyle(numFmt = "#,##0")
pctStyle <- createStyle(numFmt = "0.0")
posStyle <- createStyle(fontColour = "#006100", fgFill = "#C6EFCE", numFmt = "0.0")
negStyle <- createStyle(fontColour = "#9C0006", fgFill = "#FFC7CE", numFmt = "0.0")

# ------------------------------------------------------------------------------
# 12.2 Weight handling (robust, NO $__wt)
# ------------------------------------------------------------------------------
# Create guaranteed numeric weight column: wt_num
if ("Weight_2" %in% names(profile_df)) {
  
  w <- profile_df[["Weight_2"]]
  
  # haven_labelled часто сидить як "labelled"; unclass -> underlying numeric codes
  if (inherits(w, "haven_labelled") || inherits(w, "labelled")) {
    wt_num <- suppressWarnings(as.numeric(unclass(w)))
    
  } else if (is.factor(w)) {
    wt_num <- suppressWarnings(as.numeric(as.character(w)))
    
  } else if (is.character(w)) {
    # прибрати кому як десятковий роздільник, зайві пробіли
    ww <- gsub(",", ".", trimws(w))
    wt_num <- suppressWarnings(as.numeric(ww))
    
  } else {
    wt_num <- suppressWarnings(as.numeric(w))
  }
  
  if (all(is.na(wt_num))) {
    warning("[WARN] Weight_2 exists but cannot be converted to numeric -> using unweighted counts.")
    wt_num <- rep(1, nrow(profile_df))
  }
  
} else {
  warning("[WARN] Weight_2 not found -> using unweighted counts.")
  wt_num <- rep(1, nrow(profile_df))
}

profile_df[["wt_num"]] <- wt_num
wt_var <- "wt_num"

# ------------------------------------------------------------------------------
# 12.3 Z-test helper (difference in proportions: cluster vs rest)
# ------------------------------------------------------------------------------
get_z_score <- function(n_cluster, pct_cluster, n_total, pct_total) {
  x1 <- n_cluster * (pct_cluster / 100)
  xT <- n_total   * (pct_total / 100)
  
  n1 <- n_cluster
  n2 <- n_total - n_cluster
  x2 <- xT - x1
  
  if (is.na(n1) || is.na(n2) || n1 < 5 || n2 < 5 || n2 < 0) return(0)
  
  p_pool <- (x1 + x2) / (n1 + n2)
  if (p_pool <= 0 || p_pool >= 1) return(0)
  
  se <- sqrt(p_pool * (1 - p_pool) * (1 / n1 + 1 / n2))
  if (se == 0) return(0)
  
  ((x1 / n1) - (x2 / n2)) / se
}

# ------------------------------------------------------------------------------
# 12.4 Processing helper
# ------------------------------------------------------------------------------
process_cluster_data <- function(.data, var, cluster_var, wt_var) {
  
  if (!var %in% names(.data)) return(NULL)
  if (!cluster_var %in% names(.data)) return(NULL)
  if (!wt_var %in% names(.data)) return(NULL)
  if (all(is.na(.data[[var]]))) return(NULL)
  
  data_clean <- .data %>%
    filter(!is.na(.data[[cluster_var]])) %>%
    transmute(
      cl = as.character(.data[[cluster_var]]),
      wt = suppressWarnings(as.numeric(.data[[wt_var]])),
      target = .data[[var]],
      raw = as.character(target),
      lab = tryCatch(as.character(to_factor(target)), error = function(e) as.character(target)),
      Category = ifelse(raw == lab, lab, paste0(lab, " [", raw, "]"))
    ) %>%
    filter(!is.na(target)) %>%
    filter(!is.na(wt))
  
  if (nrow(data_clean) == 0) return(NULL)
  
  cl_stats <- data_clean %>%
    group_by(cl, Category) %>%
    summarise(n = sum(wt, na.rm = TRUE), .groups = "drop") %>%
    group_by(cl) %>%
    mutate(pct = n / sum(n) * 100) %>%
    ungroup()
  
  tot_stats <- data_clean %>%
    group_by(Category) %>%
    summarise(n_Total = sum(wt, na.rm = TRUE), .groups = "drop") %>%
    mutate(pct_Total = n_Total / sum(n_Total) * 100)
  
  cl_Ns   <- data_clean %>% group_by(cl) %>% summarise(N = sum(wt, na.rm = TRUE), .groups = "drop")
  grand_N <- sum(data_clean$wt, na.rm = TRUE)
  
  clusters <- sort(unique(data_clean$cl))
  n_cl <- length(clusters)
  
  wide_n <- cl_stats %>%
    select(Category, cl, n) %>%
    pivot_wider(names_from = cl, values_from = n, values_fill = 0, names_prefix = "N_")
  expected_n_cols <- paste0("N_", clusters)
  for (mc in setdiff(expected_n_cols, names(wide_n))) wide_n[[mc]] <- 0
  wide_n <- wide_n %>% select(Category, all_of(expected_n_cols))
  df_n <- tot_stats %>% select(Category, n_Total) %>% left_join(wide_n, by = "Category")
  
  wide_pct <- cl_stats %>%
    select(Category, cl, pct) %>%
    pivot_wider(names_from = cl, values_from = pct, values_fill = 0, names_prefix = "P_")
  expected_p_cols <- paste0("P_", clusters)
  for (mc in setdiff(expected_p_cols, names(wide_pct))) wide_pct[[mc]] <- 0
  wide_pct <- wide_pct %>% select(Category, all_of(expected_p_cols))
  df_pct <- tot_stats %>% select(Category, pct_Total) %>% left_join(wide_pct, by = "Category")
  
  final_df <- bind_cols(
    df_n,
    data.frame(GAP = rep("", nrow(df_n))),
    df_pct %>% select(-Category)
  )
  
  # significance mapping
  sig_coords <- list()
  pct_start_col <- 4 + n_cl
  start_cluster_pct_col <- pct_start_col + 1
  
  for (i in seq_len(nrow(final_df))) {
    t_pct <- df_pct$pct_Total[i]
    for (j in seq_len(n_cl)) {
      cl_name <- clusters[j]
      c_n   <- cl_Ns$N[cl_Ns$cl == cl_name]
      c_pct <- df_pct[[paste0("P_", cl_name)]][i]
      z <- get_z_score(c_n, c_pct, grand_N, t_pct)
      
      sig_type <- if (z > 1.96) "pos" else if (z < -1.96) "neg" else "none"
      if (sig_type != "none") {
        target_col_idx <- start_cluster_pct_col + (j - 1)
        sig_coords[[length(sig_coords) + 1]] <- list(row = i, col = target_col_idx, type = sig_type)
      }
    }
  }
  
  list(data = final_df, sigs = sig_coords, clusters = clusters, n_clusters = n_cl)
}

# ------------------------------------------------------------------------------
# 12.5 Workbook generation (ALL variables in refugees_final)
# ------------------------------------------------------------------------------
wb <- createWorkbook()
addWorksheet(wb, "Cluster_Profiles_ALL")
showGridLines(wb, "Cluster_Profiles_ALL", showGridLines = FALSE)

# Exclude only technical columns
tech_exclude <- unique(c(
  "ID",
  cluster_var,
  "Weight_2",
  wt_var
))

vars_to_analyze <- setdiff(names(profile_df), tech_exclude)

# Drop columns that are fully NA
vars_to_analyze <- vars_to_analyze[
  !vapply(profile_df[vars_to_analyze], function(x) all(is.na(x)), logical(1))
]

message("[INFO] Vars to analyze: ", length(vars_to_analyze))

curr_row <- 1
skipped_null <- 0
written_blocks <- 0

for (var in vars_to_analyze) {
  
  lbl <- tryCatch(var_label(profile_df[[var]]), error = function(e) NULL)
  if (is.null(lbl)) lbl <- ""
  header_text <- paste0(var, ifelse(nchar(lbl) > 0, paste0(" - ", lbl), ""))
  
  res <- tryCatch(
    process_cluster_data(profile_df, var, cluster_var, wt_var),
    error = function(e) NULL
  )
  
  if (is.null(res)) {
    skipped_null <- skipped_null + 1
    next
  }
  
  written_blocks <- written_blocks + 1
  
  df <- res$data
  sigs <- res$sigs
  clusters <- res$clusters
  n_cl <- res$n_clusters
  
  # Title
  writeData(wb, "Cluster_Profiles_ALL", header_text, startRow = curr_row, startCol = 1)
  addStyle(wb, "Cluster_Profiles_ALL", titleStyle,
           rows = curr_row, cols = 1:(n_cl * 2 + 4), stack = TRUE)
  
  # Super headers
  sh_row <- curr_row + 1
  writeData(wb, "Cluster_Profiles_ALL", "COUNTS", startRow = sh_row, startCol = 2)
  mergeCells(wb, "Cluster_Profiles_ALL", cols = 2:(2 + n_cl), rows = sh_row)
  addStyle(wb, "Cluster_Profiles_ALL", superHeaderStyle,
           rows = sh_row, cols = 2:(2 + n_cl), stack = TRUE)
  
  pct_start_col <- 4 + n_cl
  pct_end_col   <- pct_start_col + n_cl
  writeData(wb, "Cluster_Profiles_ALL", "PERCENTAGES (%)", startRow = sh_row, startCol = pct_start_col)
  mergeCells(wb, "Cluster_Profiles_ALL", cols = pct_start_col:pct_end_col, rows = sh_row)
  addStyle(wb, "Cluster_Profiles_ALL", superHeaderStyle,
           rows = sh_row, cols = pct_start_col:pct_end_col, stack = TRUE)
  
  # Column headers
  ch_row <- curr_row + 2
  headers <- c("Category", "Total", clusters, "", "Total", clusters)
  writeData(wb, "Cluster_Profiles_ALL", t(headers), startRow = ch_row, startCol = 1, colNames = FALSE)
  
  style_cols <- c(1:(2 + n_cl), pct_start_col:pct_end_col)
  addStyle(wb, "Cluster_Profiles_ALL", colHeaderStyle,
           rows = ch_row, cols = style_cols, stack = TRUE)
  
  # Data table
  data_start <- curr_row + 3
  writeData(wb, "Cluster_Profiles_ALL", df, startRow = data_start, startCol = 1, colNames = FALSE)
  
  # Formats
  count_cols <- 2:(2 + n_cl)
  pct_cols   <- pct_start_col:pct_end_col
  addStyle(wb, "Cluster_Profiles_ALL", numStyle,
           rows = data_start:(data_start + nrow(df) - 1),
           cols = count_cols, gridExpand = TRUE, stack = TRUE)
  addStyle(wb, "Cluster_Profiles_ALL", pctStyle,
           rows = data_start:(data_start + nrow(df) - 1),
           cols = pct_cols, gridExpand = TRUE, stack = TRUE)
  
  # Sig styling
  if (length(sigs) > 0) {
    for (s in sigs) {
      rr <- data_start + s$row - 1
      cc <- s$col
      addStyle(wb, "Cluster_Profiles_ALL",
               if (s$type == "pos") posStyle else negStyle,
               rows = rr, cols = cc, stack = TRUE)
    }
  }
  
  # Borders
  addStyle(wb, "Cluster_Profiles_ALL",
           createStyle(border = "TopBottomLeftRight"),
           rows = ch_row:(data_start + nrow(df) - 1),
           cols = 1:(n_cl * 2 + 4),
           gridExpand = TRUE, stack = TRUE)
  
  # spacing
  curr_row <- data_start + nrow(df) + 2
}

# Column widths
n_clusters_actual <- length(unique(as.character(profile_df[[cluster_var]])))
setColWidths(wb, "Cluster_Profiles_ALL", cols = 1, widths = 42)
setColWidths(wb, "Cluster_Profiles_ALL", cols = 2:(2 + n_clusters_actual), widths = 12)
setColWidths(wb, "Cluster_Profiles_ALL",
             cols = (4 + n_clusters_actual):(4 + 2*n_clusters_actual),
             widths = 14)

# Save
out_file <- file.path(getwd(), paste0("cluster_profiles_ALLVARS_", format(Sys.Date(), "%Y%m%d"), ".xlsx"))
saveWorkbook(wb, out_file, overwrite = TRUE)

message("\n[OK] Excel saved: ", out_file)
message("[DEBUG] written_blocks=", written_blocks,
        " | skipped_null=", skipped_null,
        " | analyzed_vars=", length(vars_to_analyze))





# ==============================================================================
# 13. EXPORT: LISTS OF VARIABLES USED IN CLUSTERING
#   - used_in_algo: реально використані у kmodes (df_algo після drop_na)
#   - stable_vars: фінальні стабільні змінні (використані у km.final)
# ==============================================================================

stopifnot(exists("df_algo"))
stopifnot(exists("stable_vars"))

used_in_algo <- names(df_algo)
used_in_final <- stable_vars

cat("\n--- VARIABLE LISTS ---\n")
cat("df_algo vars (actually used in clustering input): ", length(used_in_algo), "\n")
print(used_in_algo)

cat("\nstable_vars (final clustering vars): ", length(used_in_final), "\n")
print(used_in_final)

# Optional: save to disk (CSV + TXT)
var_list_df <- tibble(
  variable = c(used_in_algo, used_in_final),
  role = c(rep("df_algo_input", length(used_in_algo)),
           rep("stable_final", length(used_in_final)))
) %>%
  distinct() %>%
  arrange(role, variable)

write.csv(var_list_df,
          file = file.path(getwd(), paste0("clustering_variable_lists_", format(Sys.Date(), "%Y%m%d"), ".csv")),
          row.names = FALSE)

writeLines(used_in_algo,
           con = file.path(getwd(), paste0("vars_df_algo_input_", format(Sys.Date(), "%Y%m%d"), ".txt")))
writeLines(used_in_final,
           con = file.path(getwd(), paste0("vars_stable_final_", format(Sys.Date(), "%Y%m%d"), ".txt")))

message("[OK] Saved variable lists: CSV + two TXT files in working directory.")

