# ==============================================================================
# 0. SETUP & LIBRARIES
# ==============================================================================
# install.packages(c("tidyverse", "naniar", "haven"))

library(tidyverse)
library(haven)
library(naniar)

# Fix for the 'select' conflict (MASS vs dplyr)
select <- dplyr::select 

# ==============================================================================
# 1. ROBUST DATA PREPARATION
# ==============================================================================

input_file <- "901849.sav"
refugees <- read_sav(input_file, encoding = "cp1251")

##### 1. Select ONLY universal or fixable variables
# Removed: E11, E8, J7, J5, Children variables
refugees_raw <- refugees %>% select(
  ID,
  Weight_2,
  # A. Social Unit & Push Factors (Asked to all)
  A2.1_1:A2.1_9,
  А3_01:А3_19,
  # V. Integration (Asked to all)
  V3, V5, V7_1:V7_3,
  # G. Help (G1 asked to all; G2 filtered)
  G1_01:G1_11, G2,
  # D. Housing (Asked to all)
  D1,
  # E. Work (E2, E3 asked to all)
  E2_01:E2_14, E3,
  # J. Intentions (J2, J4, J12 asked to all)
  J2,
  J4.0_01:J4.0_17, # Barriers
  J4_01:J4_22, # Incentives
  J12_01:J12_15, # Difficulties
  # K. Protection (K1 asked to all; K2, K3 filtered)
  K1_1:K1_4
)

##### 2. Advanced Cleaning & Logical Imputation
refugees_clean <- refugees_raw %>% mutate(
  ### --- A. SOC-DEM PROXIES ---
  # Who do you live with?
  living_situation = case_when(
    A2.1_1 == 1 ~ "Alone",
    A2.1_2 == 1 | A2.1_3 == 1 | A2.1_4 == 1 | A2.1_5 == 1 | A2.1_6 == 1 ~ "Family",
    A2.1_7 == 1 | A2.1_8 == 1 | A2.1_9 == 1 ~ "Friends/Other",
    TRUE ~ "Family"
  ),

  # Push Factors
  push_factor_destruction = ifelse(А3_04 == 1 | А3_05 == 1, 1, 0),
  push_factor_safety = ifelse(А3_01 == 1 | А3_02 == 1 | А3_03 == 1 | А3_08 == 1, 1, 0),
  push_factor_soft = ifelse(А3_06 == 1 | А3_09 == 1 | А3_10 == 1 | А3_14 == 1, 1, 0),

  ### --- V. INTEGRATION ---
  # Language: Recode NA as Low (safe assumption if skipped, or use 1)
  lang_knowledge = case_when(
    V3 <= 2 ~ 1, # Low
    V3 > 2 & V3 <= 4 ~ 2, # Mid
    V3 >= 5 ~ 3, # High
    is.na(V3) ~ 1
  ),

  # Attitude of locals
  local_attitude = case_when(
    V5 <= 2 ~ 1, # Positive
    V5 == 3 ~ 2, # Neutral
    V5 >= 4 ~ 3, # Negative
    is.na(V5) ~ 2 # Assume Neutral if missing
  ),
  has_local_friends = ifelse(V7_1 == 1 | V7_2 == 1, 1, 0),

  ### --- D. HOUSING ---
  housing_type = case_when(
    D1 == 1 ~ 1, # Independent
    D1 == 2 | D1 == 4 ~ 2, # Shared Stable
    D1 >= 3 ~ 3, # Instable/Camps/Hotels
    TRUE ~ 3
  ),

  ### --- E. WORK & ECON ---
  # Employment Status
  emp_status = case_when(
    E2_04 == 1 | E2_05 == 1 | E2_11 == 1 ~ 1, # Local Job
    E2_01 == 1 | E2_02 == 1 | E2_10 == 1 ~ 2, # Remote UA
    E2_06 == 1 ~ 3, # Searching
    E2_07 == 1 | E2_08 == 1 | E2_09 == 1 ~ 4, # Student
    E2_12 == 1 | E2_13 == 1 | E2_14 == 1 ~ 5, # Inactive
    TRUE ~ 3 # Assume searching if unclear
  ),

  # Economic Status
  econ_status = case_when(
    E3 <= 2 ~ 1, # Poor
    E3 == 3 | E3 == 4 ~ 2, # Mid
    E3 >= 5 ~ 3, # Good
    TRUE ~ 2
  ),

  ### --- G. HELP ---
  receives_housing_help = ifelse(G1_01 == 1, 1, 0), # Adjusted based on your dump (G1_12 missing in dump)
  receives_cash_help = ifelse(G1_04 == 1, 1, 0),

  # G2: Benefit Dynamic (HANDLING NA)
  # If G2 is NA, it implies they don't receive benefits.
  benefit_dynamic = case_when(
    G2 == 1 ~ 1, # Increased
    G2 == 2 ~ 2, # Decreased
    G2 == 3 ~ 3, # Same
    G2 == 4 | G2 == 5 ~ 4, # Stopped/Started
    is.na(G2) ~ 0 # NEW CATEGORY: Never Received / Not Applicable
  ),

  ### --- J. INTENTIONS ---
  plan_return = case_when(
    J2 == 1 | J2 == 2 ~ 1, # Yes
    J2 == 3 | J2 == 4 ~ 2, # No
    TRUE ~ 3 # Hard to say
  ),

  # Barriers (J4.0) & Difficulties (J12)
  # These are Multiple Choice, so NA usually means "Not selected" -> 0
  barrier_security = ifelse(J4.0_01 == 1 | J4.0_02 == 1 | J4.0_03 == 1, 1, 0),
  barrier_integration = ifelse(J4.0_05 == 1 | J4.0_10 == 1, 1, 0),
  barrier_quality_life = ifelse(J4.0_06 == 1 | J4.0_07 == 1 | J4.0_08 == 1, 1, 0),
  diff_material = ifelse(J12_02 == 1 | J12_03 == 1 | J12_04 == 1, 1, 0),
  diff_psych = ifelse(J12_08 == 1 | J12_09 == 1 | J12_10 == 1, 1, 0),
  diff_social = ifelse(J12_01 == 1 | J12_05 == 1 | J12_12 == 1, 1, 0),

  ### --- K. PROTECTION ---

  # K1: Legal Status
  status_secure = case_when(
    K1_1 == 1 ~ 1, # Has other docs (Secure)
    K1_2 == 1 | K1_3 == 1 ~ 2, # Temp / Planning (Insecure)
    TRUE ~ 3
  )
)

##### 3. Final Selection for Clustering
refugees_clustering_ready <- refugees_clean %>%
  select(
    living_situation,
    push_factor_destruction,
    push_factor_safety,
    push_factor_soft,
    lang_knowledge,
    local_attitude,
    has_local_friends,
    housing_type,
    emp_status,
    econ_status,
    receives_housing_help,
    receives_cash_help,
    benefit_dynamic,
    plan_return,
    barrier_security,
    barrier_integration,
    barrier_quality_life,
    diff_material,
    diff_psych,
    diff_social,
    status_secure
  ) %>%
  mutate_all(as.factor)

# Verify no NAs remain
print(sum(is.na(refugees_clustering_ready)))
# View(miss_var_summary(refugees_clustering_ready))