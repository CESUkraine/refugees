library(haven)
library(dplyr)
library(tidyr)
library(writexl)

# 1. Load data
input_file <- "901849.sav"
data <- read_sav(input_file, encoding = "cp1251")

# 2. Calculate the Profile (Weighted)
representative_profile <- data %>%
  select(E4.2_1:E4.2_7, Weight_2) %>% # Must include the weight column here
  mutate(across(E4.2_1:E4.2_7, ~ replace_na(as.numeric(.), 0))) %>%
  # Use weighted.mean instead of mean
  summarise(across(E4.2_1:E4.2_7, ~ weighted.mean(., w = Weight_2, na.rm = TRUE))) %>%
  pivot_longer(
    cols = everything(),
    names_to = "Source_ID",
    values_to = "Raw_Average"
  ) %>%
  mutate(
    Source_Name = sapply(Source_ID, function(x) attr(data[[x]], "label")),
    Representative_Share = (Raw_Average / sum(Raw_Average)) * 100
  )

# 3. View the result
print(representative_profile)
sum(representative_profile$Representative_Share)

write_xlsx(representative_profile, "representative_income_structure.xlsx")



refugees_clusters <- read_sav("refugees_wave5_with_clusters.sav")

# 2. Calculate Profile per Cluster
cluster_profiles <- refugees_clusters %>%
    # Select income columns and the cluster variable
    select(cluster, E4.2_1:E4.2_7) %>%
    # Replace NA with 0 (as confirmed)
    mutate(across(E4.2_1:E4.2_7, ~ replace_na(as.numeric(.), 0))) %>%
    # Group by cluster
    group_by(cluster) %>%
    # Calculate mean for each source within each cluster
    summarise(across(everything(), mean), .groups = "drop") %>%
    # Pivot to long format for normalization and plotting
    pivot_longer(
        cols = E4.2_1:E4.2_7,
        names_to = "Source_ID",
        values_to = "Raw_Average"
    ) %>%
    # Normalize within each cluster so each cluster sums to 100%
    group_by(cluster) %>%
    mutate(
        Source_Label = sapply(Source_ID, function(x) attr(refugees_clusters[[x]], "label")),
        Representative_Share = (Raw_Average / sum(Raw_Average)) * 100
    ) %>%
    ungroup()

# 3. View the profiles
print(cluster_profiles)

write_xlsx(cluster_profiles, "spending_sources_clusters.xlsx")
