library(haven)
library(dplyr)
library(tidyr)

input_file <- "901849.sav"
data <- read_sav(input_file, encoding = "cp1251")

# Keep only what is needed
data_selected <- data %>%
  select(
    Weight_2,      # respondent weight
    A2.1_6,        # partnership status: used for 2-parent adjustment
    B1,            # number of children
    starts_with("B2_"),      # child ages
    matches("^B5_\\d+_2$"),  # attends Ukrainian school
    S1,            # respondent gender
    S2,            # respondent age
    starts_with("B3_")       # child gender (needed only to rebuild total weighted pyramid)
  ) %>%
  rename(
    two_parents = A2.1_6,
    num_children = B1,
    respondent_gender = S1,
    respondent_age = S2
  )

# -----------------------------
# 1) Child-level schooling data
# -----------------------------
children_schooling <- data_selected %>%
  pivot_longer(
    cols = c(matches("^B2_\\d+$"), matches("^B5_\\d+_2$")),
    names_to = c(".value", "child_num"),
    names_pattern = "(B2|B5)_(\\d+)(?:_2)?"
  ) %>%
  mutate(child_num = as.numeric(child_num)) %>%
  filter(child_num <= num_children) %>%
  rename(
    age = B2,
    attends_ua_school = B5
  ) %>%
  filter(!is.na(age), !is.na(attends_ua_school)) %>%
  mutate(
    two_parents = ifelse(two_parents == 1, 1, 0),
    adjusted_Weight = if_else(two_parents == 1, Weight_2 / 2, Weight_2)
  )

# ----------------------------------------------------------
# 2) Rebuild total weighted pyramid population from example
#    (respondents + children, with child weights adjusted)
# ----------------------------------------------------------

# Child part for pyramid
children_pyramid <- data_selected %>%
  pivot_longer(
    cols = matches("^B[23]_\\d+$"),
    names_to = c(".value", "child_num"),
    names_pattern = "(B[23])_(\\d+)"
  ) %>%
  mutate(
    child_num = as.numeric(child_num),
    child_gender = case_when(
      B3 == 1 ~ 2,   # keep your example's reversal
      B3 == 2 ~ 1,
      TRUE ~ NA_real_
    )
  ) %>%
  filter(child_num <= num_children) %>%
  filter(!is.na(B2) & !is.na(B3)) %>%
  transmute(
    age = B2,
    gender = as.numeric(child_gender),
    child = 1,
    two_parents = ifelse(two_parents == 1, 1, 0),
    adjusted_Weight = if_else(ifelse(two_parents == 1, 1, 0) == 1, Weight_2 / 2, Weight_2)
  )

# Respondent part for pyramid
respondents_pyramid <- data_selected %>%
  transmute(
    age = respondent_age,
    gender = as.numeric(respondent_gender),
    child = 0,
    two_parents = ifelse(two_parents == 1, 1, 0),
    adjusted_Weight = Weight_2
  )

# Total weighted pyramid population (all refugees)
final_pyramid <- bind_rows(respondents_pyramid, children_pyramid)
total_weighted_pyramid <- sum(final_pyramid$adjusted_Weight, na.rm = TRUE)

# ----------------------------------------------------------
# 3) Requested statistics: ages 6-17 inclusive
# ----------------------------------------------------------
requested_stats <- children_schooling %>%
  mutate(
    is_6_17 = age >= 6 & age <= 17,
    not_attending_ua_school = attends_ua_school == 0
  ) %>%
  summarise(
    weighted_children_6_17_not_attending_ua_school =
      sum(adjusted_Weight[is_6_17 & not_attending_ua_school], na.rm = TRUE),

    weighted_children_6_17_total =
      sum(adjusted_Weight[is_6_17], na.rm = TRUE),

    share_among_children_6_17 =
      weighted_children_6_17_not_attending_ua_school / weighted_children_6_17_total,

    total_weighted_pyramid_all_refugees =
      total_weighted_pyramid,

    share_of_total_weighted_pyramid =
      weighted_children_6_17_not_attending_ua_school / total_weighted_pyramid
  )

requested_stats

# ----------------------------------------------------------
# 4) B5 for children aged 6-17 inclusive:
#    weighted % by country + overall
#    ignore countries where weighted sum of children < 30
# ----------------------------------------------------------

# Country labels
country_labels <- c(
  `1`  = "Україна",
  `2`  = "Бельгія",
  `3`  = "Болгарія",
  `4`  = "Чехія",
  `5`  = "Данія",
  `6`  = "Німеччина",
  `7`  = "Естонія",
  `8`  = "Ірландія",
  `9`  = "Греція",
  `10` = "Іспанія",
  `11` = "Франція",
  `12` = "Хорватія",
  `13` = "Італія",
  `14` = "Кіпр",
  `15` = "Латвія",
  `16` = "Литва",
  `17` = "Люксембург",
  `18` = "Угорщина",
  `19` = "Мальта",
  `20` = "Нідерланди",
  `21` = "Австрія",
  `22` = "Польща",
  `23` = "Португалія",
  `24` = "Румунія",
  `25` = "Словенія",
  `26` = "Словаччина",
  `27` = "Фінляндія",
  `28` = "Швеція",
  `29` = "Ісландія",
  `30` = "Ліхтенштейн",
  `31` = "Норвегія",
  `32` = "Швейцарія",
  `33` = "США",
  `34` = "Канада",
  `35` = "Велика Британія",
  `36` = "Інше"
)

# B5 option labels
b5_option_labels <- c(
  `1` = "1. Школа / коледж / технікум у країні перебування",
  `2` = "2. Дистанційна школа / коледж / технікум в Україні",
  `3` = "3. Дитсадок у країні перебування",
  `4` = "4. Університет у країні перебування",
  `5` = "5. Університет в Україні",
  `6` = "6. Інше навчання",
  `7` = "7. Не навчається"
)

# Base child file
children_base <- data %>%
  select(
    ID,
    Weight_2,
    S5,
    A2.1_6,
    B1,
    starts_with("B2_"),
    matches("^B5_\\d+_\\d+$")
  ) %>%
  rename(
    country_code = S5,
    two_parents = A2.1_6,
    num_children = B1
  )

# Child ages: one row = one child
children_age <- children_base %>%
  select(ID, Weight_2, country_code, two_parents, num_children, starts_with("B2_")) %>%
  pivot_longer(
    cols = starts_with("B2_"),
    names_to = "child_num",
    names_prefix = "B2_",
    values_to = "age"
  ) %>%
  mutate(
    child_num = as.numeric(child_num),
    age = as.numeric(age),
    two_parents = ifelse(two_parents == 1, 1, 0),
    adjusted_Weight = if_else(two_parents == 1, Weight_2 / 2, Weight_2),
    country = recode(as.character(country_code), !!!country_labels)
  ) %>%
  filter(child_num <= num_children) %>%
  filter(!is.na(age), age >= 6, age <= 17) %>%
  select(ID, child_num, age, country, adjusted_Weight)

# Child B5 matrix: one row = one child-option
children_b5_long <- children_base %>%
  select(ID, matches("^B5_\\d+_\\d+$")) %>%
  pivot_longer(
    cols = matches("^B5_\\d+_\\d+$"),
    names_to = c("child_num", "option_no"),
    names_pattern = "B5_(\\d+)_(\\d+)",
    values_to = "value"
  ) %>%
  mutate(
    child_num = as.numeric(child_num),
    option_no = as.numeric(option_no),
    value = as.numeric(value),
    option_label = recode(as.character(option_no), !!!b5_option_labels)
  )

# Join ages with B5
children_b5_6_17 <- children_age %>%
  inner_join(children_b5_long, by = c("ID", "child_num"))

# Denominator by country: weighted number of children 6-17
country_denoms <- children_age %>%
  group_by(country) %>%
  summarise(
    total_weighted_children_6_17 = sum(adjusted_Weight, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  filter(total_weighted_children_6_17 >= 30)

# Numerators by country and option
country_b5_pct_long <- children_b5_6_17 %>%
  semi_join(country_denoms, by = "country") %>%
  group_by(country, option_no, option_label) %>%
  summarise(
    weighted_n = sum(adjusted_Weight[value == 1], na.rm = TRUE),
    .groups = "drop"
  ) %>%
  left_join(country_denoms, by = "country") %>%
  mutate(
    pct = 100 * weighted_n / total_weighted_children_6_17
  ) %>%
  arrange(country, option_no)

# Overall denominator
overall_denom <- children_age %>%
  summarise(
    total_weighted_children_6_17 = sum(adjusted_Weight, na.rm = TRUE)
  ) %>%
  pull(total_weighted_children_6_17)

# Overall numerators
overall_b5_pct_long <- children_b5_6_17 %>%
  group_by(option_no, option_label) %>%
  summarise(
    weighted_n = sum(adjusted_Weight[value == 1], na.rm = TRUE),
    .groups = "drop"
  ) %>%
  mutate(
    country = "Загалом",
    total_weighted_children_6_17 = overall_denom,
    pct = 100 * weighted_n / total_weighted_children_6_17
  ) %>%
  select(country, option_no, option_label, weighted_n, total_weighted_children_6_17, pct)

# Final long table
country_b5_pct_long <- bind_rows(
  country_b5_pct_long,
  overall_b5_pct_long
) %>%
  arrange(country, option_no)

# Final wide table with percentages
country_b5_pct_wide <- country_b5_pct_long %>%
  select(country, option_label, pct) %>%
  pivot_wider(
    names_from = option_label,
    values_from = pct
  )

country_b5_pct_wide

country_b5_pct_wide <- country_b5_pct_long %>%
  select(country, total_weighted_children_6_17, option_label, pct) %>%
  distinct() %>%
  pivot_wider(names_from = option_label, values_from = pct) %>%
  rename(weighted_children_6_17 = total_weighted_children_6_17)

writexl::write_xlsx(country_b5_pct_wide, "b5_children_6_17_by_country.xlsx")
