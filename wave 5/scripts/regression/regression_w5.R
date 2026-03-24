# ==============================================================================
# 1. SETUP & LIBRARIES
# ==============================================================================
# install.packages(c("haven", "margins", "sjPlot", "tidyverse", "expss", "MASS", "broom", "writexl"))
library(haven)      # For read_sav
library(margins)    # For margins
library(sjPlot)     # For plot_model
library(tidyverse)  # For data manip & plotting
library(expss)      # For lbl_na_if, add_labelled_class
library(MASS)       # For polr
library(broom)      # For tidy
library(writexl)    # For write_xlsx

# Load Data
refugees <- read_sav("901849.sav", encoding = 'CP1251')

# ==============================================================================
# 3. DATA CLEANING
# ==============================================================================
# Recoding Intention to Return: "Точно планую" becomes 5; "Точно не планую" becomes 1; "Важко сказати" is in the middle (3)
# IMPORTANT for ordinal logit!
refugees <- refugees |> 
  mutate(int_return = case_when(
    J2 == 1 ~ 5, J2 == 2 ~ 4, J2 == 5 ~ 3, J2 == 3 ~ 2, J2 == 4 ~ 1
  ))

# Handling Missing Data
# assign_missing <- function(x) { 
#   lbl_na_if(x, ~.lbl %in% c("Відмова відповідати", 'Відсутня відповідь',
#                             'Важко сказати', 'Відмова від відповіді', 'немає відповіді'))
# }
# refugees <- map_df(refugees, assign_missing)
# refugees <- add_labelled_class(refugees)

# Define Regions
attributes(refugees$Z1)$labels # double-check

west <- c(3, 7, 9, 14, 18, 20, 23, 25) # Volyn, Zakarpattia, Ivano-Frank, Lviv, Rivne, Ternopil, Khmelnytskyi, Chernivtsi
centre <- c(2, 12, 17, 24) # Vinnytsia, Kirovohrad, Poltava, Cherkasy
south <- c(1, 8, 15, 16, 22) # Crimea, Zaporizhzhia, Mykolaiv, Odesa, KHERSON???
east <- c(4, 5, 13, 21) # Dnipro, Donetsk, Luhansk, Kharkiv
north <- c(6, 10, 11, 19, 26) # Zhytomyr, Kyiv oblast, Kyiv city, Sumy, Chernihiv

warzone <- c(5, 8, 10, 13, 15, 19, 21, 22, 26) # Donetsk, Zaporizhzhia, Kyiv oblast, Luhansk, Mykolaiv, Sumy, Kharkiv, Kherson, Chernihiv 

refugees <- refugees |> 
  mutate(
    region = case_when(
      Z1 %in% west ~ 1, Z1 %in% centre ~ 2, Z1 %in% north ~ 3,
      Z1 %in% south ~ 4, Z1 %in% east ~ 5
    ),
    war_zone = ifelse(Z1 %in% warzone, 1, 0)
  )

# Recode Employment
refugees <- refugees |> 
  mutate(
    # Before War
    student_before = ifelse(E1_04 == 1, 1, 0),
    working_before = ifelse(E1_01 == 1 | E1_02 == 1, 1, 0), # full-time employed, part-time employed
    business_before = ifelse(E1_03 == 1, 1, 0), # self-employed/business owner
    unemployed_before = ifelse(E1_08 == 1, 1, 0), # unemployed
    out_of_labor_before = ifelse(E1_05 == 1 | E1_06 == 1 | E1_07 == 1, 1, 0), # don't have a job and don't look for job; pensioner; unfit for work
    
    # Now
    student_here_now = ifelse(E2_08 == 1, 1, 0), # studying at institutions/programs which are available for locals
    student_remote_now = ifelse(E2_07 == 1, 1, 0), # distant learning in UA education instituitons
    working_now = ifelse(E2_04 == 1 | E2_05 == 1, 1, 0), # full-time, part-time
    business_here_now = ifelse(E2_11 == 1, 1, 0), # business abroad
    business_remote_now = ifelse(E2_10 == 1, 1, 0), # business in Ukraine
    unemployed_now = ifelse(E2_06 == 1, 1, 0),
    working_remotely_now = ifelse(E2_01 == 1 | E2_02 == 1, 1, 0) # the same Ukrainian job, just remotely; a new Ukrainian remote job
  )

# Labels and Factors
refugees <- refugees |> 
  # rename(
  #   Студент = student_here_now,
  #   Працює = working_now,
  #   Безробітний = unemployed_now
  # ) |> 
  mutate(
    region = factor(region, levels = 1:5, labels = c('Захід', 'Центр', 'Північ', 'Південь', 'Схід')),
    income_before = factor(unclass(E3), levels = 1:6, 
                  labels = c("Найнижчий дохід до війни", 'Дохід до війни:2', 'Дохід до війни:3',
                             'Дохід до війни:4', 'Дохід до війни:5', "Дохід до війни:6")),
    income_now = factor(unclass(E3.1), levels = 1:6, 
                  labels = c("Найнижчий дохід зараз", 'Дохід зараз:2', 'Дохід зараз:3',
                             'Дохід зараз:4', 'Дохід зараз:5', "Дохід зараз:6"))
  )

# Recode Country: check with attributes(refugees$S5)$labels
refugees <- refugees |>
  mutate(Country = case_when(
    S5 == 6 ~ 1, S5 == 22 ~ 2, S5 == 4 ~ 3, S5 == 35 ~ 4,
    S5 == 33 ~ 5, S5 == 34 ~ 6, TRUE ~ 7
  )) |>
  mutate(Country = factor(Country,
    levels = 1:7,
    labels = c("Німеччина", "Польща", "Чехія", "Велика Британія", "США", "Канада", "Інші")
  ))

# Demographics
refugees <- refugees |> 
  mutate(
    male = ifelse(S1 == 1, 1, 0),
    Z4 = factor(unclass(Z4), levels = 1:7, 
                labels = c('Незакінчена середня', 'Загальна середня', 'Середня спеціальна', 'Незакінчена вища', 'Бакалавр/спеціаліст', 'Магістр', 'Науковий ступінь')),
    Z2 = factor(unclass(Z2), levels = 1:4, labels = c('Село', ' Селище', 'Невелике місто', 'Велике місто'))
  )

refugees <- refugees |>
  mutate(
    Z4_grouped = case_match(
      as.numeric(Z4),
      c(1, 2) ~ "Середня",
      3 ~ "Середня спеціальна",
      c(4, 5, 6, 7) ~ "Вища"
    ),
    # Convert the result to a factor to set the specific order
    education = factor(Z4_grouped, levels = c("Середня", "Середня спеціальна", "Вища")),
    Z2 = factor(unclass(Z2), levels = 1:4, labels = c("Село", "Селище", "Невелике місто", "Велике місто"))
  )

attr(refugees$Z4, "label") <- "Освіта"
attr(refugees$S2, "label") <- "Вік"

refugees <- refugees |> 
  mutate(
    # New year variable with the 2021 floor
    year = ifelse(S4.01 < 2021, 2021, S4.01)
  )

refugees <- refugees |>
  mutate(
    remittances_to_UA = ifelse(Z7 == 1, 1, 0)
  )

refugees <- refugees |> 
  mutate(
    # Check if ANY of the columns B2_1...B2_5 fall into the specific range
    # The '+' symbol converts TRUE/FALSE to 1/0
    child_0_5   = +if_any(c(B2_1, B2_2, B2_3, B2_4, B2_5), ~ .x >= 0 & .x <= 5),
    child_6_10  = +if_any(c(B2_1, B2_2, B2_3, B2_4, B2_5), ~ .x >= 6 & .x <= 10),
    child_11_14 = +if_any(c(B2_1, B2_2, B2_3, B2_4, B2_5), ~ .x >= 11 & .x <= 14),
    child_15_17 = +if_any(c(B2_1, B2_2, B2_3, B2_4, B2_5), ~ .x >= 15 & .x <= 17)
  ) |> 
  # If a respondent has NO children (or B2 columns are all NA), the result is NA.
  # We must replace these NAs with 0 for the regression model to work.
  mutate(across(c(child_0_5, child_6_10, child_11_14, child_15_17), ~replace_na(., 0)))

refugees <- refugees |>
  rename(age = S2,
        with_children = A2.1_2,
        settlement_type = Z2,
        spouse_in_UA = Z5_1)

# ==============================================================================
# 4. REGRESSION (Ukrainian)
# ==============================================================================
print("===")
# ==========================================================================
# 4. REGRESSION (FIXED)
# ==========================================================================

# 1. Run the model (same as before)
m1 <- polr(
  factor(int_return) ~ male + age + Country + with_children +
    # student_before + working_before + business_before + unemployed_before +
    student_here_now + student_remote_now + working_now + business_here_now +
    business_remote_now + unemployed_now + working_remotely_now +
    region + war_zone +
    settlement_type + education + spouse_in_UA + income_before + income_now + remittances_to_UA,
  data = refugees,
  weights = Weight_2,
  Hess = TRUE # This calculates the Hessian immediately
)

# 2. Get the Tidy Results (Odds Ratios and Confidence Intervals)
# note: this creates columns: term, estimate, std.error, statistic, conf.low, conf.high
# exponentiate = TRUE gives you Odds Ratios (matches your exp() logic)
# conf.int = TRUE gives confidence intervals
# conf.level = 0.9 matches your 90% request
results <- broom::tidy(m1, conf.int = TRUE, exponentiate = TRUE, conf.level = 0.9)

# 3. Calculate P-values and Significance Stars
results <- results |>
  mutate(
    # Calculate P-value based on the t-statistic (Wald test approximation)
    p.value = pnorm(abs(statistic), lower.tail = FALSE) * 2,

    # Optional: Add stars for easier reading
    significance = case_when(
      # p.value < 0.001 ~ "***",
      p.value < 0.01 ~ "***",
      p.value < 0.05 ~ "**",
      p.value < 0.1 ~ "*",
      TRUE ~ ""
    )
  )

# 4. Reorder columns for a professional look
results <- results |>
  select(term, estimate, std.error, statistic, p.value, significance, conf.low, conf.high)

# 5. Save
write_xlsx(results, "regression_results.xlsx")
