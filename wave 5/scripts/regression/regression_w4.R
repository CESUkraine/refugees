# ==============================================================================
# 1. SETUP & LIBRARIES
# ==============================================================================
library(haven)      # For read_sav
library(margins)    # For margins
library(sjPlot)     # For plot_model
library(tidyverse)  # For data manip & plotting
library(expss)      # For lbl_na_if, add_labelled_class
library(MASS)       # For polr
library(broom)      # For tidy
library(writexl)    # For write_xlsx

# Load Data
refugees <- read_sav("901737_wave4.sav", encoding = 'CP1251')
write_csv(refugees, 'refugees_wave4.csv')

# ==============================================================================
# 2. DATA INSPECTION
# ==============================================================================
# Exclude those who left Ukraine in Jan 2024 or later
ref_no_2024_leave <- refugees |> filter(S4a < 26)

# Summary with weights and percentages (Replaces the redundant table/count calls)
ref_no_2024_leave |> 
  group_by(Q16) |> 
  summarise(
    weighted_count = sum(weight), 
    perc = weighted_count / sum(ref_no_2024_leave$weight)
  )

# ==============================================================================
# 3. DATA CLEANING
# ==============================================================================
# Recoding Intention to Return
# attributes(refugees$J2)$labels
refugees <- refugees |> 
  mutate(int_return = case_when(
    J2 == 1 ~ 5, J2 == 2 ~ 4, J2 == 5 ~ 3, J2 == 3 ~ 2, J2 == 4 ~ 1
  ))

# Handling Missing Data
assign_missing <- function(x) { 
  lbl_na_if(x, ~.lbl %in% c("Відмова відповідати", 'Відсутня відповідь',
                            'Важко сказати', 'Відмова від відповіді', 'немає відповіді'))
}
refugees <- map_df(refugees, assign_missing)
refugees <- add_labelled_class(refugees)

# Define Regions
west <- c(3, 7, 9, 14, 18, 20, 23, 25)
centre <- c(2, 12, 17, 24)
south <- c(1, 8, 15, 16)
east <- c(4, 5, 13, 21)
north <- c(6, 10, 11, 19, 26)
warzone <- c(5, 8, 10, 13, 15, 19, 21, 22, 26)

refugees <- refugees |> 
  mutate(
    region = case_when(
      Z1 %in% west ~ 1, Z1 %in% centre ~ 2, Z1 %in% north ~ 3,
      Z1 %in% south ~ 4, Z1 %in% east ~ 5
    ),
    war_zone = ifelse(Z1 %in% warzone, 1, 0)
  )

# Recode Employment
# Note: 'attend_integr_courses' removed as it was unused
refugees <- refugees |> 
  mutate(
    # Before War
    student = ifelse(E1_04 == 1, 1, 0),
    working = ifelse(E1_01 == 1 | E1_02 == 1 | E1_10 == 1, 1, 0),
    business = ifelse(E1_03 == 1, 1, 0),
    unemployed = ifelse(E1_08 == 1, 1, 0),
    out_of_labor = ifelse(E1_05 == 1 | E1_06 == 1 | E1_07 == 1, 1, 0),
    
    # Now
    student_here_now = ifelse(E2_08 == 1, 1, 0),
    student_remote_now = ifelse(E2_07 == 1, 1, 0),
    working_now = ifelse(E2_05 == 1 | E2_04 == 1, 1, 0),
    business_here_now = ifelse(E2_11 == 1, 1, 0),
    business_remote_now = ifelse(E2_10 == 1, 1, 0),
    unemployed_now = ifelse(E2_06 == 1, 1, 0),
    work_remotely_now = ifelse(E2_01 == 1 | E2_02 == 1, 1, 0)
  )

# Labels and Factors
refugees <- refugees |> 
  rename(
    Студент = student_here_now,
    Працює = working_now,
    Безробітний = unemployed_now
  ) |> 
  mutate(
    region = factor(region, levels = 1:5, labels = c('Захід', 'Центр', 'Північ', 'Південь', 'Схід')),
    E3.1 = factor(unclass(E3), levels = 1:6, 
                  labels = c("Найнижчий дохід до війни", 'Дохід до війни:2', 'Дохід до війни:3',
                             'Дохід до війни:4', 'Дохід до війни:5', "Дохід до війни:6")),
    E3.2 = factor(unclass(E4), levels = 1:6, 
                  labels = c("Найнижчий дохід зараз", 'Дохід зараз:2', 'Дохід зараз:3',
                             'Дохід зараз:4', 'Дохід зараз:5', "Дохід зараз:6"))
  )

# Recode Country
refugees <- refugees |> 
  mutate(Country = case_when(
    S5 == 7 ~ 1, S5 == 8 ~ 2, S5 == 17 ~ 3, S5 == 3 ~ 4, 
    S5 == 12 ~ 5, S5 == 18 ~ 6, TRUE ~ 7
  )) |> 
  mutate(Country = factor(Country, levels = 1:7, 
                          labels = c('Німеччина', "Польща", "Чехія", "Велика Британія", "США", "Канада", "Інші"))) |> 
  rename(Бізнес = business_here_now)

# Demographics
refugees <- refugees |> 
  mutate(
    male = ifelse(S1 == 1, 1, 0),
    Z4 = factor(unclass(Z4), levels = 1:4, 
                labels = c('Незакінчена середня', 'Загальна середня', 'Середня спеціальна', 'Вища або незакінчена вища')),
    Z2 = factor(unclass(Z2), levels = 1:3, labels = c('Село', ' Селище міського типу', 'Місто'))
  )

attr(refugees$Z4, "label") <- "Освіта"
attr(refugees$S2, "label") <- "Вік"

# ==============================================================================
# 4. REGRESSION (Ukrainian)
# ==============================================================================
m1 <- polr(factor(int_return) ~ male + S2 + Country + factor(Q6) + A2.1_2 + 
             student + working + business + unemployed + 
             Студент + student_remote_now + Працює + Бізнес + business_remote_now + 
             Безробітний + work_remotely_now + region + war_zone + Z2 + Z4 + Z5_1 + E3.1 + E3.2,
           data = refugees, weights = weight)

# Export Coefficients
coefs <- tidy(m1, exponentiate = TRUE, conf.int = TRUE, conf.level = 0.90)
var_names <- c('male', 'S2', 'work_remotely_now', 'CountryПольща', 'CountryЧехія', 
               'CountryВелика Британія', 'CountryСША', 'CountryКанада', 'CountryІнші', 
               'Z2 Селище міського типу', 'Z2Місто', 'E3.1Дохід до війни:2', 
               'E3.1Дохід до війни:3', 'E3.1Дохід до війни:4', 'E3.1Дохід до війни:5', 
               'E3.1Дохід до війни:6', 'Z4Загальна середня', 'Z4Середня спеціальна', 
               'Z4Вища або незакінчена вища')

coefs |> filter(term %in% var_names) |> write_xlsx('coefs.xlsx')

# Manual Odds Calculations
results <- exp(cbind(estimate = coef(m1), confint.default(m1))) # confint.default is faster approx for wald
results <- as.data.frame(results) |> rownames_to_column(var = "term") |> filter(term %in% var_names)
write_xlsx(results, 'results.xlsx')

# Plot
plot_odds <- plot_model(m1, rm.terms = c('1|2', '2|3', '3|4', '4|5'),
                        terms = c('male', 'S2', 'work_remotely_now', 
                                  'Country [Польща, Чехія, Велика Британія, США, Канада, Інші]', 
                                  'Z2 Селище міського типу', 'Z2Місто', 
                                  'E3.1 [Дохід до війни:2,Дохід до війни:3,Дохід до війни:4,Дохід до війни:5,Дохід до війни:6]', 
                                  'Z4 [ Загальна середня,Середня спеціальна,Вища або незакінчена вища]'),
                        colors = c("#73B932", "#00509b"), show.values = TRUE, 
                        title = '', vline.color = 'black', value.offset = 0.3) + theme_bw()
ggsave('plot_odds.png', plot_odds, width = 2000, height = 1700, units = 'px')

# ==============================================================================
# 5. REGRESSION (English)
# ==============================================================================
refugees_eng <- refugees |> 
  rename(Student = Студент, Working = Працює, Unemployed = Безробітний,
         student_before = student, working_before = working, unemployed_before = unemployed) |> 
  mutate(
    region = factor(region, levels = c("Захід", "Центр", "Північ", "Південь", "Схід"),
                    labels = c('West', 'Centre', 'North', 'South', 'East')),
    E3.1 = factor(E3.1, labels = c("Lowest pre-war income", 'Pre-war income:2', 'Pre-war income:3', 
                                   'Pre-war income:4', 'Pre-war income:5', "Pre-war income:6")),
    E3.2 = factor(E3.2, labels = c("Lowest income now", 'Income now:2', 'Income now:3', 
                                   'Income now:4', 'Income now:5', "Income now:6"))
  )

m1_eng <- polr(factor(int_return) ~ factor(S1) + S2 + factor(Country) + A2_2 + 
                 student_before + working_before + business + unemployed_before + 
                 Student + student_remote_now + Working + business_here_now + business_remote_now + 
                 Unemployed + work_remotely_now + region + war_zone + factor(Z2) + factor(Z4) + Z5_1 + E3.1 + E3.2,
               data = refugees_eng)

plot_odds_eng <- plot_model(m1_eng, rm.terms = c('1|2', '2|3', '3|4', '4|5'),
                            terms = c('region [Centre,North,South,East]', 'Student', 'Unemployed', 'Working', 
                                      'E3.1 [Pre-war income:2,Pre-war income:3,Pre-war income:4,Pre-war income:5,Pre-war income:6]', 
                                      'E3.2 [Income now:2,Income now:3,Income now:4,Income now:5,Income now:6]'),
                            colors = c("#73B932", "#00509b"), show.values = TRUE, 
                            title = '', vline.color = 'black', value.offset = 0.3) + theme_bw()
ggsave('plot_odds_eng.png', plot_odds_eng, width = 2000, height = 1700, units = 'px')