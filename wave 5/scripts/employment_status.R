# install.packages(c("haven", "writexl", "dplyr", "tidyr"))

library(haven)
library(writexl)
library(dplyr)
library(tidyr)

input_file <- "901849.sav"
output_file <- "employment_status.xlsx"

data <- read_sav(input_file, encoding = "cp1251")

# Since we want to build exclusive categories based on multiple-choice answers, we choose the following approach.
# If people are labour-involved (at least one of the 6 categories), the they are classified as labour-involved.
# Even if they also belong to other categories (e.g. studying or pensioners).
# Then, similarly, similarly, for the rest: if they are not labour-involved but studying (3 categories), they are students.
# On a residual basis, respondents are classified into the rest of categories.

# THE LIST AND NUMBERING BELOW CAN CHANGE FROM SURVEY TO SURVEY. PLEASE CHECK!!!!!!

# LABOUR-INVOLVED
# E2_01 "Продовжую дистанційно працювати на роботі, на якій працював\\ла до війни"
# E2_02 "Знайшов\\ла нову роботу в Україні і працюю дистанційно"
# E2_04 "Влаштувався\\лась на роботу в цій країні на повний день"
# E2_05 "Влаштувався\\лась на роботу в цій країні на неповний день"
# E2_10 "Займаюсь бізнесом в Україні"
# E2_11 "Займаюсь бізнесом в країні перебування"

# STUDENTS
# E2_07 "Навчаюсь дистанційно в українських навчальних закладах"
# E2_08 "Навчаюсь на загальних навчальних програмах в навчальних закладах країні перебування, які доступні і для місцевих жителів"
# E2_09 "Навчаюсь на спеціальних курсах для українців в країні перебування (наприклад, мовні курси, інтеграційні курси)"

# DISABLED AND PENSIONERS
# E2_12 "Пенсіонер(-ка)"
# E2_13 "Непрацездатний(-а) (включаючи людей з обмеженими можливостями)"

# DISCOURAGED (NOT WORKING AND NOT LOOKING FOR WORK)
# E2_14 "Не працюю та не шукаю роботу (вирішую побутові питання, адаптуюсь до життя в новій країні\\планую незабаром повертатися до України, доглядаю за дітьми або іншими членами родини або не працюю з інших причин)"

# UNEMPLOYED
# E2_03 "Займаюсь волонтерством"
# E2_06 "Шукаю роботу в цій країні"
# E2_15 "Інше"

# ---- Columns by bucket ----
labour_cols <- c("E2_01", "E2_02", "E2_04", "E2_05", "E2_10", "E2_11")
student_cols <- c("E2_07", "E2_08", "E2_09")
dp_cols <- c("E2_12", "E2_13") # disabled & pensioners
disc_cols <- c("E2_14") # discouraged (not working & not looking)
unemp_cols <- c("E2_03", "E2_06", "E2_15") # unemployed/other/non-working but not in disc

# ---- Helper: any selected in a set of multi-select columns ----
# Treats any positive numeric as "selected" (robust to labelled doubles 0/1, 1/2, etc.)
# Updated to work with data frames provided by pick()
selected_any <- function(df) {
    # Coerce to numeric, then to 0/1, NAs -> 0
    mat_bin <- lapply(df, function(v) {
        v_num <- suppressWarnings(as.numeric(v))
        ifelse(is.na(v_num), 0, ifelse(v_num > 0, 1, 0))
    })
    rowSums(as.data.frame(mat_bin)) > 0
}

# ---- Build flags ----
flags <- data %>%
    mutate(
        is_labour   = selected_any(pick(all_of(labour_cols))),
        is_student  = selected_any(pick(all_of(student_cols))),
        is_dp       = selected_any(pick(all_of(dp_cols))),
        is_disc     = selected_any(pick(all_of(disc_cols))),
        is_unemp    = selected_any(pick(all_of(unemp_cols)))
    ) %>%
    # ---- Exclusive category by priority ----
    mutate(
        employment_status = case_when(
            is_labour ~ "Залучені до праці",
            !is_labour & is_student ~ "Студенти/навчання",
            !is_labour & !is_student & is_dp ~ "Пенсіонери/непрацездатні",
            !is_labour & !is_student & !is_dp & is_disc ~ "Не працює і не шукає",
            !is_labour & !is_student & !is_dp & !is_disc & is_unemp ~ "Безробітні/інше",
            TRUE ~ "Інше/невизначено"
        )
    )

# Optional: lock factor order for nicer output
status_levels <- c(
    "Залучені до праці",
    "Студенти/навчання",
    "Пенсіонери/непрацездатні",
    "Не працює і не шукає",
    "Безробітні/інше",
    "Інше/невизначено"
)

flags <- flags %>%
    mutate(employment_status = factor(employment_status, levels = status_levels))

# ---- Weighted summary ----
# Assumes weight column is named `Weight`. Adjust if different.
summary_tbl <- flags %>%
    group_by(employment_status) %>%
    summarise(
        n = n(),
        weight_sum = sum(Weight_2, na.rm = TRUE),
        .groups = "drop"
    ) %>%
    mutate(
        weight_share = weight_sum / sum(weight_sum, na.rm = TRUE)
    ) %>%
    arrange(match(employment_status, status_levels))

# ---- (Optional) diagnostics sheet: counts of raw combinations ----
# Helps sanity-check overlap patterns prior to exclusivisation
diag_tbl <- flags %>%
    transmute(
        is_labour, is_student, is_dp, is_disc, is_unemp, Weight_2
    ) %>%
    group_by(is_labour, is_student, is_dp, is_disc, is_unemp) %>%
    summarise(
        n = n(),
        weight_sum = sum(Weight_2, na.rm = TRUE),
        .groups = "drop"
    ) %>%
    arrange(desc(weight_sum))

# ---- Export ----
write_xlsx(
    list(
        "employment_status_summary" = summary_tbl,
        "diagnostics_overlap"       = diag_tbl
    ),
    path = output_file
)

cat("Written:", output_file, "\n")
