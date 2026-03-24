library(tidyverse)
library(readxl)
library(ggplot2)

# ==============================================================================
# 1. DATA IMPORT & CLEANING
# ==============================================================================

raw_data <- read_excel("regression_results.xlsx")

plot_data <- raw_data |>
  filter(coef.type == "coefficient") |>
  mutate(
    # Create readable display labels (EN)
    label = case_when(
      # Demographics
      term == "male" ~ "Male",
      term == "age" ~ "Age",
      term == "with_children" ~ "Has children",
      term == "spouse_in_UA" ~ "Partner in Ukraine",

      # Education
      term == "educationСередня спеціальна" ~ "Education: Vocational",
      term == "educationВища" ~ "Education: Higher",

      # Origin / Region
      term == "war_zone" ~ "From war zone",
      term == "settlement_typeСелище" ~ "Settlement: Town",
      term == "settlement_typeНевелике місто" ~ "Settlement: Small city",
      term == "settlement_typeВелике місто" ~ "Settlement: Large city",
      term == "regionЦентр" ~ "Region: Center",
      term == "regionПівніч" ~ "Region: North",
      term == "regionПівдень" ~ "Region: South",
      term == "regionСхід" ~ "Region: East",

      # Employment / Activity
      term == "student_here_now" ~ "Student (host country)",
      term == "student_remote_now" ~ "Student (remote)",
      term == "working_now" ~ "Working (host country)",
      term == "working_remotely_now" ~ "Working (remote)",
      term == "business_here_now" ~ "Business (host country)",
      term == "business_remote_now" ~ "Business (remote)",
      term == "unemployed_now" ~ "Unemployed",

      # Financials
      term == "remittances_to_UA" ~ "Remittances to Ukraine",

      term == "income_beforeДохід до війни:2" ~ "Pre-war income: 2",
      term == "income_beforeДохід до війни:3" ~ "Pre-war income: 3",
      term == "income_beforeДохід до війни:4" ~ "Pre-war income: 4",
      term == "income_beforeДохід до війни:5" ~ "Pre-war income: 5",
      term == "income_beforeДохід до війни:6" ~ "Pre-war income: 6",

      term == "income_nowДохід зараз:2" ~ "Current income: 2",
      term == "income_nowДохід зараз:3" ~ "Current income: 3",
      term == "income_nowДохід зараз:4" ~ "Current income: 4",
      term == "income_nowДохід зараз:5" ~ "Current income: 5",
      term == "income_nowДохід зараз:6" ~ "Current income: 6",

      # Host Countries
      term == "CountryПольща" ~ "Poland",
      term == "CountryЧехія" ~ "Czechia",
      term == "CountryВелика Британія" ~ "United Kingdom",
      term == "CountryСША" ~ "United States",
      term == "CountryКанада" ~ "Canada",
      term == "CountryІнші" ~ "Other countries",

      TRUE ~ NA_character_
    )
  ) |>
  filter(!is.na(label)) |>
  mutate(
    # Color based strictly on CI bounds relative to 1
    color_group = case_when(
      conf.low > 1  ~ "blue",   # Entire CI > 1
      conf.high < 1 ~ "green",  # Entire CI < 1
      TRUE          ~ "grey"    # CI includes 1
    ),
    stars = case_when(
      p.value < 0.001 ~ "***",
      p.value < 0.01  ~ "**",
      p.value < 0.05  ~ "*",
      TRUE            ~ ""
    ),
    text_label = paste0(sprintf("%.2f", estimate), " ", stars)
  )

# ==============================================================================
# 2. ORDERING (FIXED: must match English labels)
# ==============================================================================

desired_order <- c(
  "Current income: 6", "Current income: 5", "Current income: 4",
  "Current income: 3", "Current income: 2",
  "Pre-war income: 6", "Pre-war income: 5", "Pre-war income: 4",
  "Pre-war income: 3", "Pre-war income: 2",
  "Remittances to Ukraine",
  "Unemployed",
  "Business (remote)", "Business (host country)",
  "Working (remote)", "Working (host country)",
  "Student (remote)", "Student (host country)",
  "Other countries", "Canada", "United States", "United Kingdom", "Czechia", "Poland",
  "From war zone",
  "Region: East", "Region: South", "Region: North", "Region: Center",
  "Settlement: Large city", "Settlement: Small city", "Settlement: Town",
  "Education: Higher", "Education: Vocational",
  "Partner in Ukraine", "Has children", "Age", "Male"
)

plot_data <- plot_data |>
  mutate(
    label = factor(label, levels = desired_order)
  ) |>
  # Drop anything not present in desired_order (prevents NA factor levels from breaking y scale)
  filter(!is.na(label))

# ==============================================================================
# 3. VISUALIZATION
# ==============================================================================

col_blue  <- "#73B932"
col_green <- "#AF2D5A"
col_grey  <- "#656567"

p <- ggplot(plot_data, aes(x = estimate, y = label, color = color_group)) +
  geom_vline(xintercept = 1, color = "gray40", linewidth = 0.6) +
  geom_errorbarh(aes(xmin = conf.low, xmax = conf.high), height = 0.2, linewidth = 0.8) +
  geom_point(size = 4) +
  geom_text(aes(label = text_label), vjust = -1, show.legend = FALSE, size = 4) +
  scale_x_log10(
    breaks = c(0.1, 0.5, 1, 2, 5, 10),
    labels = c("0.1", "0.5", "1", "2", "5", "10")
  ) +
  scale_color_manual(
    values = c(
      "green" = col_green,
      "blue"  = col_blue,
      "grey"  = col_grey
    )
  ) +
  theme_minimal() +
  labs(x = "Odds ratios", y = NULL) +
  theme(
    panel.border = element_rect(color = "black", fill = NA, linewidth = 1),
    panel.grid.minor = element_blank(),
    panel.grid.major.y = element_line(color = "grey95"),
    axis.text.y = element_text(size = 14, color = "gray20"),
    axis.text.x = element_text(size = 14, color = "gray20"),
    axis.title.x = element_text(size = 14, margin = margin(t = 10)),
    legend.position = "none"
  )

# Arrow placement settings
arrow_y_pos <- -3
text_y_pos  <- -3.5

p_final <- p +
  coord_cartesian(
    clip = "off",
    ylim = c(1, length(desired_order)),
    xlim = c(0.05, 20)
  ) +
  # --- LEFT SIDE ---
  annotate(
    "segment",
    x = 0.8, xend = 0.2,
    y = arrow_y_pos, yend = arrow_y_pos,
    arrow = arrow(length = unit(0.5, "cm")),
    color = col_green,
    linewidth = 2
  ) +
  annotate(
    "text",
    x = 0.4,
    y = text_y_pos,
    label = "Less likely\nto return",
    lineheight = 0.9,
    hjust = 0.5, vjust = 1,
    size = 5,
    color = "gray30"
  ) +
  # --- RIGHT SIDE ---
  annotate(
    "segment",
    x = 1.25, xend = 5,
    y = arrow_y_pos, yend = arrow_y_pos,
    arrow = arrow(length = unit(0.5, "cm")),
    color = col_blue,
    linewidth = 2
  ) +
  annotate(
    "text",
    x = 2.5,
    y = text_y_pos,
    label = "More likely\nto return",
    lineheight = 0.9,
    hjust = 0.5, vjust = 1,
    size = 5,
    color = "gray30"
  ) +
  theme(plot.margin = margin(10, 10, 80, 10))

ggsave(
  "forest_plot_corrected_ENG.png",
  p_final,
  width = 10,
  height = 12,
  dpi = 300,
  bg = "white"
)

print(p_final)