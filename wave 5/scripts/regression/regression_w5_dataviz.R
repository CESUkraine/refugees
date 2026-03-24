library(tidyverse)
library(readxl)
library(ggplot2)

# ==============================================================================
# 1. DATA IMPORT & CLEANING
# ==============================================================================

# Read the Excel file
raw_data <- read_excel("regression_results.xlsx") 

plot_data <- raw_data |> 
  filter(coef.type == "coefficient") |> 
  mutate(
    # Create readable display labels
    label = case_when(
      # Demographics
      term == "male" ~ "Чоловік",
      term == "age" ~ "Вік",
      term == "with_children" ~ "Має дітей",
      term == "spouse_in_UA" ~ "Партнер/ка в Україні",
      
      # Education
      term == "educationСередня спеціальна" ~ "Освіта: Середня спеціальна",
      term == "educationВища" ~ "Освіта: Вища",
      
      # Origin / Region
      term == "war_zone" ~ "Із зони бойових дій",
      term == "settlement_typeСелище" ~ "Тип поселення: Селище", 
      term == "settlement_typeНевелике місто" ~ "Тип поселення: Мале місто", 
      term == "settlement_typeВелике місто" ~ "Тип поселення: Велике місто",
      term == "regionЦентр" ~ "Регіон: Центр",
      term == "regionПівніч" ~ "Регіон: Північ",
      term == "regionПівдень" ~ "Регіон: Південь",
      term == "regionСхід" ~ "Регіон: Схід",
      
      # Employment / Activity
      term == "student_here_now" ~ "Навчається (в країні переб.)",
      term == "student_remote_now" ~ "Навчається (дистанційно)",
      term == "working_now" ~ "Працює (в країні переб.)",
      term == "working_remotely_now" ~ "Працює (дистанційно)",
      term == "business_here_now" ~ "Має бізнес (в країні переб.)",
      term == "business_remote_now" ~ "Має бізнес (дистанційно)",
      term == "unemployed_now" ~ "Безробітний(а)",
      
      # Financials
      term == "remittances_to_UA" ~ "Перекази в Україну",
      
      term == "income_beforeДохід до війни:2" ~ "Дохід до:2",
      term == "income_beforeДохід до війни:3" ~ "Дохід до:3",
      term == "income_beforeДохід до війни:4" ~ "Дохід до:4",
      term == "income_beforeДохід до війни:5" ~ "Дохід до:5",
      term == "income_beforeДохід до війни:6" ~ "Дохід до:6",
      
      term == "income_nowДохід зараз:2" ~ "Дохід зараз:2",
      term == "income_nowДохід зараз:3" ~ "Дохід зараз:3",
      term == "income_nowДохід зараз:4" ~ "Дохід зараз:4",
      term == "income_nowДохід зараз:5" ~ "Дохід зараз:5",
      term == "income_nowДохід зараз:6" ~ "Дохід зараз:6",
      
      # Host Countries
      term == "CountryПольща" ~ "Польща",
      term == "CountryЧехія" ~ "Чехія",
      term == "CountryВелика Британія" ~ "Велика Британія",
      term == "CountryСША" ~ "США",
      term == "CountryКанада" ~ "Канада",
      term == "CountryІнші" ~ "Інші країни",
      
      TRUE ~ NA_character_
    )
  ) |>
  filter(!is.na(label)) |> 
  mutate(
    # --- UPDATED COLOR LOGIC BASED ON CI BOUNDS ---
    # We ignore p.value here and look strictly at whether the interval crosses 1.
    color_group = case_when(
      conf.low > 1  ~ "blue",   # Entire CI is greater than 1
      conf.high < 1 ~ "green",  # Entire CI is less than 1
      TRUE          ~ "grey"    # CI includes 1
    ),
    
    stars = case_when(
      p.value < 0.001 ~ "***",
      p.value < 0.01 ~ "**",
      p.value < 0.05 ~ "*",
      # Optional: Add dot for p < 0.1 if you want to match the data's implication
      # p.value < 0.1 ~ ".", 
      TRUE ~ ""
    ),
    text_label = paste0(sprintf("%.2f", estimate), " ", stars)
  )

# ==============================================================================
# 2. ORDERING
# ==============================================================================

desired_order <- c(
  "Дохід зараз:6", "Дохід зараз:5", "Дохід зараз:4", 
  "Дохід зараз:3", "Дохід зараз:2",
  "Дохід до:6", "Дохід до:5", "Дохід до:4", 
  "Дохід до:3", "Дохід до:2",
  "Перекази в Україну",
  "Безробітний(а)",
  "Має бізнес (дистанційно)", "Має бізнес (в країні переб.)",
  "Працює (дистанційно)", "Працює (в країні переб.)",
  "Навчається (дистанційно)", "Навчається (в країні переб.)",
  "Інші країни", "Канада", "США", "Велика Британія", "Чехія", "Польща",
  "Із зони бойових дій",
  "Регіон: Схід", "Регіон: Південь", "Регіон: Північ", "Регіон: Центр",
  "Тип поселення: Велике місто", "Тип поселення: Мале місто", "Тип поселення: Селище",
  "Освіта: Вища", "Освіта: Середня спеціальна",
  "Партнер/ка в Україні", "Має дітей", "Вік", "Чоловік"
)

plot_data$label <- factor(plot_data$label, levels = desired_order)

# ==============================================================================
# 3. VISUALIZATION
# ==============================================================================

# Define Colors
col_blue <- "#73B932"
col_green <- "#AF2D5A"
col_grey <- "#656567" 

p <- ggplot(plot_data, aes(x = estimate, y = label, color = color_group)) +
  # Darker vertical line at 1
  geom_vline(xintercept = 1, color = "gray40", linewidth = 0.6) +
  
  geom_errorbarh(aes(xmin = conf.low, xmax = conf.high), height = 0.2, linewidth = 0.8) +
  geom_point(size = 4) +
  geom_text(aes(label = text_label), vjust = -1, show.legend = FALSE, size = 4) +
  scale_x_log10(breaks = c(0.1, 0.5, 1, 2, 5, 10), labels = c("0.1", "0.5", "1", "2", "5", "10")) +
  scale_color_manual(values = c(
    "green" = col_green, 
    "blue"  = col_blue, 
    "grey"  = col_grey
  )) +
  theme_minimal() +
  labs(x = "Odds Ratios", y = NULL) +
  theme(
    # Chart Border
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
text_y_pos <- -3.5 

p_final <- p +
  coord_cartesian(
    clip = "off",
    ylim = c(1, length(desired_order)),
    xlim = c(0.05, 20)
  ) +
  # --- LEFT SIDE ---
  annotate("segment",
    x = 0.8, xend = 0.2, 
    y = arrow_y_pos, yend = arrow_y_pos,
    arrow = arrow(length = unit(0.5, "cm")),
    color = col_green,
    linewidth = 2
  ) +
  annotate("text",
    x = 0.4, 
    y = text_y_pos,
    label = "Менш схильні\nповертатись", 
    lineheight = 0.9, 
    hjust = 0.5, vjust = 1, 
    size = 5, 
    color = "gray30"
  ) +
  # --- RIGHT SIDE ---
  annotate("segment",
    x = 1.25, xend = 5, 
    y = arrow_y_pos, yend = arrow_y_pos,
    arrow = arrow(length = unit(0.5, "cm")),
    color = col_blue,
    linewidth = 2
  ) +
  annotate("text",
    x = 2.5, 
    y = text_y_pos,
    label = "Більш схильні\nповертатись", 
    lineheight = 0.9, 
    hjust = 0.5, vjust = 1, 
    size = 5,
    color = "gray30"
  )

# Add margin for the footer arrows
p_final <- p_final + theme(plot.margin = margin(10, 10, 80, 10))

ggsave("forest_plot_corrected.png", p_final, width = 10, height = 12, dpi = 300, bg = "white")

print(p_final)
