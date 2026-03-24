library(tidyverse)
library(scales)

# --- GLOBAL SETTINGS ---
# 1. Define the universal font size (in points)
# Note: ggplot uses different units for text geoms (mm) vs theme elements (pt).
# This variable sets the base point size for the theme. 
# The geom_text size is calculated as FONT_SIZE / .pt to match visually.
FONT_SIZE <- 20

# 2. Fix for "Font Not Found" Error
windowsFonts(AptosDisp = windowsFont("Aptos Display"))

# 3. Define Brand Palette
brand_palette <- c(
  "#00509B", # A1 (Blue)
  "#F0593F", # D1 (Orange)
  "#31AE87", # E1 (Teal)
  "#AF2D5A", # F1 (Berry)
  "#73B932", # B1 (Green)
  "#FAD51E", # C1 (Yellow)
  "#80A7CD"  # A2 (Soft Blue)
)

# --- DATA PREPARATION ---
df <- data.frame(
  Income_Source = c(
    "Зарплата чи підприємницький дохід з України",
    "Зарплата чи підприємницький дохід НЕ з України",
    "Пенсія, стипендія, дитячі та інші соціальні виплати з України",
    "Пенсія, стипендія, дитячі та інші соціальні виплати НЕ з України",
    "Заощадження",
    "Інші грошові перекази з України (від родичів, від здачі житла тощо)",
    "Інші грошові перекази НЕ з України (від родичів, від здачі житла тощо)"
  ),
  Pop_Weight = c(118.18, 589.91, 159.47, 271.93, 318.77, 76.48, 71.83),
  Budget_Share = c(59.36, 83.75, 41.50, 60.87, 44.24, 35.59, 49.81)
)

# Sorting
df <- df %>%
  arrange(desc(Budget_Share)) %>%
  mutate(Income_Source = factor(Income_Source, levels = Income_Source))

# Coordinates
df_plot <- df %>%
  mutate(
    width = Pop_Weight / sum(Pop_Weight),
    xmax = cumsum(width),
    xmin = xmax - width,
    x_mid = (xmin + xmax) / 2,
    height = Budget_Share / 100
  )

# --- PLOTTING ---
p <- ggplot(df_plot) +
  # 1. Draw Bars
  geom_rect(aes(xmin = xmin, xmax = xmax, ymin = 0, ymax = height, fill = Income_Source), 
            color = "white", linewidth = 0.5) +
  
  # 2. Draw Percentage Labels
  # Size conversion: ggplot 'size' is in mm, theme 'size' is in pts.
  # We divide by .pt (approx 2.8) to match the theme size perfectly.
  geom_text(aes(x = x_mid, y = height, label = paste0(round(Budget_Share, 0), "%")), 
            vjust = -0.5, 
            size = FONT_SIZE / .pt, 
            fontface = "bold", 
            color = "#333333",
            family = "AptosDisp") +
  
  # 3. Axes Config
  scale_y_continuous(labels = percent_format(), limits = c(0, 1.05), expand = c(0,0)) +
  scale_x_continuous(labels = percent_format(), expand = c(0,0)) +
  scale_fill_manual(values = brand_palette) + 
  
  # 4. Remove Titles
  labs(title = NULL, subtitle = NULL, x = NULL, y = NULL, fill = NULL, caption = NULL) +
  
  # 5. Theme & Styling
  theme_minimal(base_family = "AptosDisp") +
  theme(
    # Remove grids
    panel.grid = element_blank(),
    
    # Axis Lines & Ticks
    axis.line = element_line(color = "#333333", linewidth = 0.8),
    axis.ticks = element_line(color = "#333333", linewidth = 0.8),
    axis.ticks.length = unit(0.3, "cm"),
    
    # Axis Text (Numbers) - Set to GLOBAL FONT SIZE
    axis.text = element_text(color = "#333333", size = FONT_SIZE),
    
    # Remove Titles
    axis.title = element_blank(),
    
    # Layout
    legend.position = "none",
    plot.margin = margin(40, 40, 40, 40)
  )

# 6. Save
ggsave("income_mekko_brand_v4.png", plot = p, width = 12, height = 9, dpi = 300)

