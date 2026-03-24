library(haven)
library(dplyr)
library(tidyr)
library(writexl)

# 1. Read data
input_file <- "901849.sav"
data <- read_sav(input_file, encoding = "cp1251")

data$K2_1
data$K2_3

data$Weight_2

# Calculate sum of weights where K2_1 is 1 OR K2_3 is 1 (inclusive OR)
weighted_sum <- sum(data$Weight_2[data$K2_1 == 1 | data$K2_3 == 1], na.rm = TRUE)

print(weighted_sum)

weighted_sum <- sum(data$Weight_2[data$K3_1 == 1 | data$K3_2 == 1 | data$K3_3 == 1], na.rm = TRUE)

weighted_sum <- sum(data$Weight_2[data$K3_4 == 1 | data$K3_5 == 1], na.rm = TRUE)

