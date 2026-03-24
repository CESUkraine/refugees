# Install the package if you haven't already
# install.packages("haven")

library(haven)
library(survey)

# Read the SAV file
data <- read_sav("901849.sav", encoding = "cp1251")

# View the first few rows
head(data)
# Get the question behind the variable S5
attr(data$S5, "label")
# Get all the attributes of the variable S5
attributes(data$S5)

# Define the survey design
design <- svydesign(ids = ~ID, weights = ~Weight, data = data)

# Weighted Linear Regression
weighted_model <- svyglm(as.numeric(Z4) ~ as.factor(Рекрутинг), design = design)

model <- lm(S5 ~ Рекрутинг + S1, data = data)

# Create a cross-tabulation table
table_data <- table(data$Рекрутинг, as.factor(data$J2))

# Run the test
chisq_test <- chisq.test(table_data)

print(chisq_test)
