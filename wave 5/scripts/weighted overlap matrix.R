> library(tidyverse)
> library (writexl)

#select columns and weights
> m_data <- data %>% select(starts_with("A2.1"))
> weights <- data$Weight_2
> total_weight <- sum(weights, na.rm = TRUE) 

#calculate weighted percentage matrix
> m <- as.matrix(m_data)
> m[is.na(m)] <- 0 #for missing values
> W <- diag(weights)
> weighted_matrix <- (t(m) %*% W %*% m / total_weight) * 100

#add survey labels
> labels <- sapply(m_data, function(x) attr(x, "label"))
> colnames(weighted_matrix) <- labels
> rownames(weighted_matrix) <- labels

> print(round(weighted_matrix, 1))
