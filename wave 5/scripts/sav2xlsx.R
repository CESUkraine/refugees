library(haven)
library(dplyr)
library(writexl)

# 1. Load data
input_file <- "901849.sav"
data <- read_sav(input_file, encoding = "cp1251")

# 2. Function to transform values based on your specific rules
# - Use label if it exists
# - Use raw value if no label exists
# - Write "NA" if missing
custom_labeler <- function(x) {
    # Get the label mapping (e.g., c("Так" = 1, "Ні" = 2))
    labels <- attr(x, "labels")

    # Convert current values to character for the final output
    raw_values <- as.character(x)

    if (!is.null(labels)) {
        # Create a lookup table (names are values, values are the text labels)
        lookup <- setNames(names(labels), as.character(labels))

        # Identify which values have a corresponding label
        has_label <- raw_values %in% names(lookup)

        # Replace only those that have a match
        raw_values[has_label] <- lookup[raw_values[has_label]]
    }

    # Replace actual NAs with the string "NA"
    raw_values[is.na(raw_values)] <- "NA"

    return(raw_values)
}

# 3. Create new Headers (Variable Name + Label)
new_headers <- sapply(names(data), function(name) {
    var_label <- attr(data[[name]], "label")
    if (!is.null(var_label) && var_label != "") {
        return(paste(name, var_label)) # Result: "S5 S5. В якій країні..."
    } else {
        return(name)
    }
})

# 4. Apply the value transformation to all columns
# We use across(everything()) to process every column through our function
data_final <- data %>%
    mutate(across(everything(), custom_labeler))

# 5. Apply the new headers
colnames(data_final) <- new_headers

# 6. Export to Excel
output_file <- "901849.xlsx"
write_xlsx(data_final, output_file)

cat("File successfully created:", output_file)
