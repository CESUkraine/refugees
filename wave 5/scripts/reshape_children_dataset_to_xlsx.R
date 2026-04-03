# Children-level dataset: one row per child, with parent variables inherited
# Outputs: children_dataset.xlsx

library(haven)
library(dplyr)
library(tidyr)
library(writexl)

input_file <- "901849.sav"
output_file <- "children_dataset.xlsx"

data <- read_sav(input_file, encoding = "cp1251")

# --- Identify per-child variable stems ---
# B2_N = child age, B3_N = child gender, B5_N_J = education (7 flags),
# J11_N = return intentions, J11.1_N = change in return intentions

child_specific_patterns <- c(
    "^B2_\\d+$", "^B3_\\d+$", "^B5_\\d+_\\d+$",
    "^J11_\\d+$", "^J11\\.1_\\d+$"
)

child_specific_cols <- grep(paste(child_specific_patterns, collapse = "|"),
    names(data),
    value = TRUE
)

# Also exclude B1 (num_children) from parent cols — we'll keep it but it's metadata
parent_cols <- setdiff(names(data), c(child_specific_cols, "B1"))

# --- Build long dataset: one row per child ---

# We need to loop through child numbers 1-5 and stack
max_children <- 5

children_list <- list()

for (n in 1:max_children) {
    # Column names for this child
    age_col <- paste0("B2_", n)
    gender_col <- paste0("B3_", n)
    edu_cols <- paste0("B5_", n, "_", 1:7)
    return_col <- paste0("J11_", n)
    change_col <- paste0("J11.1_", n)

    child_cols <- c(age_col, gender_col, edu_cols, return_col, change_col)

    # Check all exist
    missing <- setdiff(child_cols, names(data))
    if (length(missing) > 0) next

    # Filter to respondents who actually have this child (B1 >= n and age not NA)
    subset_df <- data %>%
        filter(!is.na(B1) & B1 >= n & !is.na(.data[[age_col]])) %>%
        select(all_of(c(parent_cols, child_cols, "B1")))

    if (nrow(subset_df) == 0) next

    # Rename child-specific columns to generic names
    rename_map <- setNames(
        c(age_col, gender_col, paste0("B5_", n, "_", 1:7), return_col, change_col),
        c(
            "child_age", "child_gender", paste0("child_edu_", 1:7),
            "child_return_intentions", "child_return_change"
        )
    )

    subset_df <- subset_df %>%
        rename(!!!rename_map) %>%
        mutate(child_number = n)

    children_list[[n]] <- subset_df
}

children_long <- bind_rows(children_list)

# --- Adjusted weight ---
# two_parents = 1 if A2.1_6 == 1 (living with partner)
children_long <- children_long %>%
    mutate(
        two_parents = ifelse(A2.1_6 == 1, 1, 0),
        adjusted_weight = ifelse(two_parents == 1, Weight_2 / 2, Weight_2)
    )

# --- Apply human-readable labels ---

# Function: for a labelled vector, replace numeric codes with their labels
apply_labels <- function(x) {
    labels <- attr(x, "labels")
    if (is.null(labels)) {
        # No labels — return as-is (converted from labelled to plain)
        return(as.character(zap_labels(x)))
    }
    lookup <- setNames(names(labels), as.character(labels))
    raw <- as.character(x)
    has_label <- raw %in% names(lookup)
    raw[has_label] <- lookup[raw[has_label]]
    raw[is.na(x)] <- NA_character_
    return(raw)
}

# We need to apply labels BEFORE the rename happened for child-specific cols,
# but we already renamed them. The labels are still attached to the vectors
# since we just renamed, not transformed. Let's check and apply.

# Apply to all columns
children_labeled <- children_long %>%
    mutate(across(everything(), apply_labels))

# --- Build human-readable column headers (varname + label) ---

# For parent columns, get label from original data
build_header <- function(col_name, original_data, rename_map_inv = NULL) {
    # Check if this is a renamed child column
    if (!is.null(rename_map_inv) && col_name %in% names(rename_map_inv)) {
        orig_name <- rename_map_inv[[col_name]]
        # Use the first child's version (N=1) for the label text
        lbl <- attr(original_data[[orig_name]], "label")
        if (!is.null(lbl) && lbl != "") {
            return(paste(col_name, lbl))
        }
        return(col_name)
    }

    if (col_name %in% names(original_data)) {
        lbl <- attr(original_data[[col_name]], "label")
        if (!is.null(lbl) && lbl != "") {
            return(paste(col_name, lbl))
        }
    }
    return(col_name)
}

# Inverse rename map: new_name -> original_name (using child 1 as reference for labels)
rename_inv <- list(
    "child_age" = "B2_1",
    "child_gender" = "B3_1",
    "child_edu_1" = "B5_1_1",
    "child_edu_2" = "B5_1_2",
    "child_edu_3" = "B5_1_3",
    "child_edu_4" = "B5_1_4",
    "child_edu_5" = "B5_1_5",
    "child_edu_6" = "B5_1_6",
    "child_edu_7" = "B5_1_7",
    "child_return_intentions" = "J11_1",
    "child_return_change" = "J11.1_1"
)

new_headers <- sapply(names(children_labeled), function(cn) {
    build_header(cn, data, rename_inv)
})

colnames(children_labeled) <- new_headers

# --- Reorder: child-specific cols first, then child_number, adjusted_weight, then parent cols ---

child_col_names <- grep("^(child_|adjusted_weight|two_parents|child_number)",
    names(children_labeled),
    value = TRUE
)

# Also put B1 near the child info
b1_col <- grep("^B1 ", names(children_labeled), value = TRUE)
if (length(b1_col) == 0) b1_col <- "B1"

parent_col_names <- setdiff(names(children_labeled), c(child_col_names, b1_col))

# Final column order
final_order <- c(child_col_names, b1_col, parent_col_names)
children_labeled <- children_labeled[, final_order]

# --- Save ---
write_xlsx(children_labeled, output_file)

cat("Done! Output:", output_file, "\n")
cat("Total child rows:", nrow(children_labeled), "\n")