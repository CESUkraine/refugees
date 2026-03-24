> data <- read_sav("901849.sav", encoding = "cp1251")
> head(data)
#looking at labels in the data
> lapply(data[grep("A2.1", names(data))], attr, "label")

#to view table with answers
>library(tidyverse)
> data %>%
+   select(starts_with("A2.1")) %>%
+   # This gathers the labels so you don't just see "A2.1_1"
+   map_df(~{
+     label <- attr(.x, "label")
+     count <- sum(.x == 1, na.rm = TRUE)
+     prop  <- mean(.x == 1, na.rm = TRUE) * 100
+     tibble(Option = label, N = count, Percent = prop)
+   })

#define the survey design
design <- svydesign(ids = ~ID, weights = ~Weight_2, data = data)

#get weighted pct
> cols <- c("A2.1_1", "A2.1_2", "A2.1_3", "A2.1_4", "A2.1_5", "A2.1_6", "A2.1_7", "A2.1_8")
> multiple_res_summary <- map_df(cols, ~{
+   # Create a formula for each column
+   form <- as.formula(paste0("~", .x))
+   
+   # Calculate weighted mean (which is the percentage for 0/1 data)
+   stat <- svymean(form, design, na.rm = TRUE)
+   
+   tibble(
+     Option = attr(data[[.x]], "label"), # Pulls the text label
+     Weighted_Pct = as.numeric(stat) * 100
+   )
+ })
> print(multiple_res_summary)