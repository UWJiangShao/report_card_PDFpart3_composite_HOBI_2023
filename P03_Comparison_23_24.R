library(dplyr)
library(openxlsx)


# data 24
SC24 <- read.xlsx("C:\\Users\\jiang.shao\\Dropbox (UFL)\\MCO Report Card - 2024\\Program\\5. Composite\\Output\\Comparison 23_24\\data24.xlsx", sheet = 1)
SA24 <- read.xlsx("C:\\Users\\jiang.shao\\Dropbox (UFL)\\MCO Report Card - 2024\\Program\\5. Composite\\Output\\Comparison 23_24\\data24.xlsx", sheet = 2)
SP24 <- read.xlsx("C:\\Users\\jiang.shao\\Dropbox (UFL)\\MCO Report Card - 2024\\Program\\5. Composite\\Output\\Comparison 23_24\\data24.xlsx", sheet = 3)
SK24 <- read.xlsx("C:\\Users\\jiang.shao\\Dropbox (UFL)\\MCO Report Card - 2024\\Program\\5. Composite\\Output\\Comparison 23_24\\data24.xlsx", sheet = 4)


# data 23
SC23 <- read.xlsx("C:\\Users\\jiang.shao\\Dropbox (UFL)\\MCO Report Card - 2024\\Program\\5. Composite\\Output\\Comparison 23_24\\data23.xlsx", sheet = 1)
SA23 <- read.xlsx("C:\\Users\\jiang.shao\\Dropbox (UFL)\\MCO Report Card - 2024\\Program\\5. Composite\\Output\\Comparison 23_24\\data23.xlsx", sheet = 2)
SP23 <- read.xlsx("C:\\Users\\jiang.shao\\Dropbox (UFL)\\MCO Report Card - 2024\\Program\\5. Composite\\Output\\Comparison 23_24\\data23.xlsx", sheet = 3)
SK23 <- read.xlsx("C:\\Users\\jiang.shao\\Dropbox (UFL)\\MCO Report Card - 2024\\Program\\5. Composite\\Output\\Comparison 23_24\\data23.xlsx", sheet = 4)


# Process STAR
# ST23 <- ST23 %>% select(-SAper10kmm_clust)
# names(ST23) <- c('plancode', 'count', 'mm', 'count_pct', 'mm_pct', 'ae_ratio', 'score', 'center', 'rating')
# 
# ST24 <- ST24 %>% select(-MCONAME, -SERVICEAREA, -STper10kmm_clust)
# names(ST24) <- c('plancode', 'count', 'mm', 'count_pct', 'mm_pct', 'ae_ratio', 'score', 'center', 'rating')


merge_and_calculate_diff <- function(DF23, DF24) {
  # Load dplyr for data manipulation
  
  # Rename the columns except for Plan.Code
  # Remove MCO and Service.Area columns
  DF23 <- DF23 %>% select(-MCO, -Service.Area)
  DF24 <- DF24 %>% select(-MCO, -Service.Area)
  
  # Replace "---" and "-- " with NA
  DF23 <- DF23 %>% mutate(across(everything(), ~ifelse(. %in% c("---", "-- "), NA, .)))
  DF24 <- DF24 %>% mutate(across(everything(), ~ifelse(. %in% c("---", "-- "), NA, .)))
  
  # Convert the first five columns to numeric
  DF23[, 1:5] <- lapply(DF23[, 1:5], as.numeric)
  DF24[, 1:5] <- lapply(DF24[, 1:5], as.numeric)
  
  DF23$Plan.Code <- trimws(DF23$Plan.Code)
  DF24$Plan.Code <- trimws(DF24$Plan.Code)
  
  DF23_renamed <- DF23 %>% rename_with(~paste0(., "_23"), -Plan.Code)
  DF24_renamed <- DF24 %>% rename_with(~paste0(., "_24"), -Plan.Code)
  
  # Merge the datasets by Plan.Code with full join
  merged_df <- full_join(DF23_renamed, DF24_renamed, by = "Plan.Code")
  
  # Calculate the differences and create new columns
  vars_to_diff <- c("exp", "care", "prev", "chronic", "all")
  for (var in vars_to_diff) {
    merged_df[[paste0(var, "_diff")]] <- merged_df[[paste0(var, "_24")]] - merged_df[[paste0(var, "_23")]]
  }
  
  return(merged_df)
}

# Usage example
# result_df <- merge_and_calculate_diff(DF23, DF24)




SC_compare <- merge_and_calculate_diff(SC23, SC24)
SA_compare <- merge_and_calculate_diff(SA23, SA24)
SP_compare <- merge_and_calculate_diff(SP23, SP24)
SK_compare <- merge_and_calculate_diff(SK23, SK24)











# ########################
write_dataframes_to_excel <- function(df_list, file_name) {
  wb <- createWorkbook()

  for (df_name in names(df_list)) {
    addWorksheet(wb, df_name)
    writeData(wb, sheet = df_name, x = df_list[[df_name]])
  }

  saveWorkbook(wb, file = file_name, overwrite = TRUE)
}

dataframes <- list(
  SC = SC_compare,
  SA = SA_compare,
  SP = SP_compare,
  SK = SK_compare
)


write_dataframes_to_excel(dataframes, "C:\\Users\\jiang.shao\\Dropbox (UFL)\\MCO Report Card - 2024\\Program\\5. Composite\\Output\\Comparison 23_24\\Composite_comparison_final.xlsx")

