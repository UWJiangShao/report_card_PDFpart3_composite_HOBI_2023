library(readxl)
library(sas7bdat)
library(httr)

# read in the SAS dataset that we use to create table by SDA
data <- read.sas7bdat("C://Users//jiang.shao//Dropbox (UFL)//MCO Report Card - 2024//Program//5. Composite//Data//temp_data//contact_mco_info_plancode.sas7bdat")
 
website_eng <- data$web_eng
website_spa <- data$web_spa

# dedup to save time for running 
website_eng <- unique(website_eng)
website_spa <- unique(website_spa)


results_df_eng <- data.frame(Website = character(), Result = character(), stringsAsFactors = FALSE)

for (url in website_eng) {
  response <- tryCatch({
    GET(url, timeout(10))  # 10-second timeout
  }, error = function(e) {
    NULL  # In case of error, return NULL
  })
  
  if (!is.null(response)) {
    status <- status_code(response)
    if (status == 200) {
      result <- "Accessible"
    } else {
      result <- paste("Not accessible, status code:", status)
    }
  } else {
    result <- "Could not be reached or timed out"
  }
  
  results_df_eng <- rbind(results_df_eng, data.frame(Website = url, Result = result))
  
  Sys.sleep(1) 
}


results_df_spa <- data.frame(Website = character(), Result = character(), stringsAsFactors = FALSE)


for (url in website_spa) {
  response <- tryCatch({
    GET(url, timeout(10))  # 10-second timeout
  }, error = function(e) {
    NULL  # In case of error, return NULL
  })
  
  if (!is.null(response)) {
    status <- status_code(response)
    if (status == 200) {
      result <- "Accessible"
    } else {
      result <- paste("Not accessible, status code:", status)
    }
  } else {
    result <- "Could not be reached or timed out"
  }
  
  results_df_spa <- rbind(results_df_eng, data.frame(Website = url, Result = result))
  
  Sys.sleep(1) 
}


library(writexl)
write_xlsx(results_df_eng, "C://Users//jiang.shao//Dropbox (UFL)//MCO Report Card - 2024//Program//5. Composite//Output//website return value check//website_check_results_eng.xlsx")
write_xlsx(results_df_spa, "C://Users//jiang.shao//Dropbox (UFL)//MCO Report Card - 2024//Program//5. Composite//Output//website return value check//website_check_results_spa.xlsx")


