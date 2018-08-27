library(dplyr)
library(datapkg)
library(tidyr)

##################################################################
#
# Processing Script for Heroin Use
# Created by Jenna Daly
# On 08/23/2018
#
##################################################################

#Setup environment
sub_folders <- list.files()
raw_location <- grep("raw", sub_folders, value=T)
path_to_raw <- (paste0(getwd(), "/", raw_location))
raw_data <- dir(path_to_raw, recursive=T, pattern = "xlsx") 

#read in entire xls file (all sheets)
read_excel_allsheets <- function(filename) {
  sheets <- readxl::excel_sheets(filename)
  x <- lapply(sheets, function(X) readxl::read_excel(filename, sheet = X))
  names(x) <- sheets
  x
}

#Get all sheets for all data sets
for (i in 1:length(raw_data)) {
  mysheets <- read_excel_allsheets(paste0(path_to_raw, "/", raw_data[i]))
  for (j in 1:length(mysheets)) {
    sheet_name <- names(mysheets[j])
    sheet_file <- mysheets[[j]]
    assign(sheet_name, sheet_file)
  }
}

#Find all "Tables"
dfs <- ls()[sapply(mget(ls(), .GlobalEnv), is.data.frame)]
table_dfs <- grep("Table", dfs, value=T)
table_dfs <- table_dfs[grepl(paste(c(" 5"), collapse = "|"), table_dfs)]

all_data <- data.frame(stringsAsFactors = F)
for (i in 1:length(table_dfs)) {
  current_table <- get(table_dfs[i])
  description <- colnames(current_table)[1]
  current_table$`Table Description` <- description
  colnames(current_table) <- c("Order", "Table Number", "Substate Region", "12-17 (Estimate)", 
                               "12-17 MOE Lower", "12-17 MOE Upper", "18-25 (Estimate)", 
                               "18-25 MOE Lower" , "18-25 MOE Upper", "26+ (Estimate)", 
                               "26+ MOE Lower", "26+ MOE Upper", "18+ (Estimate)", 
                               "18+ MOE Lower", "18+ MOE Upper", "Table Description")
  get_year <- as.numeric(substr((unlist(gsub("[^0-9]", "", unlist(table_dfs[i])), "")), 1, 4))
  get_year <- paste0(get_year, "-", get_year+2)
  current_table$Year <- get_year
  all_data <- rbind(all_data, current_table)
}

all_data$Order <- NULL

all_data$`12-17 (Estimate)` <- as.numeric(all_data$`12-17 (Estimate)`) * 100
all_data$`12-17 MOE Lower` <- as.numeric(all_data$`12-17 MOE Lower`) * 100   
all_data$`12-17 MOE Upper` <- as.numeric(all_data$`12-17 MOE Upper`) * 100   
all_data$`18-25 (Estimate)` <- as.numeric(all_data$`18-25 (Estimate)`) * 100 
all_data$`18-25 MOE Lower` <- as.numeric(all_data$`18-25 MOE Lower`) * 100   
all_data$`18-25 MOE Upper` <- as.numeric(all_data$`18-25 MOE Upper`) * 100   
all_data$`26+ (Estimate)` <- as.numeric(all_data$`26+ (Estimate)`) * 100   
all_data$`26+ MOE Lower` <- as.numeric(all_data$`26+ MOE Lower`) * 100     
all_data$`26+ MOE Upper` <- as.numeric(all_data$`26+ MOE Upper`) * 100     
all_data$`18+ (Estimate)` <- as.numeric(all_data$`18+ (Estimate)`) * 100   
all_data$`18+ MOE Lower` <- as.numeric(all_data$`18+ MOE Lower`) * 100     
all_data$`18+ MOE Upper` <- as.numeric(all_data$`18+ MOE Upper`) * 100

#combine all years
#all_data <- rbind(all_legacy_data_sep, all_data)

#select connecticut rows
ct_rows <- which(grepl("Connecticut", all_data$`Substate Region`))
add1 <- ct_rows + 1
add2 <- ct_rows + 2
add3 <- ct_rows + 3
add4 <- ct_rows + 4
add5 <- ct_rows + 5

all_ct_rows <- c(ct_rows, add1, add2, add3, add4, add5)

#add rows for Northeast and United States
us_rows <- which(grepl("^Total United States$", all_data$`Substate Region`))
add1 <- us_rows + 1 #northeast
us_ne_rows <- c(us_rows, add1)

all_rows <- c(all_ct_rows, us_ne_rows)

test <- all_data[all_rows,]

#append "Region" to CT substate regions
regions <- c("Eastern", "North Central", "Northwestern", "South Central", "Southwest")
region_rows <- which(grepl(paste(regions, collapse = "|"), test$`Substate Region`))
test$`Substate Region`[region_rows] <- gsub("$", " Region", test$`Substate Region`[region_rows])
#rename US rows
test$`Substate Region` <- gsub("Total ", "", test$`Substate Region`)

test <- test %>% 
  mutate(`12-17 MOE` = (((as.numeric(`12-17 MOE Upper`)-as.numeric(`12-17 MOE Lower`))/2)/1.96)*1.645, 
         `18-25 MOE` = (((as.numeric(`18-25 MOE Upper`)-as.numeric(`18-25 MOE Lower`))/2)/1.96)*1.645, 
         `26+ MOE` = (((as.numeric(`26+ MOE Upper`)-as.numeric(`26+ MOE Lower`))/2)/1.96)*1.645, 
         `18+ MOE` = (((as.numeric(`18+ MOE Upper`)-as.numeric(`18+ MOE Lower`))/2)/1.96)*1.645) 
        
test <- test %>% 
  select(`Substate Region`, Year, 
         `12-17 (Estimate)`, `12-17 MOE`, 
         `18-25 (Estimate)`, `18-25 MOE`, 
         `26+ (Estimate)`, `26+ MOE`, 
         `18+ (Estimate)`, `18+ MOE`, 
         `Table Description`)

#recode description column
test$`Heroin Use` <- NA
test$`Heroin Use`[grep("Heroin", test$`Table Description`)] <- "Heroin Use in the Past Year"

#wide to long format
all_data_long <- gather(test, Variable, Value, 3:10, factor_key=F)

#assign age range
all_data_long$`Age Range` <- NA
all_data_long$`Age Range`[grep("12-17", all_data_long$Variable)] <- "12-17"
all_data_long$`Age Range`[grep("26+", all_data_long$Variable)] <- "Over 25"
all_data_long$`Age Range`[grep("18+", all_data_long$Variable)] <- "Over 17"
all_data_long$`Age Range`[grep("18-25", all_data_long$Variable)] <- "18-25"

#assign variable
all_data_long$Variable[grep("Estimate", all_data_long$Variable)] <- "Heroin Use"
all_data_long$Variable[grep("MOE", all_data_long$Variable)] <- "Margins of Error"

#assign measure type
all_data_long$`Measure Type` <- "Percent"

#Select and sort columns
all_data_final <- all_data_long %>% 
  select(`Substate Region`, Year, `Age Range`, `Heroin Use`, `Measure Type`, Variable, Value) %>% 
  rename(Region = `Substate Region`) %>% 
  arrange(Region, Year, `Age Range`, `Heroin Use`, Variable)

all_data_final$Value <- round(as.numeric(all_data_final$Value), 2)

# Write to File
write.table(
  all_data_final,
  file.path(getwd(), "data", "heroin-use-2016.csv"),
  sep = ",",
  na = "-9999",
  row.names = F
)

