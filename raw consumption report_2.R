library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)
library(skimr)
library(janitor)
library(lubridate)

#########################################################################################################################
######################################################## Data Input #####################################################
#########################################################################################################################

######################################################### DSX List ######################################################

dsx_lag1 <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/BI Forecast Backup/2023/DSX Forecast Backup - 2023.04.03.xlsx")

dsx_1 <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/BI Forecast Backup/2023/DSX Forecast Backup - 2023.05.03.xlsx")
dsx_2 <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/BI Forecast Backup/2023/DSX Forecast Backup - 2023.06.01.xlsx")
dsx_3 <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/BI Forecast Backup/2023/DSX Forecast Backup - 2023.07.05.xlsx")
dsx_4 <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/BI Forecast Backup/2023/DSX Forecast Backup - 2023.08.01.xlsx")
dsx_5 <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/BI Forecast Backup/2023/DSX Forecast Backup - 2023.09.01.xlsx")
dsx_6 <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/BI Forecast Backup/2023/DSX Forecast Backup - 2023.10.02.xlsx")
dsx_7 <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/BI Forecast Backup/2023/DSX Forecast Backup - 2023.11.01.xlsx")
dsx_8 <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/BI Forecast Backup/2023/DSX Forecast Backup - 2023.12.04.xlsx")
dsx_9 <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/BI Forecast Backup/2024/DSX Forecast Backup - 2024.01.02.xlsx")
dsx_10 <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/BI Forecast Backup/2024/DSX Forecast Backup - 2024.02.01.xlsx")
dsx_11 <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/BI Forecast Backup/2024/DSX Forecast Backup - 2024.03.01.xlsx")
dsx_12 <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/BI Forecast Backup/2024/DSX Forecast Backup - 2024.04.01.xlsx")

######################################################### Other List ######################################################

# BoM RM to sku 
rm_to_sku <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Weekly Report run/2024/04.30.2024/Raw Material Inventory Health (IQR) NEW TEMPLATE - 04.30.2024.xlsx", 
                        sheet = "RM to SKU")

# BoM Report 
bom <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/RM/Weekly Report run/2024/04.30.2024/Raw Material Inventory Health (IQR) NEW TEMPLATE - 04.30.2024.xlsx", 
                  sheet = "BoM")


## sku_actual (Make sure in the MSTR if months info input correct) 
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/7D421DDA4D4411DA73B4469771826BD9/W62--K46
sku_actual <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Oil Consumption/12 month rolling report/2024.04 Lag4/shipped.xlsx")

# Input sales orders (Make sure in the MSTR if months info input correct) 
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/7D421DDA4D4411DA73B4469771826BD9/W62--K46
sales_orders <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Oil Consumption/12 month rolling report/2024.04 Lag4/ordered.xlsx")


# Completed SKU List 
completed_sku_list <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/04.30.2024/Completed SKU list - Linda.xlsx")

###########################################################################################################################

date_1_month_ago <- format(ymd(today()) %m-% months(1), "%Y%m")
date_2_months_ago <- format(ymd(today()) %m-% months(2), "%Y%m")
date_3_months_ago <- format(ymd(today()) %m-% months(3), "%Y%m")
date_4_months_ago <- format(ymd(today()) %m-% months(4), "%Y%m")
date_5_months_ago <- format(ymd(today()) %m-% months(5), "%Y%m")
date_6_months_ago <- format(ymd(today()) %m-% months(6), "%Y%m")
date_7_months_ago <- format(ymd(today()) %m-% months(7), "%Y%m")
date_8_months_ago <- format(ymd(today()) %m-% months(8), "%Y%m")
date_9_months_ago <- format(ymd(today()) %m-% months(9), "%Y%m")
date_10_months_ago <- format(ymd(today()) %m-% months(10), "%Y%m")
date_11_months_ago <- format(ymd(today()) %m-% months(11), "%Y%m")
date_12_months_ago <- format(ymd(today()) %m-% months(12), "%Y%m")


# functions
select_forecast_columns <- function(dataframe) {
  common_columns <- c("mfg_ref", "mfg_loc", "sku", "sku_description", "label", "category", "platform", "group_no", "group", "adjusted_forecast_cases", "adjusted_forecast_pounds_lbs")
  
  forecast_column <- if (any(names(dataframe) == "forecast_month_year_code")) {
    "forecast_month_year_code"
  } else if (any(names(dataframe) == "forecast_month_year_id")) {
    "forecast_month_year_id"
  } else {
    stop("Neither forecast_month_year_code nor forecast_month_year_id found in the dataframe")
  }
  selected_data <- dplyr::select(dataframe, all_of(c(forecast_column, common_columns)))
  return(selected_data)
}



process_forecast_data <- function(dataframe) {
  # Determine the appropriate column to use
  forecast_column <- if (any(names(dataframe) == "forecast_month_year_code")) {
    "forecast_month_year_code"
  } else if (any(names(dataframe) == "forecast_month_year_id")) {
    "forecast_month_year_id"
  } else {
    stop("Neither forecast_month_year_code nor forecast_month_year_id found in the dataframe")
  }
  
  dataframe %>%
    dplyr::group_by(!!sym(forecast_column), mfg_ref, mfg_loc, sku, sku_description, label, category, platform, group_no, group) %>%
    dplyr::summarise(adjusted_forecast_cases = sum(adjusted_forecast_cases),
                     adjusted_forecast_pounds_lbs = sum(adjusted_forecast_pounds_lbs)) %>%
    dplyr::mutate(year = stringr::str_sub(!!sym(forecast_column), 1, 4)) -> processed_df
  
  forecast_month_year <- processed_df[[forecast_column]]
  month <- substr(forecast_month_year, nchar(forecast_month_year)-1, nchar(forecast_month_year)) %>%
    data.frame() %>%
    cbind(processed_df) %>%
    dplyr::rename(month = ".") %>%
    dplyr::select(-!!sym(forecast_column)) %>%
    dplyr::relocate(year, month) %>% 
    dplyr::filter(mfg_loc != 22) %>% 
    dplyr::filter(mfg_loc != 16) 
  
  return(month)
}




# Forecast dsx (for lag1, use the first day of the month) ----
# Make sure to put the date correctly few below ----

# DSX - 1
dsx_1[-1, ] -> dsx_1
colnames(dsx_1) <- dsx_1[1, ]
dsx_1[-1, ] -> dsx_1

dsx_1 %>%
  janitor::clean_names() %>%
  readr::type_convert() %>%
  data.frame() %>%
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_12_months_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_12_months_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>%    
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_1


process_forecast_data(forecast_1) -> forecast_1


# DSX - 2
dsx_2[-1, ] -> dsx_2
colnames(dsx_2) <- dsx_2[1, ]
dsx_2[-1, ] -> dsx_2

dsx_2 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_11_months_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_11_months_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>%   
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_2


process_forecast_data(forecast_2) -> forecast_2


# DSX - 3
dsx_3[-1, ] -> dsx_3
colnames(dsx_3) <- dsx_3[1, ]
dsx_3[-1, ] -> dsx_3

dsx_3 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_10_months_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_10_months_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>%    
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_3


process_forecast_data(forecast_3) -> forecast_3


# DSX - 4
dsx_4[-1, ] -> dsx_4
colnames(dsx_4) <- dsx_4[1, ]
dsx_4[-1, ] -> dsx_4

dsx_4 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_9_months_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_9_months_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>%   
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_4 


process_forecast_data(forecast_4) -> forecast_4


# DSX - 5
dsx_5[-1, ] -> dsx_5
colnames(dsx_5) <- dsx_5[1, ]
dsx_5[-1, ] -> dsx_5

dsx_5 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_8_months_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_8_months_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>%   
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_5 


process_forecast_data(forecast_5) -> forecast_5




# DSX - 6
dsx_6[-1, ] -> dsx_6
colnames(dsx_6) <- dsx_6[1, ]
dsx_6[-1, ] -> dsx_6

dsx_6 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_7_months_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_7_months_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>% 
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_6 


process_forecast_data(forecast_6) -> forecast_6




# DSX - 7
dsx_7[-1, ] -> dsx_7
colnames(dsx_7) <- dsx_7[1, ]
dsx_7[-1, ] -> dsx_7

dsx_7 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_6_months_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_6_months_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>% 
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_7


process_forecast_data(forecast_7) -> forecast_7




# DSX - 8
dsx_8[-1, ] -> dsx_8
colnames(dsx_8) <- dsx_8[1, ]
dsx_8[-1, ] -> dsx_8

dsx_8 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_5_months_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_5_months_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>% 
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_8


process_forecast_data(forecast_8) -> forecast_8




# DSX - 9
dsx_9[-1, ] -> dsx_9
colnames(dsx_9) <- dsx_9[1, ]
dsx_9[-1, ] -> dsx_9

dsx_9 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_4_months_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_4_months_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>% 
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_9


process_forecast_data(forecast_9) -> forecast_9



# DSX - 10
dsx_10[-1, ] -> dsx_10
colnames(dsx_10) <- dsx_10[1, ]
dsx_10[-1, ] -> dsx_10

dsx_10 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_3_months_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_3_months_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>% 
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_10


process_forecast_data(forecast_10) -> forecast_10



# DSX - 11
dsx_11[-1, ] -> dsx_11
colnames(dsx_11) <- dsx_11[1, ]
dsx_11[-1, ] -> dsx_11

dsx_11 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_2_months_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_2_months_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>% 
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_11


process_forecast_data(forecast_11) -> forecast_11




# DSX - 12
dsx_12[-1, ] -> dsx_12
colnames(dsx_12) <- dsx_12[1, ]
dsx_12[-1, ] -> dsx_12

dsx_12 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_1_month_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_1_month_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>% 
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_12


process_forecast_data(forecast_12) -> forecast_12



#################################################################################################################

rbind(forecast_1, forecast_2, forecast_3, forecast_4, forecast_5, forecast_6, forecast_7, forecast_8, forecast_9, forecast_10,
      forecast_11, forecast_12) -> forecast


forecast %>%
  dplyr::mutate(year = as.double(year),
                month = as.double(month),
                date_ref = paste0(year, "_", month, "_", mfg_loc, "_", sku)) %>% 
  dplyr::relocate(date_ref) -> forecast






##################################################################################################################


# BoM RM to sku ----
# rm_to_sku <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/Raw Material Inventory Health (IQR) - 03.29.23.xlsx", 
#                        sheet = "RM to SKU")

rm_to_sku %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  dplyr::select(1:3) %>%
  tidyr::separate(comp_ref, c("location", "component"), sep = "-") %>% 
  dplyr::select(-1) %>% 
  dplyr::rename(sku = parent_item_number) %>%
  dplyr::slice(-n()) %>% 
  dplyr::filter(!is.na(component)) -> rm_to_sku




rm_to_sku %>%
  dplyr::select(component) %>% 
  dplyr::mutate(component = as.double(component)) -> rm_to_sku_comp


rm_to_sku %>% 
  dplyr::mutate(component = as.double(component)) -> rm_to_sku






forecast %>% 
  dplyr::mutate(year = as.double(year),
                month = as.double(month),
                date_ref = paste0(year, "_", month, "_", mfg_loc, "_", sku))-> forecast_2




# BoM Report ----
# bom <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/Raw Material Inventory Health (IQR) - 03.29.23.xlsx", 
#                   sheet = "BoM")

bom %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(ref = gsub("-", "_", ref)) %>% 
  dplyr::select(ref, comp_ref, business_unit, parent_item_number, comp_number_labor_code, quantity_w_scrap) %>% 
  dplyr::rename(mfg_loc = business_unit,
                sku = parent_item_number,
                component = comp_number_labor_code) %>% 
  dplyr::mutate(mfg_ref = paste0(mfg_loc, "_", sku),
                mfg_comp_ref = paste0(mfg_loc, "_", component)) %>% 
  dplyr::select(-mfg_loc) %>% 
  dplyr::relocate(ref, comp_ref, mfg_ref, mfg_comp_ref, sku, component, quantity_w_scrap) %>% 
  dplyr::mutate(quantity_w_scrap = round(quantity_w_scrap, 2)) -> bom





## sku_actual ----
# sku_actual <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Oil Consumption/12 month rolling report/Mar.2023/Order and Shipped History (20).xlsx")


sku_actual %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>%
  dplyr::rename(mfg_loc = product_manufacturing_location,
                sku = product_label_sku,
                actual_shipped_cases = cases,
                actual_shipped_lbs = net_pounds_lbs,
                year = calendar_year,
                month = calendar_month_no) %>% 
  dplyr::select(year, month, sku, mfg_loc, actual_shipped_lbs, actual_shipped_cases) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                date_ref = paste0(year, "_", month, "_", mfg_loc, "_", sku)) %>% 
  dplyr::relocate(date_ref)-> sku_actual

sku_actual %>% 
  dplyr::group_by(date_ref, mfg_ref) %>% 
  dplyr::summarise(actual_shipped_lbs = sum(actual_shipped_lbs),
                   actual_shipped_cases = sum(actual_shipped_cases)) %>% 
  dplyr::mutate(actual_shipped_lbs = replace(actual_shipped_lbs, is.na(actual_shipped_lbs), 0),
                actual_shipped_cases = replace(actual_shipped_cases, is.na(actual_shipped_cases), 0)) -> sku_actual_pivot


# combine with dsx_with_oil x open_order

forecast_2 %>% 
  dplyr::left_join(sku_actual_pivot %>% select(-mfg_ref), by = "date_ref") %>% 
  dplyr::mutate(actual_shipped_cases = replace(actual_shipped_cases, is.na(actual_shipped_cases), 0)) %>% 
  dplyr::mutate(actual_shipped_lbs = replace(actual_shipped_lbs, is.na(actual_shipped_lbs), 0),
                adjusted_forecast_pounds_lbs = round(adjusted_forecast_pounds_lbs, 0)) -> raw_comsumption_comparison




################################################# Sales Orders ##################################################
# Input sales orders ----
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/7D421DDA4D4411DA73B4469771826BD9/W62--K46

# sales_orders <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Oil Consumption/12 month rolling report/Mar.2023/Order and Shipped History (21).xlsx")

sales_orders %>% 
  janitor::clean_names() %>% 
  dplyr::rename(mfg_loc = product_manufacturing_location,
                sku = product_label_sku,
                description = x6,
                order_qty_final = ordered_final_qty,
                order_qty_original = ordered_original_qty,
                year = calendar_year,
                month = calendar_month_no) %>% 
  dplyr::mutate(order_qty_final = replace(order_qty_final, is.na(order_qty_final), 0),
                order_qty_original = replace(order_qty_original, is.na(order_qty_original), 0),
                sku = gsub("-", "", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                date_ref = paste0(year, "_", month, "_", mfg_loc, "_", sku)) %>%
  dplyr::relocate(date_ref) %>% 
  readr::type_convert() -> sales_orders


sales_orders %>% 
  dplyr::group_by(date_ref, mfg_ref) %>% 
  dplyr::summarise(order_qty_final = sum(order_qty_final),
                   order_qty_original = sum(order_qty_original)) %>% 
  dplyr::mutate(order_qty_final = ifelse(order_qty_final < 0, 0, order_qty_final),
                order_qty_original = ifelse(order_qty_original < 0, 0, order_qty_original)) -> sales_orders_pivot



raw_comsumption_comparison %>% 
  dplyr::left_join(sales_orders_pivot %>% select(-mfg_ref), by = "date_ref") -> raw_comsumption_comparison


# NA to 0
raw_comsumption_comparison %>% 
  dplyr::mutate(order_qty_final = replace(order_qty_final, is.na(order_qty_final), 0),
                order_qty_original = replace(order_qty_original, is.na(order_qty_original), 0)) ->  raw_comsumption_comparison



################################################ second phase #######################################
bom %>% 
  dplyr::select(sku, component, quantity_w_scrap) -> bom_2

bom_2[!duplicated(bom_2[,c("sku", "component")]),] -> bom_2

raw_comsumption_comparison %>% 
  dplyr::left_join(bom_2) -> raw_comsumption_comparison_ver2


raw_comsumption_comparison_ver2 %>% 
  dplyr::mutate(forecasted_oil_qty = adjusted_forecast_cases * quantity_w_scrap,
                consumption_qty_actual_shipped = actual_shipped_cases * quantity_w_scrap,
                consumption_percent_adjusted_actual_shipped = consumption_qty_actual_shipped / forecasted_oil_qty) %>%
  
  dplyr::mutate(consumption_qty_sales_order_qty = order_qty_final * quantity_w_scrap,
                consumption_percent_adjusted_sales_order = consumption_qty_sales_order_qty / forecasted_oil_qty) %>% 
  
  
  dplyr::mutate(consumption_percent_adjusted_actual_shipped = replace(consumption_percent_adjusted_actual_shipped, is.na(consumption_percent_adjusted_actual_shipped) | is.nan(consumption_percent_adjusted_actual_shipped) | is.infinite(consumption_percent_adjusted_actual_shipped), 0)) %>% 
  dplyr::mutate(consumption_percent_adjusted_actual_shipped = sprintf("%1.2f%%", 100*consumption_percent_adjusted_actual_shipped)) %>% 
  dplyr::mutate(consumption_percent_adjusted_sales_order = replace(consumption_percent_adjusted_sales_order, is.na(consumption_percent_adjusted_sales_order) | is.nan(consumption_percent_adjusted_sales_order) | is.infinite(consumption_percent_adjusted_sales_order), 0)) %>% 
  dplyr::mutate(consumption_percent_adjusted_sales_order = sprintf("%1.2f%%", 100*consumption_percent_adjusted_sales_order)) %>% 
  
  
  dplyr::mutate(diff_between_forecast_actual =  forecasted_oil_qty - consumption_qty_actual_shipped,
                diff_between_forecast_original = forecasted_oil_qty - consumption_qty_sales_order_qty) %>% 
  
  
  dplyr::arrange(mfg_ref) %>% 
  dplyr::relocate(component, .after = group) -> raw_comsumption_comparison_ver2




raw_comsumption_comparison_ver2 %>% 
  dplyr::filter(mfg_loc != "-1") -> raw_comsumption_comparison_ver2




# Duplicated values delete
raw_comsumption_comparison_ver2[!duplicated(raw_comsumption_comparison_ver2[,c("date_ref", "mfg_ref", "sku", "component", "quantity_w_scrap")]),] -> raw_comsumption_comparison_ver2



#################################################################################################################################################
#################################################################################################################################################

# final round up
raw_comsumption_comparison_ver2 %>% 
  dplyr::mutate(adjusted_forecast_cases = round(adjusted_forecast_cases, 0),
                forecasted_oil_qty = round(forecasted_oil_qty, 0),
                consumption_qty_actual_shipped = round(consumption_qty_actual_shipped, 0),
                diff_between_forecast_actual = round(diff_between_forecast_actual, 0),
                consumption_qty_sales_order_qty = round(consumption_qty_sales_order_qty, 0),
                diff_between_forecast_original = round(diff_between_forecast_original, 0)) -> raw_comsumption_comparison_ver2


# final touch
raw_comsumption_comparison_ver2 %>% 
  dplyr::mutate(mfg_ref = gsub("_", "-", mfg_ref)) -> raw_comsumption_comparison_ver2 


# column rename
raw_comsumption_comparison_final <- raw_comsumption_comparison_ver2

raw_comsumption_comparison_ver2 %>% 
  dplyr::select(year, month, mfg_ref, mfg_loc, sku, sku_description, label, category, platform, group_no, group, component,
                quantity_w_scrap, adjusted_forecast_cases, forecasted_oil_qty, 
                actual_shipped_cases, 
                consumption_qty_actual_shipped, consumption_percent_adjusted_actual_shipped,
                diff_between_forecast_actual, order_qty_final, order_qty_original, consumption_qty_sales_order_qty, 
                consumption_percent_adjusted_sales_order, diff_between_forecast_original) %>% 
  dplyr::arrange(year, month, mfg_loc, sku) -> raw_comsumption_comparison_ver2


#######################################################################################################################
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################
#################################################### Lag 1 - 4 work ###################################################
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################


###### Lag 1 #####
# DSX - Lag_1
dsx_lag1[-1, ] -> dsx_lag1
colnames(dsx_lag1) <- dsx_lag1[1, ]
dsx_lag1[-1, ] -> dsx_lag1

dsx_lag1 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_12_months_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_12_months_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>%   
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_lag_1_1


process_forecast_data(forecast_lag_1_1) -> forecast_lag_1_1



# DSX - 1

dsx_1 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_11_months_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_11_months_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>%   
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_1_1


process_forecast_data(forecast_1_1) -> forecast_1_1


# DSX - 2

dsx_2 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_10_months_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_10_months_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>%   
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_2_1


process_forecast_data(forecast_2_1) -> forecast_2_1


# DSX - 3

dsx_3 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_9_months_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_9_months_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>%   
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_3_1


process_forecast_data(forecast_3_1) -> forecast_3_1


# DSX - 4

dsx_4 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_8_months_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_8_months_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>%   
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_4_1


process_forecast_data(forecast_4_1) -> forecast_4_1


# DSX - 5

dsx_5 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_7_months_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_7_months_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>%   
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_5_1


process_forecast_data(forecast_5_1) -> forecast_5_1




# DSX - 6

dsx_6 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_6_months_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_6_months_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>%   
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_6_1


process_forecast_data(forecast_6_1) -> forecast_6_1




# DSX - 7

dsx_7 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_5_months_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_5_months_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>%   
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_7_1


process_forecast_data(forecast_7_1) -> forecast_7_1




# DSX - 8

dsx_8 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_4_months_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_4_months_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>%   
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_8_1


process_forecast_data(forecast_8_1) -> forecast_8_1




# DSX - 9

dsx_9 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_3_months_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_3_months_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>%   
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_9_1


process_forecast_data(forecast_9_1) -> forecast_9_1




# DSX - 10

dsx_10 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_2_months_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_2_months_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>%   
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_10_1


process_forecast_data(forecast_10_1) -> forecast_10_1



# DSX - 11

dsx_11 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(
    temp_filter = (
      (if("forecast_month_year_code" %in% names(.)) forecast_month_year_code == date_1_month_ago else FALSE) |
        (if("forecast_month_year_id" %in% names(.)) forecast_month_year_id == date_1_month_ago else FALSE)
    )
  ) %>%
  dplyr::filter(temp_filter) %>%
  dplyr::select(-temp_filter) %>%   
  dplyr::rename(mfg_loc = product_manufacturing_location_code,
                location = location_no,
                sku = product_label_sku_code,
                sku_description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group_no = product_group_code,
                group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  select_forecast_columns %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast_11_1


process_forecast_data(forecast_11_1) -> forecast_11_1



rbind(forecast_lag_1_1, forecast_1_1, forecast_2_1, forecast_3_1, forecast_4_1, forecast_5_1, forecast_6_1, forecast_7_1, 
      forecast_8_1, forecast_9_1, forecast_10_1, forecast_11_1) -> forecast1


forecast1 %>%
  dplyr::mutate(year = as.double(year),
                month = as.double(month),
                date_ref = paste0(year, "_", month, "_", mfg_loc, "_", sku)) %>% 
  dplyr::relocate(date_ref) -> forecast1


forecast1 %>% 
  dplyr::mutate(year = as.double(year),
                month = as.double(month),
                date_ref = paste0(year, "_", month, "_", mfg_loc, "_", sku))-> forecast_2_lag_1


forecast_2_lag_1 %>% 
  dplyr::left_join(sku_actual_pivot %>% select(-mfg_ref), by = "date_ref") %>% 
  dplyr::mutate(actual_shipped_cases = replace(actual_shipped_cases, is.na(actual_shipped_cases), 0)) %>% 
  dplyr::mutate(actual_shipped_lbs = replace(actual_shipped_lbs, is.na(actual_shipped_lbs), 0),
                adjusted_forecast_pounds_lbs = round(adjusted_forecast_pounds_lbs, 0)) -> raw_comsumption_comparison_lag_1


raw_comsumption_comparison_lag_1 %>% 
  dplyr::left_join(sales_orders_pivot %>% select(-mfg_ref), by = "date_ref") -> raw_comsumption_comparison_lag_1


# NA to 0
raw_comsumption_comparison_lag_1 %>% 
  dplyr::mutate(order_qty_final = replace(order_qty_final, is.na(order_qty_final), 0),
                order_qty_original = replace(order_qty_original, is.na(order_qty_original), 0)) ->  raw_comsumption_comparison_lag_1


raw_comsumption_comparison_lag_1 %>% 
  dplyr::left_join(bom_2) -> raw_comsumption_comparison_ver2_lag_1


raw_comsumption_comparison_ver2_lag_1 %>% 
  dplyr::mutate(forecasted_oil_qty = adjusted_forecast_cases * quantity_w_scrap,
                consumption_qty_actual_shipped = actual_shipped_cases * quantity_w_scrap,
                consumption_percent_adjusted_actual_shipped = consumption_qty_actual_shipped / forecasted_oil_qty) %>%
  
  dplyr::mutate(consumption_qty_sales_order_qty = order_qty_final * quantity_w_scrap,
                consumption_percent_adjusted_sales_order = consumption_qty_sales_order_qty / forecasted_oil_qty) %>% 
  
  
  dplyr::mutate(consumption_percent_adjusted_actual_shipped = replace(consumption_percent_adjusted_actual_shipped, is.na(consumption_percent_adjusted_actual_shipped) | is.nan(consumption_percent_adjusted_actual_shipped) | is.infinite(consumption_percent_adjusted_actual_shipped), 0)) %>% 
  dplyr::mutate(consumption_percent_adjusted_actual_shipped = sprintf("%1.2f%%", 100*consumption_percent_adjusted_actual_shipped)) %>% 
  dplyr::mutate(consumption_percent_adjusted_sales_order = replace(consumption_percent_adjusted_sales_order, is.na(consumption_percent_adjusted_sales_order) | is.nan(consumption_percent_adjusted_sales_order) | is.infinite(consumption_percent_adjusted_sales_order), 0)) %>% 
  dplyr::mutate(consumption_percent_adjusted_sales_order = sprintf("%1.2f%%", 100*consumption_percent_adjusted_sales_order)) %>% 
  
  
  dplyr::mutate(diff_between_forecast_actual =  forecasted_oil_qty - consumption_qty_actual_shipped,
                diff_between_forecast_original = forecasted_oil_qty - consumption_qty_sales_order_qty) %>% 
  
  
  dplyr::arrange(mfg_ref) %>% 
  dplyr::relocate(component, .after = group) -> raw_comsumption_comparison_ver2_lag_1


raw_comsumption_comparison_ver2_lag_1 %>% 
  dplyr::filter(mfg_loc != "-1") -> raw_comsumption_comparison_ver2_lag_1




# Duplicated values delete
raw_comsumption_comparison_ver2_lag_1[!duplicated(raw_comsumption_comparison_ver2_lag_1[,c("date_ref", "mfg_ref", "sku", "component", "quantity_w_scrap")]),] -> raw_comsumption_comparison_ver2_lag_1


raw_comsumption_comparison_ver2_lag_1 %>% 
  dplyr::mutate(adjusted_forecast_cases = round(adjusted_forecast_cases, 0),
                forecasted_oil_qty = round(forecasted_oil_qty, 0),
                consumption_qty_actual_shipped = round(consumption_qty_actual_shipped, 0),
                diff_between_forecast_actual = round(diff_between_forecast_actual, 0),
                consumption_qty_sales_order_qty = round(consumption_qty_sales_order_qty, 0),
                diff_between_forecast_original = round(diff_between_forecast_original, 0)) -> raw_comsumption_comparison_ver2_lag_1


# final touch
raw_comsumption_comparison_ver2_lag_1 %>% 
  dplyr::mutate(mfg_ref = gsub("_", "-", mfg_ref)) -> raw_comsumption_comparison_ver2_lag_1 

raw_comsumption_comparison_final_lag_1 <- raw_comsumption_comparison_ver2_lag_1

raw_comsumption_comparison_final_lag_1 %>% 
  dplyr::select(year, month, mfg_ref, mfg_loc, sku, sku_description, label, category, platform, group_no, group, component,
                quantity_w_scrap, adjusted_forecast_cases, forecasted_oil_qty, 
                actual_shipped_cases, 
                consumption_qty_actual_shipped, consumption_percent_adjusted_actual_shipped,
                diff_between_forecast_actual, order_qty_final, order_qty_original, consumption_qty_sales_order_qty, 
                consumption_percent_adjusted_sales_order, diff_between_forecast_original) %>% 
  dplyr::arrange(year, month, mfg_loc, sku) -> raw_comsumption_comparison_final_lag_1






##### Combine all files
raw_comsumption_comparison_final %>% 
  dplyr::mutate(date_ref = paste0(year, "_", month, "_", mfg_loc, "_", sku, "_", component)) %>% 
  dplyr::rename(adjusted_forecast_cases_lag0 = adjusted_forecast_cases,
                forecasted_oil_qty_lag0 = forecasted_oil_qty,
                consumption_percent_adjusted_actual_shipped_lag0 = consumption_percent_adjusted_actual_shipped,
                diff_between_forecast_actual_lag0 = diff_between_forecast_actual,
                consumption_percent_adjusted_sales_order_lag0 = consumption_percent_adjusted_sales_order,
                diff_between_forecast_original_lag0 = diff_between_forecast_original) -> raw_comsumption_comparison_final



raw_comsumption_comparison_final_lag_1 %>% 
  dplyr::mutate(date_ref = paste0(year, "_", month, "_", mfg_loc, "_", sku, "_", component)) %>% 
  dplyr::select(date_ref, adjusted_forecast_cases, forecasted_oil_qty, consumption_percent_adjusted_actual_shipped, diff_between_forecast_actual,
                consumption_percent_adjusted_sales_order, diff_between_forecast_original) %>% 
  dplyr::rename(adjusted_forecast_cases_lag1 = adjusted_forecast_cases,
                forecasted_oil_qty_lag1 = forecasted_oil_qty,
                consumption_percent_adjusted_actual_shipped_lag1 = consumption_percent_adjusted_actual_shipped,
                diff_between_forecast_actual_lag1 = diff_between_forecast_actual,
                consumption_percent_adjusted_sales_order_lag1 = consumption_percent_adjusted_sales_order,
                diff_between_forecast_original_lag1 = diff_between_forecast_original) -> raw_comsumption_comparison_final_lag_1





raw_comsumption_comparison_final %>% 
  dplyr::left_join(raw_comsumption_comparison_final_lag_1) %>% 
  dplyr::select(year, month, mfg_ref, mfg_loc, sku, sku_description, label, category, platform, group_no, group, component,
                  quantity_w_scrap, actual_shipped_cases, consumption_qty_actual_shipped, order_qty_final, order_qty_original,
                  consumption_qty_sales_order_qty,
                  adjusted_forecast_cases_lag0, forecasted_oil_qty_lag0, consumption_percent_adjusted_actual_shipped_lag0,
                  diff_between_forecast_actual_lag0, consumption_percent_adjusted_sales_order_lag0, diff_between_forecast_original_lag0,
                  adjusted_forecast_cases_lag1, forecasted_oil_qty_lag1, consumption_percent_adjusted_actual_shipped_lag1,
                  diff_between_forecast_actual_lag1, consumption_percent_adjusted_sales_order_lag1, diff_between_forecast_original_lag1) -> raw_comsumption_comparison_final




#######################################################################################################################
#######################################################################################################################
#######################################################################################################################

forecast_2_lag_1 %>% 
  dplyr::select(date_ref) -> f_1



# Take out oil skus only from sales order
sales_orders_pivot %>% 
  tidyr::separate(mfg_ref, c("mfg_loc", "sku"), sep = "_") %>% 
  dplyr::select(date_ref) -> s_1



dplyr::intersect(f_1, s_1) %>% 
  dplyr::mutate(both = "both exist") -> i_1


s_1 %>% 
  dplyr::left_join(i_1) %>% 
  dplyr::filter(is.na(both)) %>% 
  dplyr::select(date_ref) %>%
  dplyr::mutate(date_ref_2 = date_ref) %>% 
  tidyr::separate(date_ref_2, c("year", "month", "mfg_loc", "sku"), sep = "_") %>% 
  dplyr::mutate(ref = paste0(mfg_loc, "-", sku)) %>% 
  dplyr::relocate(date_ref, ref, mfg_loc, sku, year, month) %>% 
  dplyr::arrange(year, month, mfg_loc, sku) -> identitied_skus_not_existing



# Export to Excel File ----
writexl::write_xlsx(identitied_skus_not_existing, "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 25/Raw consumption Lag 1/Monthly recurring reports/05.15.2024/Identitied skus not existing.xlsx")



#######################################################################################################################
################################################# Missing SKU #########################################################
#######################################################################################################################

# completed sku list for missing sku category & platform
completed_sku_list[-1:-2, ] -> completed_sku_list
completed_sku_list %>% 
  janitor::clean_names() %>% 
  dplyr::select(x6, x9, x11) %>% 
  dplyr::rename(sku = x6,
                category = x9,
                platform = x11) %>% 
  dplyr::mutate(sku = gsub("-", "", sku)) -> completed_sku_list

completed_sku_list[!duplicated(completed_sku_list[,c("sku")]),] -> completed_sku_list




# sales order for missing sku
sales_orders %>% dplyr::select(sku, x8) -> sales_orders_for_missing_sku
sales_orders_for_missing_sku[!duplicated(sales_orders_for_missing_sku[,c("sku")]),] -> sales_orders_for_missing_sku

sales_orders_pivot %>% 
  dplyr::select(-mfg_ref) -> sales_orders_for_missing_sku_2

# sku actual for missing sku
sku_actual_pivot %>% 
  dplyr::select(date_ref, actual_shipped_cases) -> sku_actual_for_missing_sku



# missing sku final page
identitied_skus_not_existing %>% 
  data.frame() %>% 
  dplyr::rename(mfg_ref = ref) %>%
  dplyr::mutate(mfg_ref = gsub("-", "_", mfg_ref)) %>% 
  dplyr::relocate(date_ref, year, month, mfg_ref, mfg_loc, sku) %>% 
  dplyr::left_join(sales_orders_for_missing_sku) %>% 
  dplyr::rename(sku_description = x8) %>% 
  dplyr::mutate(label = stringr::str_sub(sku, 6, 9)) %>% 
  dplyr::left_join(completed_sku_list) %>% 
  dplyr::mutate(group_no = "n/a",
                group = "n/a",
                adjusted_forecast_cases_lag0 = 0,
                forecasted_oil_qty_lag0 = 0,
                consumption_percent_adjusted_actual_shipped_lag0 = "n/a",
                diff_between_forecast_actual_lag0 = "n/a",
                consumption_percent_adjusted_sales_order_lag0 = "n/a",
                diff_between_forecast_original_lag0 = "n/a",
                adjusted_forecast_cases_lag1 = 0,
                forecasted_oil_qty_lag1 = 0,
                consumption_percent_adjusted_actual_shipped_lag1 = "n/a",
                diff_between_forecast_actual_lag1 = "n/a",
                consumption_percent_adjusted_sales_order_lag1 = "n/a",
                diff_between_forecast_original_lag1 = "n/a") %>% 
  dplyr::left_join(bom_2) %>% 
  dplyr::left_join(sku_actual_for_missing_sku) %>% 
  dplyr::mutate(consumption_qty_actual_shipped = actual_shipped_cases * quantity_w_scrap) %>% 
  dplyr::left_join(sales_orders_for_missing_sku_2) %>% 
  dplyr::mutate(consumption_qty_sales_order_qty = order_qty_final * quantity_w_scrap) %>% 
  dplyr::select(year, month, mfg_ref, mfg_loc, sku, sku_description, label, category, platform, group_no, group, component,
                quantity_w_scrap, actual_shipped_cases, consumption_qty_actual_shipped, order_qty_final,
                order_qty_original,	consumption_qty_sales_order_qty, 
                
                adjusted_forecast_cases_lag0, forecasted_oil_qty_lag0, consumption_percent_adjusted_actual_shipped_lag0, diff_between_forecast_actual_lag0, 
                consumption_percent_adjusted_sales_order_lag0, diff_between_forecast_original_lag0, 
                
                adjusted_forecast_cases_lag1, forecasted_oil_qty_lag1, consumption_percent_adjusted_actual_shipped_lag1, diff_between_forecast_actual_lag1, 
                consumption_percent_adjusted_sales_order_lag1, diff_between_forecast_original_lag1) %>% 
  
  dplyr::arrange(year, month, mfg_loc, sku) -> identitied_skus_not_existing_2




#######################################################################################################################
################################################# Final Paper #########################################################
#######################################################################################################################

raw_comsumption_comparison_final %>% 
  dplyr::mutate(dsx = "Y") -> raw_comsumption_comparison_final


identitied_skus_not_existing_2 %>% 
  dplyr::mutate(dsx = "N") -> identitied_skus_not_existing_2

rbind(raw_comsumption_comparison_final, identitied_skus_not_existing_2) -> final_paper



final_paper %>% 
  dplyr::filter(!(year == year(Sys.Date()) & month == month(Sys.Date()))) -> final_paper


final_paper %>% 
  dplyr::mutate(mfg_ref = gsub("_", "-", mfg_ref)) -> final_paper

final_paper %>% 
  dplyr::left_join(rm_to_sku %>% dplyr::select(component, comp_description) %>% dplyr::distinct(component, .keep_all = TRUE), by = "component") -> final_paper





colnames(final_paper)[1]	<-	"Year"
colnames(final_paper)[2]	<-	"Month"
colnames(final_paper)[3]	<-	"ref"
colnames(final_paper)[4]	<-	"Location"
colnames(final_paper)[5]	<-	"SKU (FG)"
colnames(final_paper)[6]	<-	"Description"
colnames(final_paper)[7]	<-	"Label"
colnames(final_paper)[8]	<-	"Category"
colnames(final_paper)[9]	<-	"Platform"
colnames(final_paper)[10]	<-	"Group Code"
colnames(final_paper)[11]	<-	"Group Name"
colnames(final_paper)[12]	<-	"Component"
colnames(final_paper)[13]	<-	"Quantity w/Scrap"
colnames(final_paper)[14]	<-	"Actual Shipped Cases"
colnames(final_paper)[15]	<-	"Consumption Quantity (Actual Shipped)"
colnames(final_paper)[16]	<-	"Sales Order Qty Final (Cases)"
colnames(final_paper)[17]	<-	"Sales Order Qty Original (Cases)"
colnames(final_paper)[18]	<-	"Consumption Quantity (Original Sales Order Qty)"
colnames(final_paper)[19]	<-	"Adjusted Forecast Cases (Lag 0)"
colnames(final_paper)[20]	<-	"Forecasted Oil Qty (Lag 0)"
colnames(final_paper)[21]	<-	"Consumption % (by Adjusted forecast - Actual Shipped) (Lag 0)"
colnames(final_paper)[22]	<-	"Diff (Forecasted - Actual Shipped) (Lag 0)"
colnames(final_paper)[23]	<-	"Consumption % (by Adjusted forecast - Original Sales Order Qty) (Lag 0)"
colnames(final_paper)[24]	<-	"Diff (Forecasted - Original Sales Order) (Lag 0)"
colnames(final_paper)[25]	<-	"Adjusted Forecast Cases (Lag 1)"
colnames(final_paper)[26]	<-	"Forecasted Oil Qty (Lag 1)"
colnames(final_paper)[27]	<-	"Consumption % (by Adjusted forecast - Actual Shipped) (Lag 1)"
colnames(final_paper)[28]	<-	"Diff (Forecasted - Actual Shipped) (Lag 1)"
colnames(final_paper)[29]	<-	"Consumption % (by Adjusted forecast - Original Sales Order Qty) (Lag 1)"
colnames(final_paper)[30]	<-	"Diff (Forecasted - Original Sales Order) (Lag 1)"
colnames(final_paper)[31]	<-	"DSX"
colnames(final_paper)[32]	<-	"Comp Description"


final_paper %>% 
  dplyr::filter(Location != "-1") -> final_paper



######################## Split the data frame in half #######################
# Define the function to split the final_paper dataframe in half based on a given column's median value
split_final_paper <- function(final_paper, column) {
  if (!column %in% names(final_paper)) {
    stop(paste("Column", column, "not found in the final_paper dataframe."))
  }
  
  # Ensure that the column can be coerced to numeric
  numeric_column <- as.numeric(final_paper[[column]])
  # Handle NA values which might cause the median function to fail
  if (any(is.na(numeric_column))) {
    warning("NA values found in numeric_column. Excluding them from median calculation.")
    numeric_column <- na.omit(numeric_column)
  }
  cutoff_value <- median(numeric_column)
  
  first_half <- final_paper %>%
    filter(!!as.symbol(column) <= cutoff_value)
  
  second_half <- final_paper %>%
    filter(!!as.symbol(column) > cutoff_value)
  
  return(list(first_half = first_half, second_half = second_half))
}


result <- split_final_paper(final_paper, 'Location')

# The result will be a list containing first_half and second_half dataframes
result$first_half -> final_paper_1
result$second_half -> final_paper_2


# Export to Excel File ----
write_xlsx(final_paper_1, "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 25/Raw consumption Lag 1/Monthly recurring reports/05.15.2024/raw consumption comparison_1.xlsx")
write_xlsx(final_paper_2, "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 25/Raw consumption Lag 1/Monthly recurring reports/05.15.2024/raw consumption comparison_2.xlsx")


