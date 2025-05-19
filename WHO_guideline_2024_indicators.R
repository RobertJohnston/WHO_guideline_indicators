# WHO_guideline_2024_indicators

# Calculation of prevalence of all screening criteria for
# screening criteria included in 2024 WHO Guideline

# Clear environment
rm(list = ls())

# install.packages("matrixStats")
# install.packages("labelled")
# install.packages("expss")

# Load libraries
library(readr)
library(haven)
library(ggplot2)
library(labelled)
library(matrixStats)
library(expss)
library(dplyr)
library(openxlsx)

# Host-specific setting of hostname
hostname <- Sys.info()[['nodename']]   

# Setting work directory based on host
if (hostname == "992224APL0X0061") {
  # Robert UNICEF PC
  workdir <- "C:/Users/rojohnston/UNICEF/Data and Analytics Nutrition - Analysis Space/Child Anthropometry/1- Anthropometry Analysis Script/Prepped Country Data Files/CSV"
} else if (hostname == "MY-LAPTOP") {
  # Your laptop
  workdir <- "D:/Projects/Seasonality/"
} else {
  stop("Unrecognized hostname, Set 'workdir' manually.")
}

# Set other directories
datadir <- file.path(workdir, "Data")

search_name = "Burkina"

# Collect list of files with name of country included
files <- list.files(path = workdir, pattern = search_name, full.names = TRUE)
file_names <- basename(files)
print(file_names)

# Burkina_Faso-2010-SMART-ANT.dta
# Burkina_Faso-2011-SMART-ANT.dta
# Burkina_Faso-2012-SMART-ANT.dta
# Burkina_Faso-2013-SMART-ANT.dta
# Burkina_Faso-2014-LSMS-ANT.dta
# Burkina_Faso-2014-SMART-ANT.dta
# Burkina_Faso-2015-SMART-ANT.dta
# Burkina_Faso-2016-SMART-ANT.dta
# Burkina_Faso-2017-SMART-ANT.dta
# Burkina_Faso-2018-SMART-ANT.dta
# Burkina_Faso-2019-SMART-ANT.dta
# Burkina_Faso-2020-Micronutrient-ANT.dta
# Burkina_Faso-2020-SMART-ANT.dta
# Burkina_Faso-2021-DHS-ANT.dta
# Burkina_Faso-2021-SMART-ANT.dta
# Burkina_Faso-2022-SMART-ANT.dta

# read_csv - includes the indicator label names.  read_dta does not. 

# Loop over filenames 

for (file in file_names) {
    df <- read_csv(file.path(workdir, file))

# df <- read_csv(file.path(workdir, "Burkina_Faso-1993-DHS-ANT.csv" ))
# View(df)
  
  # Valid data review
  # View(df %>% filter(is.na(sample_wgt)))
  # View(df %>% filter(is.na(sex)))
  
  # if sex is missing, then z-scores cannot be calculated
  # Change this line if a dataset of only muac is used. 
  fre(df$sex)
  df <- df %>% filter(!is.na(sex))

  
  # WAZ
  # WHZ
  # MUAC
  # Oedema
  # Agedays
  # Not BF
  
  # Data Cleaning
  
  # create sample_wgt
  df <- df %>% mutate(sample_wgt = sw)
  summary(df$sample_wgt)
  
  df <- df %>% mutate(Region= gregion)
  
  # Test if variables are completely missing  ?
  indicators <- c("sex", "agemons", "waz", "whz","measure", "muac","oedema")
  sapply(df[indicators], function(x) all(is.na(x)))
  
  indicators_original <- paste0(indicators, "_original")
  indicators_original <- indicators_original[indicators_original %in% names(df)]
  sapply(df[indicators_original], function(x) all(is.na(x)))
  
  # If oedema is missing, replace oedema with oedema_original
  if ("oedema_original" %in% names(df) && all(is.na(df$oedema))) {
    df$oedema <- df$oedema_original
  }
  
  # is MUAC saved as CM or MM ?
  # If ave(muac) > 110 then saved in MM
  if ("muac" %in% names(df) && all(!is.na(df$muac))) {  # if muac is not present or all missing - skip
    if (mean(df$muac, na.rm = TRUE) > 12 & mean(df$muac, na.rm = TRUE) < 18) {
      df$muac <- df$muac * 10  # 1 cm = 10 mm
    }
  }
  
  # sev_wast		"Severely wasted child under 5 years - WHZ"
  # wast			  "Wasted child under 5 years - WHZ"
  # mean_whz		"Mean z-score for weight-for-height for children under 5 years"
  # sev_uwt	"Severely underweight child under 5 years"
  # uwt		  "Underweight child under 5 years"
  # mean_waz		"Mean weight-for-age for children under 5 years"
  # muac_110  "Infant with severe wasting by MUAC (MUAC < 110mm) " 
  # muac_115  "Child with severe wasting by MUAC (MUAC < 115mm)" 
  # muac_125  "Child with  wasting by MUAC (MUAC < 125mm)" 
  # edema 
  # not_bf "Infant not breastfed" 
  
  # Infants under 6 months of age at risk of poor growth and development 
  
  # Percentage of infant contacts with WAZ <-2 SD of WHO child growth standards
  # Percentage of infant contacts with severe wasting by WHZ (WHZ < - 3 SD of WHO child growth standards)
  # Percentage of infant contacts with severe wasting by MUAC (MUAC < 110mm)
  # Percentage of infant contacts with nutritional oedema 
  # Percentage of infant contacts with recent weight loss (not available from surveys)
  # Percentage of infant contacts with ineffective breastfeeding, feeding concerns if not breastfed
  # Percentage of infant contacts with IMCI danger signs or acute medical problems under severe classification as per IMCI (not available from surveys)
  # COMBINED AT RISK
  
  # Children from 6 to 59 months of age 
  
  # Percentage of 6-59M child contacts with severe wasting (WHZ < - 3 SD of WHO child growth standards)
  # Percentage of 6-59M child contacts with severe wasting by MUAC (MUAC < 115mm)
  # Percentage of 6-59M child contacts with nutritional oedema 
  # Combined SAM
  
  
  # Underweight 
  df <- df %>%
    mutate(uwt =
             case_when(
               waz < -2  ~ 1,
               waz >= -2 ~ 0,
               is.na(whz) ~ NA_real_  # handle missing values
             )) %>%
    set_value_labels(uwt = c("Yes" = 1, "No" = 0)) %>%
    set_variable_labels(uwt = "WAZ<-2SD")
  
  # Severely wasted 
  df <- df %>%
    mutate(sev_wast =
             case_when(
               whz < -3  ~ 1,
               whz >= -3 ~ 0,
               is.na(whz) ~ NA_real_  # handle missing values
             )) %>%
    set_value_labels(sev_wast = c("Yes" = 1, "No" = 0)) %>%
    set_variable_labels(sev_wast = "WHZ<-3SD")
  
  # wasted 
  df <- df %>%
    mutate(wast =
             case_when(
               whz < -2  ~ 1,
               whz >= -2 ~ 0,
               is.na(whz) ~ NA_real_  # handle missing values
             )) %>%
    set_value_labels(wast = c("Yes" = 1, "No" = 0)) %>%
    set_variable_labels(wast = "WHZ<-2SD")
  
  # MUAC 125 
  df <- df %>%
    mutate(muac_125 =
             case_when(
               muac < 125  ~ 1,
               muac >= 125 ~ 0,
               is.na(muac) ~ NA_real_  # handle missing values
             )) %>%
    set_value_labels(muac_125 = c("Yes" = 1, "No" = 0)) %>%
    set_variable_labels(muac_125 = "MUAC<125mm")
  
  # MUAC 115 
  df <- df %>%
    mutate(muac_115 =
             case_when(
               muac < 115  ~ 1,
               muac >= 115 ~ 0,
               is.na(muac) ~ NA_real_  # handle missing values
             )) %>%
    set_value_labels(muac_115 = c("Yes" = 1, "No" = 0)) %>%
    set_variable_labels(muac_115 = "MUAC<115mm")
  
  # MUAC 110 
  df <- df %>%
    mutate(muac_110 =
             case_when(
               muac < 110 ~ 1,
               muac >= 110 ~ 0,
               is.na(muac) ~ NA_real_  # handle missing values
             )) %>%
    set_value_labels(muac_110 = c("Yes" = 1, "No" = 0)) %>%
    set_variable_labels(muac_110 = "MUAC<110mm")
  
  #  Oedema
  # if oedema is not recoded, check oedema_original
  df <- df %>%
    mutate(oedema =
             case_when(
               oedema == "Oui" ~ 1,
               oedema == "n" ~ 0,
               is.na(oedema) ~ NA_real_  # handle missing values
             )) %>%
    set_value_labels(oedema = c("Yes" = 1, "No" = 0)) %>%
    set_variable_labels(oedema = "bilateral oedema")
  
  # At Risk Combined
  df <- df %>%
    mutate(
      valid_inputs = rowSums(!is.na(across(c(uwt, sev_wast, muac_110, oedema)))),
      at_risk = case_when(
        valid_inputs == 0 ~ NA_real_,
        rowSums(across(c(uwt, sev_wast, muac_110, oedema)) == 1, na.rm = TRUE) > 0 ~ 1,
        TRUE ~ 0
      )
    ) %>%
    set_value_labels(at_risk = c("Yes" = 1, "No" = 0)) %>%
    set_variable_labels(at_risk = "At Risk Combined")
  
  
  # Severe Acute Malnutrition 
  df <- df %>%
    mutate(
      valid_inputs = rowSums(!is.na(across(c(sev_wast, muac_115, oedema)))),
      sam = case_when(
        valid_inputs == 0 ~ NA_real_,
        rowSums(across(c(sev_wast, muac_115, oedema)) == 1, na.rm = TRUE) > 0 ~ 1,
        TRUE ~ 0
      )
    ) %>%
    set_value_labels(sam = c("Yes" = 1, "No" = 0)) %>%
    set_variable_labels(sam = "SAM Combined")
  
  # Global Acute Malnutrition 
  df <- df %>%
    mutate(
      valid_inputs = rowSums(!is.na(across(c(wast, muac_125, oedema)))),
      gam = case_when(
        valid_inputs == 0 ~ NA_real_,
        rowSums(across(c(wast, muac_125, oedema)) == 1, na.rm = TRUE) > 0 ~ 1,
        TRUE ~ 0
      )
    ) %>%
    set_value_labels(gam = c("Yes" = 1, "No" = 0)) %>%
    set_variable_labels(gam = "GAM Combined")
  
  df <- df %>%
    mutate(blank = NA) %>%
    set_variable_labels(blank = "-")
  
  
  # fre(df$wast)
  # fre(df$sev_wast)
  # cro_rpct(df$Region, df$sev_wast)
  # fre(df$muac_125)
  # fre(df$muac_115)
  # fre(df$oedema)
  # 
  # fre(df$sam)
  # fre(df$gam)
  
  
  # View valid N of all indicators
  
  # List of indicators
  indicators <- c("sev_wast", "muac_115", "oedema", "sam", "wast", "muac_125", "gam")
  
  df %>%
    group_by(Region) %>%
    summarise(
      across(all_of(indicators), ~ sum(!is.na(.)), .names = "{.col}_N"),
      .groups = "drop"
    )
  total_row <- df %>%
    summarise(
      across(all_of(indicators), ~ sum(!is.na(.)), .names = "{.col}_N")
    ) %>%
    mutate(Region = "Total") %>%
    select(Region, everything())
  
  # Combine
  valid_n_table <- bind_rows(
    df %>%
      group_by(Region) %>%
      summarise(
        across(all_of(indicators), ~ sum(!is.na(.)), .names = "{.col}_N"),
        .groups = "drop"
      ),
    total_row
  )
  
  View(valid_n_table)
  
  
  # summarise function (used in tables)
  summarise_prev_table <- function(data) {
    data %>%
      filter(!is.na(sample_wgt)) %>%
      summarise(across(
        all_of(indicators),
        list(
          `%` = ~ round(weighted.mean(.x, sample_wgt, na.rm = TRUE) * 100, 1),
          N   = ~ sum(!is.na(.x))
        ),
        .names = "{.col} ({.fn})"
      ))
  }
  
  replace_names_with_labels <- function(df_table, reference_df, indicators) {
    label_lookup <- sapply(indicators, function(x) var_lab(reference_df[[x]]), USE.NAMES = TRUE)
    names(df_table) <- sapply(names(df_table), function(name) {
      if (name == "gregion") return(name)
      match <- regexec("^(.+?)\\s*(\\(.*\\))$", name)
      parts <- regmatches(name, match)[[1]]
      if (length(parts) == 3) {
        var <- parts[2]
        suffix <- parts[3]
        label <- label_lookup[[var]]
        if (is.null(label) || label == "") label <- var
        return(paste0(label, " ", suffix))
      } else {
        return(name)
      }
    })
    df_table
  }
  
  # **************************************************************************************************
  # * Anthropometric indicators for children under age 6 months
  # **************************************************************************************************
  
  df_0_5m <- df %>% filter(agemons >=0 & agemons < 6)
  
  # AT RISK TABLE
  indicators <- c("sev_wast", "muac_110",  "oedema", "uwt", "at_risk")
  
  table_at_risk_0_5m <- df_0_5m %>%
    group_by(Region) %>%
    summarise_prev_table()
  
  total_row <- summarise_prev_table(df_0_5m) %>%
    mutate(Region = "Total") %>%
    select(Region, everything())
  
  table_at_risk_0_5m <- bind_rows(table_at_risk_0_5m, total_row)
  for (var in indicators) {
    pct_col <- paste0(var, " (%)")
    n_col   <- paste0(var, " (N)")
    
    if (pct_col %in% names(table_at_risk_0_5m) && n_col %in% names(table_at_risk_0_5m)) {
      table_at_risk_0_5m[[pct_col]] <- ifelse(
        is.na(table_at_risk_0_5m[[n_col]]) | table_at_risk_0_5m[[n_col]] < 30,
        " - ",
        as.character(table_at_risk_0_5m[[pct_col]])
      )
    }
  }
  
  table_at_risk_0_5m <- replace_names_with_labels(table_at_risk_0_5m, df_0_5m, indicators)
  View(table_at_risk_0_5m)
  
  # **************************************************************************************************
  # * Anthropometric indicators for children from 6- 59 months
  # **************************************************************************************************
  
  df_6_59m <- df %>% filter(agemons > 5 & agemons < 60)
  
  
  # SAM TABLE
  indicators <- c("sev_wast", "muac_115", "oedema", "sam")
  
  table_sam_6_59m <- df_6_59m %>%
    group_by(Region) %>%
    summarise_prev_table()
  
  total_row <- summarise_prev_table(df_6_59m) %>%
    mutate(Region = "Total") %>%
    select(Region, everything())
  
  table_sam_6_59m <- bind_rows(table_sam_6_59m, total_row)
  for (var in indicators) {
    pct_col <- paste0(var, " (%)")
    n_col   <- paste0(var, " (N)")
    
    if (pct_col %in% names(table_sam_6_59m) && n_col %in% names(table_sam_6_59m)) {
      table_sam_6_59m[[pct_col]] <- ifelse(
        is.na(table_sam_6_59m[[n_col]]) | table_sam_6_59m[[n_col]] < 30,
        " - ",
        as.character(table_sam_6_59m[[pct_col]])
      )
    }
  }
  table_sam_6_59m <- replace_names_with_labels(table_sam_6_59m, df_6_59m, indicators)
  View(table_sam_6_59m)
  
  
  # GAM TABLE
  indicators <- c("wast", "muac_125", "oedema", "gam")
  
  table_gam_6_59m <- df_6_59m %>%
    group_by(Region) %>%
    summarise_prev_table()
  
  total_row <- summarise_prev_table(df_6_59m) %>%
    mutate(Region = "Total") %>%
    select(Region, everything())
  
  table_gam_6_59m <- bind_rows(table_gam_6_59m, total_row)
  for (var in indicators) {
    pct_col <- paste0(var, " (%)")
    n_col   <- paste0(var, " (N)")
    
    if (pct_col %in% names(table_gam_6_59m) && n_col %in% names(table_gam_6_59m)) {
      table_gam_6_59m[[pct_col]] <- ifelse(
        is.na(table_gam_6_59m[[n_col]]) | table_gam_6_59m[[n_col]] < 30,
        " - ",
        as.character(table_gam_6_59m[[pct_col]])
      )
    }
  }
  table_gam_6_59m <- replace_names_with_labels(table_gam_6_59m, df_6_59m, indicators)
  View(table_gam_6_59m)
  
  
  # Remove everything before the first dash and after -ANT.csv
  cleaned_name <- sub("^[^-]+-", "", file)              # Remove before first dash
  cleaned_name <- sub("-ANT\\.csv$", "", cleaned_name)  # Remove -ANT.csv at end
  print(cleaned_name)
  
  
  
  country_name <- df$country[!is.na(df$country)][1]
  survey_name  <- df$survey[!is.na(df$survey)][1]
  survey_year  <- df$year[!is.na(df$year)][1]
  start <- min(df$date_measure, na.rm = TRUE)
  end <- max(df$date_measure, na.rm = TRUE)
  
  
  # Define file, sheet, and cell position
  file_path <- paste0("C:/Users/rojohnston/Downloads/WHO_indicators_", country_name, ".xlsx")
  sheet_name <- cleaned_name # use 
  start_cell <- "B2"
  
  if (!file.exists(file_path)) {
    wb <- createWorkbook()
  } else {
    wb <- loadWorkbook(file_path)
    
    if (sheet_name %in% names(wb)) {
      removeWorksheet(wb, sheet_name)  # drop if not clean
    }
  }
  addWorksheet(wb, sheet_name)
  
  
  
  x = 2
  y = 5
  add_y = length(unique(df$Region)) +4  # add rows between each pasted table
  
  # Write Country, Survey Type, Start and End Date
  note1 <- paste("Country:", country_name,"   Survey:", survey_name, survey_year)
  note2 <- paste("Survey data collection from", start, "to", end)
  
  writeData(wb, sheet = sheet_name, x = "WHO Guideline 2024 - Indicators for at risk and acute malnutrition", startCol = 2, startRow = 1)
  writeData(wb, sheet = sheet_name, x = note1, startCol = 2, startRow = 2)
  writeData(wb, sheet = sheet_name, x = note2, startCol = 2, startRow = 3)
  
  writeData(wb, sheet = sheet_name, x = "Infants from 0-5m at risk of poor growth and development", startCol = x, startRow = y)
  writeData(wb, sheet = sheet_name, x = table_at_risk_0_5m, startCol = x, startRow = y+1)
  y = y + add_y
  
  writeData(wb, sheet = sheet_name, x = "Children 6-59m with Severe Acute Malnutrition", startCol = x, startRow = y)
  writeData(wb, sheet = sheet_name, x = table_sam_6_59m, startCol = x, startRow = y+1)
  y = y + add_y
  
  writeData(wb, sheet = sheet_name, x = "Children 6-59m with Global Acute Malnutrition", startCol = x, startRow = y)
  writeData(wb, sheet = sheet_name, x = table_gam_6_59m, startCol = x, startRow = y+1)
  
  # Save the file
  saveWorkbook(wb, file_path, overwrite = TRUE)
  
}

END









whz_summary <- df %>%
  group_by(Region) %>%
  summarise(
    count = sum(!is.na(whz)),
    mean_whz = mean(whz, na.rm = TRUE),
    median_whz = median(whz, na.rm = TRUE),
    sd_whz = sd(whz, na.rm = TRUE),
    min_whz = min(whz, na.rm = TRUE),
    max_whz = max(whz, na.rm = TRUE)
  ) %>%
  arrange(Region)

print(whz_summary)



# WHZ distributions
df_clean <- df %>%
  filter(!is.na(whz), !is.na(Region))

# Plot: Histogram as % + Gaussian density curve
ggplot(df_clean, aes(x = whz)) +
  geom_histogram(aes(y = ..density.. * 100), 
                 binwidth = 0.2, fill = "skyblue", color = "white") +
  stat_function(fun = function(x) dnorm(x, mean = 0, sd = 1) * 100,
                color = "gray", size = 1) +
  facet_wrap(~ Region) +
  scale_y_continuous(name = "Percent Density") +
  scale_x_continuous(name = "WHZ", breaks = seq(-5, 5, by = 1)) +
  labs(title = "WHZ Percent Distribution with Gaussian Curve by Region") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 0, hjust = 1))


# MUAC
df_clean <- df %>%
  filter(!is.na(whz), !is.na(Region))

# Plot: Histogram as % + Gaussian density curve
muac_min <- floor(min(df_clean$muac, na.rm = TRUE) / 10) * 10
muac_max <- ceiling(max(df_clean$muac, na.rm = TRUE) / 10) * 10

# Plot with dynamic x-axis
ggplot(df_clean, aes(x = muac)) +
  geom_histogram(aes(y = ..density.. * 100), 
                 binwidth = 0.2, fill = "pink", color = "pink") +
  facet_wrap(~ Region) +
  scale_y_continuous(name = "Percent Density") +
  scale_x_continuous(
    name = "MUAC",
    breaks = seq(muac_min, muac_max, by = 10),
    limits = c(muac_min, muac_max)
  ) +
  labs(title = "MUAC Percent Distribution by Region") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 90, hjust = 1))

# table_temp <-  df %>% 
#   calc_cro_rpct(
#     cell_vars = list(Region, total()),
#     col_vars = list(sev_wast, wast, mean_haz,
#                     # , uwt, mean_waz, ,muac_110, muac_115, muac_125, edema 
#                     ),
#     weight = sample_wgt,
#     total_label = "Weighted N",
#     total_statistic = "w_cases",
#     total_row_position = c("below"),
#     expss_digits(digits=1)) %>%
#   set_caption("Child's anthropometric indicators")
# #Note the mean haz, whz, and waz are computed for the total and not by each background variable. 
# write.xlsx(table_temp, "WHO_guideline_indicators.xls", sheetName = "WHO_criteria", append=TRUE)


