# WHO_guideline_2024_indicators


# CHECKS

# double check oedema

# If SD of WAZ or WHZ is >1.2 then calculate prev from mean. 

# Add proportion WHZ only / MUAC only / WHZ + MUAC 

# Set minimum of graph label to zero


# Purpose of the code :
# Calculation of prevalence of all screening criteria for
# screening criteria included in 2024 WHO Guideline
# & Methods to Estimate Caseloads of Children with Moderate Wasting using Individual Risk Factors.

# Clear environment
rm(list = ls())

# Host-specific setting of hostname
hostname <- Sys.info()[['nodename']]  # use Sys.info()[["nodename"]] to find hostname

# Setting work directory based on host
if (hostname == "992224APL0X0061") {
  # Robert UNICEF PC
  workdir <- "C:/Users/rojohnston/UNICEF"
} else if (hostname == "DESKTOP-IMTUODA") {
  # Your laptop
  workdir <- "C:/Users/stupi/UNICEF"
} else {
  stop("Unrecognized hostname, Set 'workdir' manually.")
}

# Set other directories
datadir <- file.path(workdir, "Data and Analytics Nutrition - Analysis Space/Child Anthropometry/1- Anthropometry Analysis Script/Prepped Country Data Files/CSV")
outputdir <- file.path(workdir, "Data and Analytics Nutrition - Working Space/Wasting Cascade/WHO 2024 Country Profiles")

search_name = "Burkina"


# install.packages("matrixStats")
# install.packages("labelled")
# install.packages("expss")
# install.packages("maditr")

# Load libraries
library(readr)
library(haven)
library(ggplot2)
library(labelled)
library(matrixStats)
library(expss)
library(dplyr)
library(openxlsx)

# Collect list of files with name of country included
files <- list.files(path = datadir, pattern = search_name, full.names = TRUE)
file_names <- basename(files)
print(file_names)

# NOTE read_csv - includes the indicator label names.  read_dta does not. 

# To open one specific dataframe
# file_names <- "South_Sudan-2018-FSNMS_23-ANT.csv"

# Loop over filenames 
for (file in file_names) {
    df <- read_csv(file.path(datadir, file))
    # df <- read_csv(file.path(datadir, file, locale = locale(encoding = "UTF-8")))
    # df <- read_csv(file.path(datadir, file, locale = locale(encoding = "ISO-8859-1")))  # European languages
    # df <- read_csv(file.path(datadir, file, locale = locale(encoding = "Windows-1252"))) # Older windows excel

  # Data Cleaning
  
  # create sample_wgt
  df <- df %>% mutate(sample_wgt = sw)
  # summary(df$sample_wgt)
  
  df <- df %>% mutate(Region = iconv(gregion, from = "", to = "UTF-8", sub = "byte"))
  # fre(df$gregion)
  
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
  # fre(df$oedema)

  # If MUAC is saved in CM, convert to MM
  if ("muac" %in% names(df) && any(!is.na(df$muac))) {  # if muac is present or not all missing - continue
    if (mean(df$muac, na.rm = TRUE) < 26) {  # 26 is max in CM that can be measured with child MUAC tape. 
      df$muac <- df$muac * 10  # 1 cm = 10 mm
    }
  }
  
  # If MUAC = outlier, set to missing 
  df$muac <- ifelse(df$muac < 30 | df$muac > 260, NA_real_, df$muac)
  
  # Convert date_measure to a date
  df <- df %>% mutate(date_measure = as.Date(date_measure, format = "%d%b%Y"))
  
  
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
  # Percentage of infant contacts with wasting by WHZ (WHZ <-2 SD of WHO child growth standards)
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
  
  # Moderate Wasting at risk
  # muac_115_125
  # sev_uwt
  # mod_muac_suwt
  # muac_115_125_24m
  # sev_uwt_24m
  # mod_muac_suwt_24m
  
  # Underweight 
  df <- df %>%
    mutate(uwt = wazB) %>%
    # mutate(uwt =
    #          case_when(
    #            waz <  -2~ 1,
    #            waz >= -2 ~ 0,
    #            is.na(waz) ~ NA_real_  # handle missing values
    #          )) %>%
    set_value_labels(uwt = c("Yes" = 1, "No" = 0)) %>%
    set_variable_labels(uwt = "WAZ<-2SD")
  
  # Severe Underweight 
  df <- df %>%
    mutate(sev_uwt = wazA) %>%
    # mutate(sev_uwt =
    #          case_when(
    #            waz <  -3 ~ 1,
    #            waz >= -3 ~ 0,
    #            is.na(waz) ~ NA_real_  # handle missing values
    #          )) %>%
    set_value_labels(sev_uwt = c("Yes" = 1, "No" = 0)) %>%
    set_variable_labels(sev_uwt = "WAZ<-3SD")
    if (!"sev_uwt" %in% names(df)) stop("Variable 'sev_uwt' does not exist in the dataset.")
  
  # Severe Underweight in Children < 24M 
  df <- df %>%
    mutate(sev_uwt_24m = if_else(agemons < 24, wazA, NA_real_)) %>%
    # mutate(sev_uwt_24m =
    #          case_when(
    #            agemons >=24 ~ NA_real_, 
    #            waz <  -3 ~ 1,
    #            waz >= -3 ~ 0,
    #            is.na(waz) ~ NA_real_  # handle missing values
    #          )) %>%
    set_value_labels(sev_uwt_24m = c("Yes" = 1, "No" = 0)) %>%
    set_variable_labels(sev_uwt_24m = "WAZ<-3SD in children <24M")
  
  # Wasted 
  df <- df %>%
    mutate(wast = whzB) %>%
    # mutate(wast =
    #          case_when(
    #            whz < -2  ~ 1,
    #            whz >= -2 ~ 0,
    #            is.na(whz) ~ NA_real_  # handle missing values
    #          )) %>%
    set_value_labels(wast = c("Yes" = 1, "No" = 0)) %>%
    set_variable_labels(wast = "WHZ<-2SD")
  
  # Severely wasted 
  df <- df %>%
    mutate(sev_wast = whzA) %>%
        # mutate(sev_wast =
    #          case_when(
    #            whz < -3  ~ 1,
    #            whz >= -3 ~ 0,
    #            is.na(whz) ~ NA_real_  # handle missing values
    #          )) %>%
    set_value_labels(sev_wast = c("Yes" = 1, "No" = 0)) %>%
    set_variable_labels(sev_wast = "WHZ<-3SD")
    if (!"sev_wast" %in% names(df)) stop("Variable 'sev_wast' does not exist in the dataset.")
  
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
  
  # MUAC 110 for infants from 6 weeks to <6 months of age 
  df <- df %>%
    mutate(muac_110 =
             case_when(
               muac < 110 ~ 1,
               muac >= 110 ~ 0,
               is.na(muac) ~ NA_real_,  # handle missing values
               agemons < 1.5 ~ NA_real_  # If <6 weeks, set to missing
             )) %>%
    set_value_labels(muac_110 = c("Yes" = 1, "No" = 0)) %>%
    set_variable_labels(muac_110 = "MUAC<110mm")
  
  #muac_115_119
  df <- df %>%
    mutate(muac_115_119 =
             case_when(
               muac >= 115 & muac < 120 ~ 1,
               muac < 115 | muac >= 120 ~ 0,
               is.na(muac) ~ NA_real_  # handle missing values
             )) %>%
    set_value_labels(muac_115_119 = c("Yes" = 1, "No" = 0)) %>%
    set_variable_labels(muac_115_119 = "MUAC 115-119mm")
  
  # muac_115_119_24m - muac_115_119 in children under 24M
  df <- df %>%
    mutate(muac_115_119_24m  =
             case_when(
               agemons >=24 ~ NA_real_, 
               muac >= 115 & muac < 120 ~ 1,
               muac < 115 | muac >= 120 ~ 0,
               is.na(muac) ~ NA_real_  # handle missing values
             )) %>%
    set_value_labels(muac_115_119_24m  = c("Yes" = 1, "No" = 0)) %>%
    set_variable_labels(muac_115_119_24m  = "MUAC 115-119mm <24M")
  
# mod_muac_suwt - Variable representing combined condition of 
# child under 59m who has muac_115_119 AND sev_uwt
  df <- df %>%
    mutate(
      # Check if both inputs are available in a row
      mod_muac_suwt = case_when(
        is.na(sev_uwt) | is.na(muac_115_119) ~ NA_real_,  # Either input missing
        sev_uwt == 1 & muac_115_119 == 1 ~ 1,             # Both are 1
        TRUE ~ 0                                          # Otherwise 0
      )
    ) %>%
    set_value_labels(mod_muac_suwt = c("Yes" = 1, "No" = 0)) %>%
    set_variable_labels(mod_muac_suwt = "MUAC 115-119 & Severe UWT in children under 59M")
  
# mod_muac_suwt_24m - Variable representing combined condition of 
# child under 24m who has muac_115_119 AND sev_uwt
  df <- df %>%
    mutate(
      # Check if both inputs are available in a row
      mod_muac_suwt_24m = case_when(
        is.na(sev_uwt_24m) | is.na(muac_115_119_24m) ~ NA_real_,  # Either input missing
        sev_uwt_24m == 1 & muac_115_119_24m == 1 ~ 1,             # Both are 1
        TRUE ~ 0                                                  # Otherwise 0
      )
   ) %>%
   set_value_labels(mod_muac_suwt_24m = c("Yes" = 1, "No" = 0)) %>%
   set_variable_labels(mod_muac_suwt_24m = "MUAC 115-119 & Severe UWT in children under 24M")
 
  #  Oedema
  # if oedema is not recoded, check oedema_original
  df <- df %>%
    mutate(oedema = 
             case_when(
               oedema == "Oui" ~ 1,
               oedema == "Yes" ~ 1,
               oedema == "yes" ~ 1,
               oedema == "Y" ~ 1,
               oedema == "y" ~ 1,
               oedema == "1" ~ 1,
               oedema == "n" ~ 0,
               oedema == "N" ~ 0,
               oedema == "non" ~ 0,
               oedema == "Non" ~ 0,
               oedema == "0" ~ 0,
               is.na(oedema) ~ NA_real_  # handle missing values
             )) %>%
    set_value_labels(oedema = c("Yes" = 1, "No" = 0)) %>%
    set_variable_labels(oedema = "bilateral oedema")
  
  # At Risk Combined (0-5 months)
  df <- df %>%
    mutate(
      valid_inputs = rowSums(!is.na(across(c(uwt, wast, muac_110, oedema)))),
      at_risk = case_when(
        valid_inputs == 0 ~ NA_real_,
        rowSums(across(c(uwt, wast, muac_110, oedema)) == 1, na.rm = TRUE) > 0 ~ 1,
        TRUE ~ 0
      )
    ) %>%
    set_value_labels(at_risk = c("Yes" = 1, "No" = 0)) %>%
    set_variable_labels(at_risk = "At Risk Combined")
  
  # Severe Acute Malnutrition (6-59 months)
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
  
  # View valid N of all indicators
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
  # View(valid_n_table)
  
  
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
  
  df_0_5m <- df %>% filter(agemons >=0 & agemons <6)
  
  # AT RISK TABLE
  # WHO Guideline 2024 - Indicators for at risk and acute malnutrition
  table_name <- "at_risk_0_5m"
  df_name <- "df_0_5m"
  indicators <- c("wast", "muac_110",  "oedema", "uwt", "at_risk")
  
  main_table <- get(df_name) %>%
    group_by(Region) %>%
    summarise_prev_table()
  
  total_row <- summarise_prev_table(get(df_name)) %>%
    mutate(Region = "Total") %>%
    select(Region, everything())
  
  full_table <- bind_rows(main_table, total_row)
  
  # Suppress % if N < 30
  for (var in indicators) {
    pct_col <- paste0(var, " (%)")
    n_col   <- paste0(var, " (N)")
    
    if (pct_col %in% names(full_table) && n_col %in% names(full_table)) {
      mask <- is.na(full_table[[n_col]]) | full_table[[n_col]] < 30
      if (any(mask)) {
        full_table[[pct_col]] <- as.character(full_table[[pct_col]])
        full_table[[pct_col]][mask] <- " - "
      }
    }
  }
  full_table <- replace_names_with_labels(full_table, get(df_name), indicators)
  assign(table_name, full_table)
  # View(get(table_name))
  
  # **************************************************************************************************
  # * Anthropometric indicators for children from 6- 59 months
  # **************************************************************************************************
  
  df_6_59m <- df %>% filter(agemons > 5 & agemons < 60)
  
  # SAM TABLE
  table_name <- "sam_6_59m"
  df_name <- "df_6_59m"
  indicators <- c("sev_wast", "muac_115", "oedema", "sam")
  
  main_table <- get(df_name) %>%
    group_by(Region) %>%
    summarise_prev_table()

  total_row <- summarise_prev_table(get(df_name)) %>%
    mutate(Region = "Total") %>%
    select(Region, everything())
  
  full_table <- bind_rows(main_table, total_row)
  
  # Suppress % if N < 30
  for (var in indicators) {
    pct_col <- paste0(var, " (%)")
    n_col   <- paste0(var, " (N)")
    
    if (pct_col %in% names(full_table) && n_col %in% names(full_table)) {
      mask <- is.na(full_table[[n_col]]) | full_table[[n_col]] < 30
      if (any(mask)) {
        full_table[[pct_col]] <- as.character(full_table[[pct_col]])
        full_table[[pct_col]][mask] <- " - "
      }
    }
  }
  full_table <- replace_names_with_labels(full_table, get(df_name), indicators)
  assign(table_name, full_table)
  # View(get(table_name))
  
  # GAM TABLE
  table_name <- "gam_6_59m"
  df_name <- "df_6_59m"
    indicators <- c("wast", "muac_125", "oedema", "gam")

  main_table <- get(df_name) %>%
    group_by(Region) %>%
    summarise_prev_table()
  
  total_row <- summarise_prev_table(get(df_name)) %>%
    mutate(Region = "Total") %>%
    select(Region, everything())
  
  full_table <- bind_rows(main_table, total_row)
  
  # Suppress % if N < 30
  for (var in indicators) {
    pct_col <- paste0(var, " (%)")
    n_col   <- paste0(var, " (N)")
    
    if (pct_col %in% names(full_table) && n_col %in% names(full_table)) {
      mask <- is.na(full_table[[n_col]]) | full_table[[n_col]] < 30
      if (any(mask)) {
        full_table[[pct_col]] <- as.character(full_table[[pct_col]])
        full_table[[pct_col]][mask] <- " - "
      }
    }
  }
  full_table <- replace_names_with_labels(full_table, get(df_name), indicators)
  assign(table_name, full_table)
  # View(get(table_name))

  # **************************************************************************************************
  # * Anthropometric indicators for children from 0- 59 months
  # **************************************************************************************************
  
  # mod_wast_0_59m TABLE   
  table_name <- "mod_wast_0_59m"
  df_name <- "df"  # Use full dataset of children 0-59M
  indicators <- c("muac_115_119", "sev_uwt", "mod_muac_suwt" ,"muac_115_119_24m", "sev_uwt_24m", "mod_muac_suwt_24m")
  
  main_table <- get(df_name) %>%
    group_by(Region) %>%
    summarise_prev_table()
  
  total_row <- summarise_prev_table(get(df_name)) %>%
    mutate(Region = "Total") %>%
    select(Region, everything())
  
  full_table <- bind_rows(main_table, total_row)
  
  # Suppress % if N < 30 - updated - convert to function and place above
  for (var in indicators) {
    pct_col <- paste0(var, " (%)")
    n_col   <- paste0(var, " (N)")
    
    if (pct_col %in% names(full_table) && n_col %in% names(full_table)) {
      
      # Create mask for low sample size OR variable is all NA
      all_na <- all(is.na(df[[var]]))  # Check the original variable in df
      mask <- is.na(full_table[[n_col]]) | full_table[[n_col]] < 30 | all_na
      
      if (any(mask)) {
        full_table[[pct_col]] <- as.character(full_table[[pct_col]])
        full_table[[pct_col]][mask] <- " - "
      }
    }
  }
  full_table <- replace_names_with_labels(full_table, get(df_name), indicators)
  assign(table_name, full_table)
  # View(get(table_name))
  
  
  # Children 0-59m with Wasting, MUAC, Underweight and Bilateral Oedema (all_0_59m)
  table_name <- "all_0_59m"
  df_name <- "df"  # Use full dataset of children 0-59M
  indicators <- c("wast", "muac_125","oedema", "uwt" )
  
  main_table <- get(df_name) %>%
    group_by(Region) %>%
    summarise_prev_table()
  
  total_row <- summarise_prev_table(get(df_name)) %>%
    mutate(Region = "Total") %>%
    select(Region, everything())
  
  full_table <- bind_rows(main_table, total_row)
  
  # Suppress % if N < 30
  for (var in indicators) {
    pct_col <- paste0(var, " (%)")
    n_col   <- paste0(var, " (N)")
    
    if (pct_col %in% names(full_table) && n_col %in% names(full_table)) {
      mask <- is.na(full_table[[n_col]]) | full_table[[n_col]] < 30
      if (any(mask)) {
        full_table[[pct_col]] <- as.character(full_table[[pct_col]])
        full_table[[pct_col]][mask] <- " - "
      }
    }
  }
  full_table <- replace_names_with_labels(full_table, get(df_name), indicators)
  assign(table_name, full_table)
  # View(get(table_name))
  
  
  # to label all tabs - use cleaned name - Remove everything before the first dash and after -ANT.csv
  cleaned_name <- sub("^[^-]+-", "", file)              # Remove before first dash
  cleaned_name <- sub("-ANT\\.csv$", "", cleaned_name)  # Remove -ANT.csv at end
  print(cleaned_name)
  
  country_name <- df$country[!is.na(df$country)][1]
  survey_name  <- df$survey[!is.na(df$survey)][1]
  survey_year  <- df$year[!is.na(df$year)][1]
  
  # Correct date representation
  start <- format(min(df$date_measure, na.rm = TRUE), "%d-%b-%Y")
  end <- format(max(df$date_measure, na.rm = TRUE), "%d-%b-%Y")
  

  
  # Define file, sheet, and cell position
  file_name <- file.path(outputdir, paste0("WHO_indicators_", country_name, ".xlsx"))
  tab_name <- cleaned_name # use 

  if (!file.exists(file_name)) {
    wb <- createWorkbook()
  } else {
    wb <- loadWorkbook(file_name)
    
    if (tab_name %in% names(wb)) {
      removeWorksheet(wb, tab_name)  # drop if not clean
    }
  }
  addWorksheet(wb, tab_name)
  
  x = 2
  y = 5
  y_note = length(unique(df$Region)) 
  add_y = y_note + 5  # add rows between each pasted table
  
  # Create a cell style with wrap text
  wrap_style <- createStyle(wrapText = TRUE, halign = "center")
  
  # Write Country, Survey Type, Start and End Date
  note1 <- paste("Country:", country_name,"   Survey:", survey_name, survey_year)
  note2 <- paste("Survey data collection from", start, "to", end)
  note_u6M <- paste("Note: All N are unweighted cases. Estimates with <30 unweighted cases are presented as '-'.  MUAC was often not collected for children <6M.")
  note3 <- paste("Note: All N are unweighted cases. Estimates with <30 unweighted cases are presented as '-'.")
  
  writeData(wb, sheet = tab_name, x = "WHO Guideline 2024 - Indicators for at risk and acute malnutrition", startCol = 2, startRow = 1)
  writeData(wb, sheet = tab_name, x = note1, startCol = 2, startRow = 2)
  writeData(wb, sheet = tab_name, x = note2, startCol = 2, startRow = 3)
  
  writeData(wb, sheet = tab_name, x = "Infants from 0-5m at risk of poor growth and development", startCol = x, startRow = y)
  writeData(wb, sheet = tab_name, x = at_risk_0_5m, startCol = x, startRow = y+1)
  writeData(wb, sheet = tab_name, x = note_u6M, startCol = 2, startRow = y + y_note + 3)
  addStyle(wb, sheet = tab_name, style = wrap_style, cols = 3:14, rows = 6, gridExpand = TRUE)
  y = y + add_y
  
  writeData(wb, sheet = tab_name, x = "Children 6-59m with Severe Acute Malnutrition", startCol = x, startRow = y)
  writeData(wb, sheet = tab_name, x = sam_6_59m, startCol = x, startRow = y+1)
  writeData(wb, sheet = tab_name, x = note3, startCol = 2, startRow = y + y_note + 3)
  addStyle(wb, sheet = tab_name, style = wrap_style, cols = 3:14, rows = y+1, gridExpand = TRUE)
  y = y + add_y
  
  writeData(wb, sheet = tab_name, x = "Children 6-59m with Global Acute Malnutrition", startCol = x, startRow = y)
  writeData(wb, sheet = tab_name, x = gam_6_59m, startCol = x, startRow = y+1)
  writeData(wb, sheet = tab_name, x = note3, startCol = 2, startRow = y + y_note + 3)
  addStyle(wb, sheet = tab_name, style = wrap_style, cols = 3:14, rows = y+1, gridExpand = TRUE)
  y = y + add_y
  
  writeData(wb, sheet = tab_name, x = "Children 0-59m and 0-23m with Moderate Wasting", startCol = x, startRow = y)
  writeData(wb, sheet = tab_name, x = mod_wast_0_59m, startCol = x, startRow = y+1)
  writeData(wb, sheet = tab_name, x = note3, startCol = 2, startRow = y + y_note + 3)
  addStyle(wb, sheet = tab_name, style = wrap_style, cols = 3:14, rows = y+1, gridExpand = TRUE)
  y = y + add_y
  
  writeData(wb, sheet = tab_name, x = "Children 0-59m with Wasting, MUAC, Underweight and Bilateral Oedema", startCol = x, startRow = y)
  writeData(wb, sheet = tab_name, x = all_0_59m, startCol = x, startRow = y+1)
  writeData(wb, sheet = tab_name, x = note3, startCol = 2, startRow = y + y_note + 3)
  addStyle(wb, sheet = tab_name, style = wrap_style, cols = 3:14, rows = y+1, gridExpand = TRUE)
  
  
  # add graphs
  
  df <- df %>% mutate(agemons_complete = floor(agemons))

  # Create 'agemons_complete' and calculate weighted means by month
  df_summary <- df %>%
    ungroup() %>%
    filter(!is.na(sample_wgt)) %>%
    group_by(agemons_complete) %>%
    summarise(
      across(
        .cols = c(wast, sev_wast, uwt, muac_110, muac_115, oedema, at_risk, sam),
        .fns = list(
          N = ~ sum(!is.na(.x)),
          mean = ~ ifelse(sum(!is.na(.x)) < 30, NA_real_,
                          100 * weighted.mean(.x, sample_wgt, na.rm = TRUE))
        ),
        .names = "{.col}_{.fn}"
      ),
      .groups = "drop"
    ) %>%
    rename(
      wasting = wast_mean,
      severe_wasting = sev_wast_mean,
      underweight = uwt_mean,
      muac_110mm = muac_110_mean,
      muac_115mm = muac_115_mean,
      oedema = oedema_mean, 
      at_risk_combined = at_risk_mean,
      sam_combined = sam_mean
    )
  
  # If muac is missing, replace with NA_real
  if (all(is.na(df$muac_115))) {
    df_summary$muac_115mm <- NA_real_
  }
  if (all(is.na(df$muac_110))) {
    df_summary$muac_110mm <- NA_real_
  }
  
  # graph of prevalence of infants at risk of poor growth and development (0-5M) by agemons
  
  plot_path <- file.path(tempdir(), paste0("at_risk_plot_", tab_name, ".png"))
  
  #  Pivot to long format for plotting
  df_long <- df_summary %>%
    pivot_longer(cols = c(wasting, underweight, muac_110mm, at_risk_combined),
                 names_to = "indicator",
                 values_to = "mean_value")
  
  # Set colors and line types
  indicator_colors <- c("underweight" = "#1e90ff", "wasting" = "#ff69b4", "muac_110mm" = "red", "at_risk_combined" = "purple" )
  indicator_linetypes <- c("underweight" = "solid", "wasting" = "solid" , "muac_110mm" = "dotted", "at_risk_combined" = "solid")
  
  # Create and save At Risk plot
  png(plot_path, width = 1200, height = 800, res = 150)
  at_risk_plot <-  ggplot(df_long, aes(x = agemons_complete, y = mean_value, color = indicator, linetype = indicator)) +
    # geom_point(alpha = 0.5) +
    geom_smooth(method = "loess", span = 0.9, se = FALSE, size = 1) +
    scale_color_manual(values = indicator_colors) +
    scale_linetype_manual(values = indicator_linetypes) +
    scale_y_continuous(limits = c(0, NA)) +                # Start y-axis at 0
    scale_x_continuous(limits = c(0,5)) +  # Set x-axis to 0 - 5
    labs(
      title = "Prevalence of infants at risk of poor growth and development by age in months",
      x = "Age in Months",
      y = "Prevalence (%)",
      color = "Indicator",
      linetype = "Indicator",
      caption = "Note: If the planned data collection did not include MUAC measures for infants <6 months, the MUAC 115mm trend is not representative"
    ) +
    theme_minimal() +
    theme(plot.caption = element_text(hjust = 0))  # left-align caption
  print(at_risk_plot)
  dev.off()
  
  if (!file.exists(plot_path) || file.info(plot_path)$size == 0) {
    cat("At risk plot image was not saved: ", plot_path)
  } else {
    cat(" Inserting At Risk Plot for:", tab_name, "\n")
      insertImage(
      wb,
      sheet = tab_name,
      file = plot_path,
      startRow = 2,
      startCol = 16,
      width = 8,
      height = 5.33,
      units = "in"
    )
  }
  
  # graph of severe acute malnutrition by age from 0-59m
  
  plot_path <- file.path(tempdir(), paste0("sam_plot_", tab_name, ".png"))
  
  #  Pivot to long format for plotting
  df_long <- df_summary %>%
    pivot_longer(cols = c(severe_wasting, muac_115mm, oedema, sam_combined),
                 names_to = "indicator",
                 values_to = "mean_value")
  
  # Set colors and line types
  indicator_colors <- c("severe_wasting" = "red", "muac_115mm" = "red", oedema = "black", "sam_combined" = "darkred" )
  indicator_linetypes <- c("severe_wasting" = "solid" , "muac_115mm" = "dotted", "oedema" = "solid", "sam_combined" = "solid")
  
  png(plot_path, width = 1200, height = 800, res = 150)
  sam_plot <-ggplot(df_long, aes(x = agemons_complete, y = mean_value, color = indicator, linetype = indicator)) +
    # geom_point(alpha = 0.5) +
    geom_smooth(method = "loess", span = 0.75, se = TRUE, size = 1) +
    scale_color_manual(values = indicator_colors) +
    scale_linetype_manual(values = indicator_linetypes) +
    scale_y_continuous(
      breaks = seq(0, max(df_long$mean_value, na.rm = TRUE), by = 10),
      expand = expansion(mult = c(0, 0.05))) +
    scale_x_continuous(breaks = seq(0, max(df_long$agemons_complete, na.rm = TRUE), by = 6)) +  # Set x-axis to 0, 6, 12, ...
    labs(
      title = "Prevalence of severe acute malnutrition by age in months",
      x = "Age in Months",
      y = "Prevalence (%)",
      color = "Indicator",
      linetype = "Indicator"
    ) +
    theme_minimal()
  
  print(sam_plot)
  dev.off()
  
  if (!file.exists(plot_path) || file.info(plot_path)$size == 0) {
    cat("SAM plot image was not saved: ", plot_path)
  } else {
      cat(" Inserting SAM Plot for:", tab_name, "\n")
      insertImage(
        wb,
        sheet = tab_name,
        file = plot_path,
        startRow = 26,
        startCol = 16,
        width = 8,
        height = 5.33,
        units = "in"
      )
  }
  

  
  
  # WHZ Plot
  if ("whz" %in% names(df) && any(!is.na(df$whz))) { # if whz exists or at least one non missing - continue
    df_clean <- df %>% filter(!is.na(whz), !is.na(Region))
    
    plot_path <- file.path(tempdir(), paste0("whz_plot_", tab_name, ".png"))
    
    # Create and save WHZ plot
    png(plot_path, width = 1200, height = 800, res = 150)
    whz_plot <- ggplot(df_clean, aes(x = whz)) +
      geom_histogram(aes(y = ..density.. * 100), binwidth = 0.2, fill = "skyblue", color = "white") +
      stat_function(fun = function(x) dnorm(x, mean = 0, sd = 1) * 100,
                    color = "gray", linewidth = 1) +
      facet_wrap(~ Region) +
      scale_y_continuous(name = "Percent Density") +
      scale_x_continuous(name = "WHZ", breaks = seq(-5, 5, by = 1)) +
      labs(title = "WHZ Percent Distribution in children 0–59M by Region") +
      theme_minimal() +
      theme(axis.text.x = element_text(angle = 0, hjust = 1))
    print(whz_plot)
    dev.off()
    
    if (!file.exists(plot_path) || file.info(plot_path)$size == 0) {
      cat(" WHZ plot image was not saved: ", plot_path)
    }
    
    cat(" Inserting WHZ plot for:", tab_name, "\n")
    
    insertImage(
      wb,
      sheet = tab_name,
      file = plot_path,
      startRow = 52,
      startCol = 16,
      width = 8,
      height = 5.33,
      units = "in"
    )
  } else {
    cat("️ No valid 'whz' data found in:", tab_name, "\n")
  }
  
  # MUAC plot
  if ("muac" %in% names(df) && any(!is.na(df$muac))) {  # if muac exists or at least one non missing - continue
    df_clean <- df %>% filter(!is.na(muac), !is.na(Region))
    
    plot_path <- file.path(tempdir(), paste0("muac_plot_", tab_name, ".png"))
    
    # Calculate limits for x axis label
    muac_min <- max(60, floor(min(df$muac, na.rm = TRUE) / 10) * 10)  # 60 is min
    muac_max <- min(260, floor(max(df$muac, na.rm = TRUE) / 10) * 10)  # 260 is max
    
    # Create and save MUAC plot
    png(plot_path, width = 1200, height = 800, res = 150)
    muac_plot <- ggplot(df, aes(x = muac)) +
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
      theme(axis.text.x = element_text(angle = 90, hjust = 1)) +
      coord_cartesian(expand = FALSE)
    
    print(muac_plot)
    dev.off()
    
    if (!file.exists(plot_path) || file.info(plot_path)$size == 0) {
      cat(" MUAC plot image was not saved: ", plot_path)
    }
    cat(" Inserting MUAC plot for:", tab_name, "\n")
    
    insertImage(
      wb,
      sheet = tab_name,
      file = plot_path,
      startRow = 75,
      startCol = 16,   # Plot below WHZ graph
      width = 8,
      height = 5.33,
      units = "in"
    )
  } else {
    cat("️ No valid 'MUAC' data found in:", tab_name, "\n")
  }
  
  # Save workbook
  saveWorkbook(wb, file_name, overwrite = TRUE)
  cat(" Saved Excel file to:", file_name, "\n")
  
}



