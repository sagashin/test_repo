# 1st opinionと2nd オピニオンの比較表を作る。Accounting Reserveも必要であれば作る
# by LoB
#   MotorはCoverage別
#   その他はHealth、Pet
# CYとPYにSplitする
# Run 2nd opinion program before running as reading 2nd opinion exhibits

#
#
current_tri_wd <- file.path(mainDir, subDir_2)#"P:/AXADJDivision/RiskManagement/05.Insurance（保険引受リスク）/43.リザービング/2021/Second Opinion/202103/01_Triangle_Data"
current_work_wd <- file.path(mainDir, subDir_5)#"P:/AXADJDivision/RiskManagement/05.Insurance（保険引受リスク）/43.リザービング/2021/Second Opinion/202103/99 work"
previous_wd <- "C:/Users/shin.sagara/OneDrive - AXA/02. 保険引受リスク（P&C)/43.リザービング/2023/PCFY/99 work" #WD for previous summary file
#
#

library(haven)
library(dplyr)
library(tidyverse)
library(data.table)
library(reshape2)
library(ggplot2)
library(directlabels)
library(GGally)
library(readxl)
library(pdftools)
library(R6)
library(ChainLadder)
library(janitor)
library(rowr)
library(knitr)
library(formattable)
library(kableExtra)
library(qpcR)
library(DT)
library(bit64)
library(flexdashboard)
library(stringr)
library(reshape2)
library("writexl")
library(res2ndopi)
#
# set WD and file name ----------------------------------------------
#
setwd(current_tri_wd)
#fst_motor_filename <- "1st_opinion_data (Motor)_LinkOut.xlsx" Not available yet 20200608
#fst_auto_filename <- "01_IBNR_Auto_new_Linkout202012Finance_v4.xlsx"
#fst_bike_filename <- "01_IBNR_Bike_new_Linkout202012Finance_v4.xlsx"
#fst_pet_filename <- "01_IBNR_Pet_new_Linkout202012Finance.xlsx"
#fst_health_filename <- "01_IBNR_Health_new_Linkout202012Finance.xlsx"
prev_summary_filename <- "Sep23_summary_1st_2nd_v2.xlsm"
#
# ------------------------------------------------------------
#
# define functions -------------------------------------------
#
# 1stのEAXAやリザーブを取得  -------------------------------------------------
#
#origin_date <- read_excel("origin_ym_199906_202012.xlsx") # 基ファイルはエクセルでアップデート。AYを直近までのばす
#
# Function for reading excel file for 1st opinion
# For Auto, Bike, and Health (Petは別下記Function)
#
get_1st_opinion <- function(filename, sheetname, excl_row1, excl_row2, excl_row3, biz, prod, covname) {

  df <- read_excel(filename, sheet = sheetname, range = cell_cols("A:AA"))
  df <- df[-c(1:excl_row1, excl_row2:(dim(df)[1])), ] # mod excl_row3 is not necessary anymore
  #df <- df[-c(1:excl_row1, excl_row2:excl_row3), ]
  df <- cbind(origin_date, df) #reading origin_data from outside
  df$BIZ_LINE <- biz
  df$PRODUCTS <- prod
  df$COV_NAME <- sheetname
  df$YEAR <- year(origin_date[[1]])
  colnames(df)[10] <- "EXPOSURE"
  colnames(df)[11] <-"EARNINGS"
  colnames(df)[13] <-"CLCNT_PAID"
  colnames(df)[14] <-"CLCNT_OS"
  colnames(df)[15] <-"CLCNT_INC"
  colnames(df)[16] <-"CLCNT_LDF"
  colnames(df)[17] <-"CLCNT_LDF_SEL"
  colnames(df)[19] <-"CLCNT_EAXA"
  colnames(df)[20] <-"CLCNT_IBNR"
  colnames(df)[21] <-"PAID"
  colnames(df)[22] <-"OS"
  colnames(df)[23] <-"INC"
  colnames(df)[24] <-"LOSS_LDF"
  colnames(df)[25] <-"LOSS_LDF_SEL"
  colnames(df)[26] <-"TARGET_LR"
  colnames(df)[27] <-"EAXA"
  colnames(df)[28] <-"IBNR"

  df$EXPOSURE <- as.numeric(df['EXPOSURE'][[1]])
  df$EARNINGS <- as.numeric(df['EARNINGS'][[1]])
  df$CLCNT_PAID <- as.numeric(df['CLCNT_PAID'][[1]])
  df$CLCNT_OS <- as.numeric(df['CLCNT_OS'][[1]])
  df$CLCNT_INC <- as.numeric(df['CLCNT_INC'][[1]])
  df$CLCNT_LDF <- as.numeric(df['CLCNT_LDF'][[1]])
  df$CLCNT_LDF_SEL <- as.numeric(df['CLCNT_LDF_SEL'][[1]])
  df$CLCNT_EAXA <- as.numeric(df['CLCNT_EAXA'][[1]])
  df$CLCNT_IBNR <- as.numeric(df['CLCNT_IBNR'][[1]])
  df$PAID <- as.numeric(df['PAID'][[1]])
  df$OS <- as.numeric(df['OS'][[1]])
  df$INC <- as.numeric(df['INC'][[1]])
  df$LOSS_LDF <- as.numeric(df['LOSS_LDF'][[1]])
  df$LOSS_LDF_SEL <- as.numeric(df['LOSS_LDF_SEL'][[1]])
  df$TARGET_LR <- as.numeric(df['TARGET_LR'][[1]])
  df$EAXA <- as.numeric(df['EAXA'][[1]])
  df$IBNR <- as.numeric(df['IBNR'][[1]])
  df$EAXA_RES <- as.numeric(df['OS'][[1]]) + as.numeric(df['IBNR'][[1]])
  df <- df[, c(1,32,c(29:31),10,11,c(13:17,19:28),33)]
  return(df)
}
#
# For Pet
#

get_1st_opinion_pet <- function(filename, sheetname, excl_row1, excl_row2, excl_row3, biz, prod, covname) {

  df <- read_excel(filename, sheet = sheetname, range = cell_cols("A:AD")) #Petはエクセルのフォーマットが異なるので、読み込み範囲が異なるので注意
  df <- df[-c(1:excl_row1, excl_row2:excl_row3), ]
  df <- cbind(origin_date, df) #reading origin_data from outside
  df$BIZ_LINE <- biz
  df$PRODUCTS <- prod
  df$COV_NAME <- sheetname
  df$YEAR <- year(origin_date[[1]])
  colnames(df)[10] <- "EXPOSURE"
  colnames(df)[11] <-"EARNINGS"
  colnames(df)[13] <-"CLCNT_PAID"
  colnames(df)[14] <-"CLCNT_OS"
  colnames(df)[15] <-"CLCNT_INC"
  colnames(df)[16] <-"CLCNT_LDF"
  colnames(df)[17] <-"CLCNT_LDF_SEL"
  colnames(df)[19] <-"CLCNT_EAXA"
  colnames(df)[20] <-"CLCNT_IBNR"
  colnames(df)[21] <-"PAID"
  colnames(df)[22] <-"OS"
  colnames(df)[23] <-"INC"
  colnames(df)[24] <-"LOSS_LDF"
  colnames(df)[25] <-"LOSS_LDF_SEL"
  colnames(df)[27] <-"TARGET_LR"
  colnames(df)[30] <-"EAXA"
  colnames(df)[31] <-"IBNR"

  df$EXPOSURE <- as.numeric(df['EXPOSURE'][[1]])
  df$EARNINGS <- as.numeric(df['EARNINGS'][[1]])
  df$CLCNT_PAID <- as.numeric(df['CLCNT_PAID'][[1]])
  df$CLCNT_OS <- as.numeric(df['CLCNT_OS'][[1]])
  df$CLCNT_INC <- as.numeric(df['CLCNT_INC'][[1]])
  df$CLCNT_LDF <- as.numeric(df['CLCNT_LDF'][[1]])
  df$CLCNT_LDF_SEL <- as.numeric(df['CLCNT_LDF_SEL'][[1]])
  df$CLCNT_EAXA <- as.numeric(df['CLCNT_EAXA'][[1]])
  df$CLCNT_IBNR <- as.numeric(df['CLCNT_IBNR'][[1]])
  df$PAID <- as.numeric(df['PAID'][[1]])
  df$OS <- as.numeric(df['OS'][[1]])
  df$INC <- as.numeric(df['INC'][[1]])
  df$LOSS_LDF <- as.numeric(df['LOSS_LDF'][[1]])
  df$LOSS_LDF_SEL <- as.numeric(df['LOSS_LDF_SEL'][[1]])
  df$TARGET_LR <- as.numeric(df['TARGET_LR'][[1]])
  df$EAXA <- as.numeric(df['EAXA'][[1]])
  df$IBNR <- as.numeric(df['EAXA'][[1]]) - as.numeric(df['INC'][[1]])
  df$EAXA_RES <- as.numeric(df['OS'][[1]]) + as.numeric(df['IBNR'][[1]])
  #df <- df[, c(1,35,c(32:34),10,11,c(13:17,19:25,27,28,31),36)]
  return(df)
}
#
# Function for Summarise by AY and make tables showing 1st and 2nd together ---------------
#
summarise_1st_opinion <- function(df, sec_exhibit, prev_sheetname) {
  df_summary <- df %>%
    group_by(YEAR) %>%
    summarise(
      EXPOSURE = sum(EXPOSURE),
      EARNINGS = sum(EARNINGS),
      CLCNT_PAID = sum(CLCNT_PAID),
      CLCNT_OS = sum(CLCNT_OS),
      CLCNT_INC = sum(CLCNT_INC),
      FST_CLCNT_EAXA = accounting(sum(CLCNT_EAXA), 0),
      FST_CLCNT_IBNR = sum(CLCNT_IBNR),
      PAID = sum(PAID),
      OS = sum(OS),
      INC = sum(INC),
      FST_EAXA = accounting(sum(EAXA), 0),
      FST_IBNR = accounting(sum(IBNR), 0),
      FST_EAXA_RES = accounting(sum(EAXA_RES), 0)
    )
  tot <- df_summary %>%
    summarize_if(is.numeric, sum, na.rm=TRUE) %>%
    mutate(YEAR="Total")
  cy <- df_summary %>%
    dplyr::filter(YEAR == max(YEAR)) %>%
    summarize_if(is.numeric, sum, na.rm=TRUE) %>%
    mutate(YEAR="CY_Total")
  py <- df_summary %>%
    dplyr::filter(YEAR < max(YEAR)) %>%
    summarize_if(is.numeric, sum, na.rm=TRUE) %>%
    mutate(YEAR="PY_Total")
  df_summary <- rbind(df_summary, py, cy, tot)
  df_summary$FST_EAXA_LR <- percent(df_summary$FST_EAXA / df_summary$EARNINGS, digits = 3, format = "f")
  df_summary$FST_EAXA_FREQ <- percent(df_summary$FST_CLCNT_EAXA / df_summary$EXPOSURE, digits = 3, format = "f")
  df_summary$FST_EAXA_SEV <- accounting(df_summary$FST_EAXA / df_summary$FST_CLCNT_EAXA, 0)
  df_summary <- cbind(sec_exhibit, df_summary[c("FST_CLCNT_EAXA", "FST_EAXA", "FST_IBNR", "FST_EAXA_RES", "FST_EAXA_LR", "FST_EAXA_FREQ", "FST_EAXA_SEV")])
  df_summary$EAXA_RES_DIF <- accounting(df_summary$EAXA_RES - df_summary$FST_EAXA_RES, 0)
  df_summary$EAXA_RES_DIF_PCT <-  ifelse(df_summary$FST_EAXA_RES == 0, 0, percent(df_summary$EAXA_RES / df_summary$FST_EAXA_RES - 1, digits = 3, format = "f"))
  df_summary$EAXA_LR_DIF <- percent(df_summary$EAXA_LR - df_summary$FST_EAXA_LR, digits = 3, format = "f")
  #
  # Read previous summary
  #
  setwd(previous_wd)
  prev <- read_excel(prev_summary_filename, sheet = prev_sheetname)
  prev <- prev %>%
    dplyr::select(AY, INCURRED, EAXA, EAXA_RES, FF_LR, EAXA_LR, FST_EAXA, FST_EAXA_RES, FST_EAXA_LR) %>%
    dplyr::rename(Prev_INCURRED = INCURRED, Prev_EAXA = EAXA, Prev_EAXA_RES = EAXA_RES, Prev_FF_LR = FF_LR, Prev_EAXA_LR = EAXA_LR,
                  Prev_FST_EAXA = FST_EAXA, Prev_FST_EAXA_RES = FST_EAXA_RES, Prev_FST_EAXA_LR = FST_EAXA_LR)
  # merge current and previous summary table
  df_summary <- merge(x = df_summary, y = prev, by = "AY", all.x = TRUE)
  df_summary$FF_BM <- df_summary$Prev_INCURRED - df_summary$INCURRED
  df_summary$SEC_EAXA_BM <- df_summary$Prev_EAXA - df_summary$EAXA
  df_summary$FST_EAXA_BM <- df_summary$Prev_FST_EAXA - df_summary$FST_EAXA
  df_summary <- df_summary[c(1:(dim(df_summary)[1]-3), (dim(df_summary)[1]-1), (dim(df_summary)[1]-2), (dim(df_summary)[1])),]
  df_summary$cov <- substring(prev_sheetname, 5)


  #df_summary2 <- setDT(as.data.frame(t(df_summary)), keep.rownames = TRUE)[]
  #colnames(df_summary2) <- c("AY", as.character(df_summary$AY))
  #df_summary <- df_summary2[-1, ]

  return(df_summary)
}
#
# PY CY summary for Flash report ------------------
#
get_1st_py_cy <- function(input_data, colname) {
  # This function summarizes 1st opinion by PY and CY

  # EAXA Reserve
  amt <- as.numeric(as.matrix(input_data[colname]))
  # Previous Year
  amt_py <- amt[len(amt)-2]
  # Current Year
  amt_cy <- amt[len(amt)-1]
  # column bind PY and CY
  amt_py_cy <- cbind(amt_py, amt_cy)
  return(amt_py_cy)
}
#
# ----------------------------------------------------------
#
#
# Run get_1st_opinion function to get 1st opinion
#
# Auto
fst_auto_total <- get_1st_opinion(fst_auto_filename, sheetname="TOTAL", excl_row1=8, excl_row2=excl_rows2_input_health_and_total, biz="Motor", prod="AUTO")
fst_auto_bil_sc <- get_1st_opinion(fst_auto_filename, sheetname="BIL_attr", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=387, biz="Motor", prod="AUTO")
fst_auto_bil_lc <- get_1st_opinion(fst_auto_filename, sheetname="BIL_large", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=387, biz="Motor", prod="AUTO")
fst_auto_bil <- cbind(fst_auto_bil_sc[, c(1:7)], fst_auto_bil_sc[, -c(1:7)] + fst_auto_bil_lc[, -c(1:7)]) # adding sc and lc EPまで共通
fst_auto_pdl_sc <- get_1st_opinion(fst_auto_filename, sheetname="PDL_attr", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=387, biz="Motor", prod="AUTO")
fst_auto_pdl_lc <- get_1st_opinion(fst_auto_filename, sheetname="PDL_large", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=387, biz="Motor", prod="AUTO")
fst_auto_pdl <- cbind(fst_auto_pdl_sc[, c(1:7)], fst_auto_pdl_sc[, -c(1:7)] + fst_auto_pdl_lc[, -c(1:7)]) # adding sc and lc EPまで共通
fst_auto_ppa <- get_1st_opinion(fst_auto_filename, sheetname="PPA_attr", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=387, biz="Motor", prod="AUTO")
fst_auto_od_attr <- get_1st_opinion(fst_auto_filename, sheetname="OD_attr", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=387, biz="Motor", prod="AUTO")
fst_auto_od_nat <- get_1st_opinion(fst_auto_filename, sheetname="OD_nat", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=387, biz="Motor", prod="AUTO")
fst_auto_pe_attr <- get_1st_opinion(fst_auto_filename, sheetname="PE_attr", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=387, biz="Motor", prod="AUTO")
fst_auto_pe_nat <- get_1st_opinion(fst_auto_filename, sheetname="PE_nat", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=387, biz="Motor", prod="AUTO")
fst_auto_odpe_attr <- cbind(fst_auto_od_attr[, c(1:6)], fst_auto_od_attr[, -c(1:6)] + fst_auto_pe_attr[, -c(1:6)]) # adding od and pe Exposureまで共通
fst_auto_odpe_nat <- cbind(fst_auto_od_nat[, c(1:6)], fst_auto_od_nat[, -c(1:6)] + fst_auto_pe_nat[, -c(1:6)]) # adding od and pe Exposureまで共通
# ODPE NAT total
fst_auto_odpe_attr_nat <- cbind(fst_auto_odpe_attr[, c(1:7)], fst_auto_odpe_attr[, -c(1:7)] + fst_auto_odpe_nat[, -c(1:7)]) # adding nat EPまで共通
fst_auto_pi_sc <- get_1st_opinion(fst_auto_filename, sheetname="PI_attr", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=387, biz="Motor", prod="AUTO")
fst_auto_pi_lc <- get_1st_opinion(fst_auto_filename, sheetname="PI_large", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=387, biz="Motor", prod="AUTO")
fst_auto_pi <- cbind(fst_auto_pi_sc[, c(1:7)], fst_auto_pi_sc[, -c(1:7)] + fst_auto_pi_lc[, -c(1:7)]) # adding sc and lc EPまで共通
fst_auto_axap <- get_1st_opinion(fst_auto_filename, sheetname="AXA+_attr", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=387, biz="Motor", prod="AUTO")
fst_auto_fb <- get_1st_opinion(fst_auto_filename, sheetname="FB_attr", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=387, biz="Motor", prod="AUTO")
fst_auto_le <- get_1st_opinion(fst_auto_filename, sheetname="LX_attr", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=387, biz="Motor", prod="AUTO")
fst_auto_eq <- get_1st_opinion(fst_auto_filename, sheetname="EQ_attr", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=387, biz="Motor", prod="AUTO")
# Bike
fst_bike_total <- get_1st_opinion(fst_bike_filename, sheetname="TOTAL", excl_row1=8, excl_row2=excl_rows2_input_health_and_total, biz="Motor", prod="BIKE")
fst_bike_bil_sc <- get_1st_opinion(fst_bike_filename, sheetname="BIL_attr", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=387, biz="Motor", prod="BIKE")#[-c(1:71), ] #バイクは2005年以降
fst_bike_bil_lc <- get_1st_opinion(fst_bike_filename, sheetname="BIL_large", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=387, biz="Motor", prod="BIKE")#[-c(1:71), ]
fst_bike_bil <- cbind(fst_bike_bil_sc[, c(1:7)], fst_bike_bil_sc[, -c(1:7)] + fst_bike_bil_lc[, -c(1:7)]) # adding sc and lc EPまで共通
fst_bike_pdl <- get_1st_opinion(fst_bike_filename, sheetname="PDL_attr", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=387, biz="Motor", prod="BIKE")#[-c(1:71), ]
fst_bike_ppa <- get_1st_opinion(fst_bike_filename, sheetname="PPA_attr", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=387, biz="Motor", prod="BIKE")#[-c(1:71), ]
fst_bike_pi_sc <- get_1st_opinion(fst_bike_filename, sheetname="PI_attr", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=387, biz="Motor", prod="BIKE")#[-c(1:71), ]
fst_bike_pi_lc <- get_1st_opinion(fst_bike_filename, sheetname="PI_large", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=387, biz="Motor", prod="BIKE")#[-c(1:71), ]
fst_bike_pi <- cbind(fst_bike_pi_sc[, c(1:7)], fst_bike_pi_sc[, -c(1:7)] + fst_bike_pi_lc[, -c(1:7)]) # adding sc and lc EPまで共通
fst_bike_le <- get_1st_opinion(fst_bike_filename, sheetname="LX_attr", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=387, biz="Motor", prod="BIKE")#[-c(1:71), ]
# Motor
fst_motor_total <- cbind(fst_auto_total[, c(1:5)], fst_auto_total[, -c(1:5)] + fst_bike_total[, -c(1:5)]) # Auto + Bike
fst_motor_total$PRODUCTS <- "MOTOR"
fst_motor_bil <- cbind(fst_auto_bil[, c(1:5)], fst_auto_bil[, -c(1:5)] + fst_bike_bil[, -c(1:5)]) # Auto + Bike
fst_motor_pdl <- cbind(fst_auto_pdl[, c(1:5)], fst_auto_pdl[, -c(1:5)] + fst_bike_pdl[, -c(1:5)]) # Auto + Bike
fst_motor_ppa <- cbind(fst_auto_ppa[, c(1:5)], fst_auto_ppa[, -c(1:5)] + fst_bike_ppa[, -c(1:5)]) # Auto + Bike
fst_motor_pi <- cbind(fst_auto_pi[, c(1:5)], fst_auto_pi[, -c(1:5)] + fst_bike_pi[, -c(1:5)]) # Auto + Bike
fst_motor_le <- cbind(fst_auto_le[, c(1:5)], fst_auto_le[, -c(1:5)] + fst_bike_le[, -c(1:5)]) # Auto + Bike
# Health
fst_health_total <- get_1st_opinion(fst_health_filename, sheetname="PA_TOTAL", excl_row1=8, excl_row2=excl_rows2_input_health_and_total, biz="PA", prod="PA")
# Pet
fst_pet_total <- get_1st_opinion_pet(fst_pet_filename, sheetname="PET", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=10000, biz="PET", prod="PET")
#
#
# Run summarise_1st_opinion function
# Run 2nd opinion program before running below as reading 2nd opinion exhibits
#
#
# Auto
fst_auto_bil_sc_by_AY <- summarise_1st_opinion(fst_auto_bil_sc, auto_bil_sc_exhibit, "auto_bil_sc")
fst_auto_bil_lc_by_AY <- summarise_1st_opinion(fst_auto_bil_lc, auto_bil_lc_exhibit, "auto_bil_lc")
fst_auto_pdl_by_AY <- summarise_1st_opinion(fst_auto_pdl, auto_pdl_exhibit, "auto_pdl")
fst_auto_ppa_by_AY <- summarise_1st_opinion(fst_auto_ppa, auto_ppa_exhibit, "auto_ppa")
fst_auto_odpe_attr_by_AY <- summarise_1st_opinion(fst_auto_odpe_attr, auto_odpe_exhibit, "auto_odpe_attr")
fst_auto_odpe_nat_by_AY <- summarise_1st_opinion(fst_auto_odpe_nat, auto_odpe_ev_exhibit, "auto_odpe_nat")
fst_auto_odpe_attr_nat_by_AY <- summarise_1st_opinion(fst_auto_odpe_attr_nat, auto_od_pe_ev_exhibit, "auto_odpe_attr") # シート名は仮 ODPE attr + nat
fst_auto_pi_sc_by_AY <- summarise_1st_opinion(fst_auto_pi_sc, auto_pi_sc_exhibit, "auto_pi_sc")
fst_auto_pi_lc_by_AY <- summarise_1st_opinion(fst_auto_pi_lc, auto_pi_lc_exhibit, "auto_pi_lc")
fst_auto_axap_by_AY <- summarise_1st_opinion(fst_auto_axap, auto_axap_exhibit, "auto_axap")
fst_auto_fb_by_AY <- summarise_1st_opinion(fst_auto_fb, auto_fb_exhibit, "auto_fb")
fst_auto_le_by_AY <- summarise_1st_opinion(fst_auto_le, auto_le_exhibit, "auto_le")
fst_auto_eq_by_AY <- summarise_1st_opinion(fst_auto_eq, auto_eq_exhibit, "auto_eq")
#fst_auto_pdl_sc_by_AY <- summarise_1st_opinion(fst_auto_pdl_sc, auto_pdl_exhibit) # PDL 2nd opinionはLCとSCの区分なしなので不要
#fst_auto_pe_by_AY <- summarise_1st_opinion(fst_auto_pe) # 2nd opinionはODとPEを合算しているのでPEの区分なしで不要
# Bike
add_old_ay_to_bike <- function(bike_exhibit) {
  df <- eaxa_exhibit(c("1999", "2000", "2001", "2002", "2003", "2004"), bike_exhibit)
  return(df)
}
fst_bike_bil_sc_by_AY <- summarise_1st_opinion(fst_bike_bil_sc, add_old_ay_to_bike(bike_bil_sc_exhibit), "bike_bil_sc")
fst_bike_bil_lc_by_AY <- summarise_1st_opinion(fst_bike_bil_lc, add_old_ay_to_bike(bike_bil_lc_exhibit), "bike_bil_lc")
fst_bike_bil_by_ay <- summarise_1st_opinion(fst_bike_bil, add_old_ay_to_bike(bike_bil_exhibit), "bike_bil_lc") # シート名は仮
fst_bike_pdl_by_AY <- summarise_1st_opinion(fst_bike_pdl, add_old_ay_to_bike(bike_pdl_exhibit), "bike_pdl")
fst_bike_ppa_by_AY <- summarise_1st_opinion(fst_bike_ppa, add_old_ay_to_bike(bike_ppa_exhibit), "bike_ppa")
fst_bike_pi_sc_by_AY <- summarise_1st_opinion(fst_bike_pi_sc, add_old_ay_to_bike(bike_pi_sc_exhibit), "bike_pi_sc")
fst_bike_pi_lc_by_AY <- summarise_1st_opinion(fst_bike_pi_lc, add_old_ay_to_bike(bike_pi_lc_exhibit), "bike_pi_lc")
fst_bike_le_by_AY <- summarise_1st_opinion(fst_bike_le, add_old_ay_to_bike(bike_le_exhibit), "bike_le")
# Motor
fst_motor_total_by_AY <- summarise_1st_opinion(fst_motor_total, motor_exhibit, "motor_tot")
fst_motor_bil_by_AY <- summarise_1st_opinion(fst_motor_bil, motor_bil_exhibit, "motor_bil") 
fst_motor_pdl_by_AY <- summarise_1st_opinion(fst_motor_pdl, motor_pdl_exhibit, "motor_pdl")
fst_motor_ppa_by_AY <- summarise_1st_opinion(fst_motor_ppa, motor_ppa_exhibit, "motor_ppa")
fst_motor_pi_by_AY <- summarise_1st_opinion(fst_motor_pi, motor_pi_exhibit, "motor_pi")
fst_motor_le_by_AY <- summarise_1st_opinion(fst_motor_le, motor_le_exhibit, "motor_le")
# Health
fst_health_total_by_AY <- summarise_1st_opinion(fst_health_total, pa_exhibit, "health_total")
# Pet
fst_pet_total_by_AY <- summarise_1st_opinion(fst_pet_total, pet_exhibit, "pet_total")
#
# PY CY summary for Flash report ------------------
#
# Motor
# EAXA Reserve
fst_motor_bil <- get_1st_py_cy(fst_auto_bil_sc_by_AY, "FST_EAXA_RES") + get_1st_py_cy(fst_auto_bil_lc_by_AY, "FST_EAXA_RES") +
             get_1st_py_cy(fst_bike_bil_sc_by_AY, "FST_EAXA_RES") +　get_1st_py_cy(fst_bike_bil_lc_by_AY, "FST_EAXA_RES")
fst_motor_pdl <- get_1st_py_cy(fst_auto_pdl_by_AY, "FST_EAXA_RES") + get_1st_py_cy(fst_bike_pdl_by_AY, "FST_EAXA_RES")
fst_motor_ppa <- get_1st_py_cy(fst_auto_ppa_by_AY, "FST_EAXA_RES") + get_1st_py_cy(fst_bike_ppa_by_AY, "FST_EAXA_RES")
fst_motor_odpe_attr_nat <- get_1st_py_cy(fst_auto_odpe_attr_by_AY, "FST_EAXA_RES") + get_1st_py_cy(fst_auto_odpe_nat_by_AY, "FST_EAXA_RES") # ODはバイクなし
fst_motor_pi <- get_1st_py_cy(fst_auto_pi_sc_by_AY, "FST_EAXA_RES") + get_1st_py_cy(fst_auto_pi_lc_by_AY, "FST_EAXA_RES") +
  get_1st_py_cy(fst_bike_pi_sc_by_AY, "FST_EAXA_RES") +　get_1st_py_cy(fst_bike_pi_lc_by_AY, "FST_EAXA_RES")
fst_motor_fb <- get_1st_py_cy(fst_auto_fb_by_AY, "FST_EAXA_RES") # FBはバイクなし
fst_motor_le <- get_1st_py_cy(fst_auto_le_by_AY, "FST_EAXA_RES") + get_1st_py_cy(fst_bike_le_by_AY, "FST_EAXA_RES")
fst_motor_axap <- get_1st_py_cy(fst_auto_axap_by_AY, "FST_EAXA_RES") # AXA+はバイクなし
fst_motor_eq <- get_1st_py_cy(fst_auto_eq_by_AY, "FST_EAXA_RES") # EQはバイクなし

# 1st opinion
fst_motor_total_eres <- get_1st_py_cy(fst_motor_total_by_AY, "FST_EAXA_RES")
fst_motor_bil_eres <- get_1st_py_cy(fst_motor_bil_by_AY, "FST_EAXA_RES")
fst_motor_pdl_eres <- get_1st_py_cy(fst_motor_pdl_by_AY, "FST_EAXA_RES")
fst_motor_ppa_eres <- get_1st_py_cy(fst_motor_ppa_by_AY, "FST_EAXA_RES")
fst_motor_odpe_eres <- get_1st_py_cy(fst_auto_odpe_attr_nat_by_AY, "FST_EAXA_RES")
fst_motor_pi_eres <- get_1st_py_cy(fst_motor_pi_by_AY, "FST_EAXA_RES")
fst_motor_fb_eres <- get_1st_py_cy(fst_auto_fb_by_AY, "FST_EAXA_RES")
fst_motor_le_eres <- get_1st_py_cy(fst_motor_le_by_AY, "FST_EAXA_RES")
fst_motor_axap_eres <- get_1st_py_cy(fst_auto_axap_by_AY, "FST_EAXA_RES")
fst_motor_eq_eres <- get_1st_py_cy(fst_auto_eq_by_AY, "FST_EAXA_RES")
# Health and Pet
fst_health_eres <- get_1st_py_cy(fst_health_total_by_AY, "FST_EAXA_RES")
fst_pet_eres <- get_1st_py_cy(fst_pet_total_by_AY, "FST_EAXA_RES")
fst_eres <- rbind(
  fst_motor_total_eres,
  fst_motor_bil_eres,
  fst_motor_pdl_eres,
  fst_motor_ppa_eres,
  fst_motor_odpe_eres,
  fst_motor_pi_eres,
  fst_motor_fb_eres,
  fst_motor_le_eres,
  fst_motor_axap_eres,
  fst_motor_eq_eres,
  fst_health_eres,
  fst_pet_eres
)
# 2nd opinion
sec_motor_total_eres <- get_1st_py_cy(fst_motor_total_by_AY, "EAXA_RES")
sec_motor_bil_eres <- get_1st_py_cy(fst_motor_bil_by_AY, "EAXA_RES")
sec_motor_pdl_eres <- get_1st_py_cy(fst_motor_pdl_by_AY, "EAXA_RES")
sec_motor_ppa_eres <- get_1st_py_cy(fst_motor_ppa_by_AY, "EAXA_RES")
sec_motor_odpe_eres <- get_1st_py_cy(fst_auto_odpe_attr_nat_by_AY, "EAXA_RES")
sec_motor_pi_eres <- get_1st_py_cy(fst_motor_pi_by_AY, "EAXA_RES")
sec_motor_fb_eres <- get_1st_py_cy(fst_auto_fb_by_AY, "EAXA_RES")
sec_motor_le_eres <- get_1st_py_cy(fst_motor_le_by_AY, "EAXA_RES")
sec_motor_axap_eres <- get_1st_py_cy(fst_auto_axap_by_AY, "EAXA_RES")
sec_motor_eq_eres <- get_1st_py_cy(fst_auto_eq_by_AY, "EAXA_RES")
# Health and Pet
sec_health_eres <- get_1st_py_cy(fst_health_total_by_AY, "EAXA_RES")
sec_pet_eres <- get_1st_py_cy(fst_pet_total_by_AY, "EAXA_RES")
sec_eres <- rbind(
  sec_motor_total_eres,
  sec_motor_bil_eres,
  sec_motor_pdl_eres,
  sec_motor_ppa_eres,
  sec_motor_odpe_eres,
  sec_motor_pi_eres,
  sec_motor_fb_eres,
  sec_motor_le_eres,
  sec_motor_axap_eres,
  sec_motor_eq_eres,
  sec_health_eres,
  sec_pet_eres
)
# EAXA LR
# 1st opinion
fst_motor_total_elr <- get_1st_py_cy(fst_motor_total_by_AY, "FST_EAXA_LR")
fst_motor_bil_elr <- get_1st_py_cy(fst_motor_bil_by_AY, "FST_EAXA_LR")
fst_motor_pdl_elr <- get_1st_py_cy(fst_motor_pdl_by_AY, "FST_EAXA_LR")
fst_motor_ppa_elr <- get_1st_py_cy(fst_motor_ppa_by_AY, "FST_EAXA_LR")
fst_motor_odpe_elr <- get_1st_py_cy(fst_auto_odpe_attr_nat_by_AY, "FST_EAXA_LR")
fst_motor_pi_elr <- get_1st_py_cy(fst_motor_pi_by_AY, "FST_EAXA_LR")
fst_motor_fb_elr <- get_1st_py_cy(fst_auto_fb_by_AY, "FST_EAXA_LR")
fst_motor_le_elr <- get_1st_py_cy(fst_motor_le_by_AY, "FST_EAXA_LR")
fst_motor_axap_elr <- get_1st_py_cy(fst_auto_axap_by_AY, "FST_EAXA_LR")
fst_motor_eq_elr <- get_1st_py_cy(fst_auto_eq_by_AY, "FST_EAXA_LR")
# Health and Pet
fst_health_elr <- get_1st_py_cy(fst_health_total_by_AY, "FST_EAXA_LR")
fst_pet_elr <- get_1st_py_cy(fst_pet_total_by_AY, "FST_EAXA_LR")
fst_elr <- rbind(
                fst_motor_total_elr,
                fst_motor_bil_elr,
                fst_motor_pdl_elr,
                fst_motor_ppa_elr,
                fst_motor_odpe_elr,
                fst_motor_pi_elr,
                fst_motor_fb_elr,
                fst_motor_le_elr,
                fst_motor_axap_elr,
                fst_motor_eq_elr,
                fst_health_elr,
                fst_pet_elr
                )

# 2nd opinion
sec_motor_total_elr <- get_1st_py_cy(fst_motor_total_by_AY, "EAXA_LR")
sec_motor_bil_elr <- get_1st_py_cy(fst_motor_bil_by_AY, "EAXA_LR")
sec_motor_pdl_elr <- get_1st_py_cy(fst_motor_pdl_by_AY, "EAXA_LR")
sec_motor_ppa_elr <- get_1st_py_cy(fst_motor_ppa_by_AY, "EAXA_LR")
sec_motor_odpe_elr <- get_1st_py_cy(fst_auto_odpe_attr_nat_by_AY, "EAXA_LR")
sec_motor_pi_elr <- get_1st_py_cy(fst_motor_pi_by_AY, "EAXA_LR")
sec_motor_fb_elr <- get_1st_py_cy(fst_auto_fb_by_AY, "EAXA_LR")
sec_motor_le_elr <- get_1st_py_cy(fst_motor_le_by_AY, "EAXA_LR")
sec_motor_axap_elr <- get_1st_py_cy(fst_auto_axap_by_AY, "EAXA_LR")
sec_motor_eq_elr <- get_1st_py_cy(fst_auto_eq_by_AY, "EAXA_LR")
# Health and Pet
sec_health_elr <- get_1st_py_cy(fst_health_total_by_AY, "EAXA_LR")
sec_pet_elr <- get_1st_py_cy(fst_pet_total_by_AY, "EAXA_LR")
sec_elr <- rbind(
  sec_motor_total_elr,
  sec_motor_bil_elr,
  sec_motor_pdl_elr,
  sec_motor_ppa_elr,
  sec_motor_odpe_elr,
  sec_motor_pi_elr,
  sec_motor_fb_elr,
  sec_motor_le_elr,
  sec_motor_axap_elr,
  sec_motor_eq_elr,
  sec_health_elr,
  sec_pet_elr
)

# EAXA Frequency
# 1st opinion
fst_motor_total_efreq <- get_1st_py_cy(fst_motor_total_by_AY, "FST_EAXA_FREQ")
fst_motor_bil_efreq <- get_1st_py_cy(fst_motor_bil_by_AY, "FST_EAXA_FREQ")
fst_motor_pdl_efreq <- get_1st_py_cy(fst_motor_pdl_by_AY, "FST_EAXA_FREQ")
fst_motor_ppa_efreq <- get_1st_py_cy(fst_motor_ppa_by_AY, "FST_EAXA_FREQ")
fst_motor_odpe_efreq <- get_1st_py_cy(fst_auto_odpe_attr_nat_by_AY, "FST_EAXA_FREQ")
fst_motor_pi_efreq <- get_1st_py_cy(fst_motor_pi_by_AY, "FST_EAXA_FREQ")
fst_motor_fb_efreq <- get_1st_py_cy(fst_auto_fb_by_AY, "FST_EAXA_FREQ")
fst_motor_le_efreq <- get_1st_py_cy(fst_motor_le_by_AY, "FST_EAXA_FREQ")
fst_motor_axap_efreq <- get_1st_py_cy(fst_auto_axap_by_AY, "FST_EAXA_FREQ")
fst_motor_eq_efreq <- get_1st_py_cy(fst_auto_eq_by_AY, "FST_EAXA_FREQ")
# Health and Pet
fst_health_efreq <- get_1st_py_cy(fst_health_total_by_AY, "FST_EAXA_FREQ")
fst_pet_efreq <- get_1st_py_cy(fst_pet_total_by_AY, "FST_EAXA_FREQ")
fst_efreq <- rbind(
  fst_motor_total_efreq,
  fst_motor_bil_efreq,
  fst_motor_pdl_efreq,
  fst_motor_ppa_efreq,
  fst_motor_odpe_efreq,
  fst_motor_pi_efreq,
  fst_motor_fb_efreq,
  fst_motor_le_efreq,
  fst_motor_axap_efreq,
  fst_motor_eq_efreq,
  fst_health_efreq,
  fst_pet_efreq
)
# 2nd opinion
sec_motor_total_efreq <- get_1st_py_cy(fst_motor_total_by_AY, "EAXA_FREQ")
sec_motor_bil_efreq <- get_1st_py_cy(fst_motor_bil_by_AY, "EAXA_FREQ")
sec_motor_pdl_efreq <- get_1st_py_cy(fst_motor_pdl_by_AY, "EAXA_FREQ")
sec_motor_ppa_efreq <- get_1st_py_cy(fst_motor_ppa_by_AY, "EAXA_FREQ")
sec_motor_odpe_efreq <- get_1st_py_cy(fst_auto_odpe_attr_nat_by_AY, "EAXA_FREQ")
sec_motor_pi_efreq <- get_1st_py_cy(fst_motor_pi_by_AY, "EAXA_FREQ")
sec_motor_fb_efreq <- get_1st_py_cy(fst_auto_fb_by_AY, "EAXA_FREQ")
sec_motor_le_efreq <- get_1st_py_cy(fst_motor_le_by_AY, "EAXA_FREQ")
sec_motor_axap_efreq <- get_1st_py_cy(fst_auto_axap_by_AY, "EAXA_FREQ")
sec_motor_eq_efreq <- get_1st_py_cy(fst_auto_eq_by_AY, "EAXA_FREQ")
# Health and Pet
sec_health_efreq <- get_1st_py_cy(fst_health_total_by_AY, "EAXA_FREQ")
sec_pet_efreq <- get_1st_py_cy(fst_pet_total_by_AY, "EAXA_FREQ")
sec_efreq <- rbind(
  sec_motor_total_efreq,
  sec_motor_bil_efreq,
  sec_motor_pdl_efreq,
  sec_motor_ppa_efreq,
  sec_motor_odpe_efreq,
  sec_motor_pi_efreq,
  sec_motor_fb_efreq,
  sec_motor_le_efreq,
  sec_motor_axap_efreq,
  sec_motor_eq_efreq,
  sec_health_efreq,
  sec_pet_efreq
)

# EAXA Severity
# 1st opinion
fst_motor_total_esev <- get_1st_py_cy(fst_motor_total_by_AY, "FST_EAXA_SEV")
fst_motor_bil_esev <- get_1st_py_cy(fst_motor_bil_by_AY, "FST_EAXA_SEV")
fst_motor_pdl_esev <- get_1st_py_cy(fst_motor_pdl_by_AY, "FST_EAXA_SEV")
fst_motor_ppa_esev <- get_1st_py_cy(fst_motor_ppa_by_AY, "FST_EAXA_SEV")
fst_motor_odpe_esev <- get_1st_py_cy(fst_auto_odpe_attr_nat_by_AY, "FST_EAXA_SEV")
fst_motor_pi_esev <- get_1st_py_cy(fst_motor_pi_by_AY, "FST_EAXA_SEV")
fst_motor_fb_esev <- get_1st_py_cy(fst_auto_fb_by_AY, "FST_EAXA_SEV")
fst_motor_le_esev <- get_1st_py_cy(fst_motor_le_by_AY, "FST_EAXA_SEV")
fst_motor_axap_esev <- get_1st_py_cy(fst_auto_axap_by_AY, "FST_EAXA_SEV")
fst_motor_eq_esev <- get_1st_py_cy(fst_auto_eq_by_AY, "FST_EAXA_SEV")
# Health and Pet
fst_health_esev <- get_1st_py_cy(fst_health_total_by_AY, "FST_EAXA_SEV")
fst_pet_esev <- get_1st_py_cy(fst_pet_total_by_AY, "FST_EAXA_SEV")
fst_esev <- rbind(
  fst_motor_total_esev,
  fst_motor_bil_esev,
  fst_motor_pdl_esev,
  fst_motor_ppa_esev,
  fst_motor_odpe_esev,
  fst_motor_pi_esev,
  fst_motor_fb_esev,
  fst_motor_le_esev,
  fst_motor_axap_esev,
  fst_motor_eq_esev,
  fst_health_esev,
  fst_pet_esev
)
# 2nd opinion
sec_motor_total_esev <- get_1st_py_cy(fst_motor_total_by_AY, "EAXA_AVG_COST")
sec_motor_bil_esev <- get_1st_py_cy(fst_motor_bil_by_AY, "EAXA_AVG_COST")
sec_motor_pdl_esev <- get_1st_py_cy(fst_motor_pdl_by_AY, "EAXA_AVG_COST")
sec_motor_ppa_esev <- get_1st_py_cy(fst_motor_ppa_by_AY, "EAXA_AVG_COST")
sec_motor_odpe_esev <- get_1st_py_cy(fst_auto_odpe_attr_nat_by_AY, "EAXA_AVG_COST")
sec_motor_pi_esev <- get_1st_py_cy(fst_motor_pi_by_AY, "EAXA_AVG_COST")
sec_motor_fb_esev <- get_1st_py_cy(fst_auto_fb_by_AY, "EAXA_AVG_COST")
sec_motor_le_esev <- get_1st_py_cy(fst_motor_le_by_AY, "EAXA_AVG_COST")
sec_motor_axap_esev <- get_1st_py_cy(fst_auto_axap_by_AY, "EAXA_AVG_COST")
sec_motor_eq_esev <- get_1st_py_cy(fst_auto_eq_by_AY, "EAXA_AVG_COST")
# Health and Pet
sec_health_esev <- get_1st_py_cy(fst_health_total_by_AY, "EAXA_AVG_COST")
sec_pet_esev <- get_1st_py_cy(fst_pet_total_by_AY, "EAXA_AVG_COST")
sec_esev <- rbind(
  sec_motor_total_esev,
  sec_motor_bil_esev,
  sec_motor_pdl_esev,
  sec_motor_ppa_esev,
  sec_motor_odpe_esev,
  sec_motor_pi_esev,
  sec_motor_fb_esev,
  sec_motor_le_esev,
  sec_motor_axap_esev,
  sec_motor_eq_esev,
  sec_health_esev,
  sec_pet_esev
)
covname <- c("Motor", "BIL", "PDL", "PPA", "OD&PE", "PI", "FB", "LE", "AXAP", "EQ", "Health", "Pet")
df_fst_sec_metrics <- data.frame(covname, fst_eres, sec_eres, fst_elr, sec_elr, fst_efreq, sec_efreq, fst_esev, sec_esev)
colnames(df_fst_sec_metrics) <- c("covname", "fst_py_eres", "fst_cy_eres", "sec_py_eres", "sec_cy_eres",
                                  "fst_py_elr", "fst_cy_elr", "sec_py_elr", "sec_cy_elr",
                                  "fst_py_efreq", "fst_cy_efreq", "sec_py_efreq", "sec_cy_efreq",
                                  "fst_py_esev", "fst_cy_esev", "sec_py_esev", "sec_cy_esev"
                                  )

# Auto
fst_auto_bil <- get_1st_py_cy(fst_auto_bil_sc_by_AY, "FST_EAXA_RES") + get_1st_py_cy(fst_auto_bil_lc_by_AY, "FST_EAXA_RES")
fst_auto_pdl <- get_1st_py_cy(fst_auto_pdl_by_AY, "FST_EAXA_RES")
fst_auto_ppa <- get_1st_py_cy(fst_auto_ppa_by_AY, "FST_EAXA_RES")
fst_auto_odpe_attr_nat <- get_1st_py_cy(fst_auto_odpe_attr_by_AY, "FST_EAXA_RES") + get_1st_py_cy(fst_auto_odpe_nat_by_AY, "FST_EAXA_RES") # ODはバイクなし
fst_auto_pi <- get_1st_py_cy(fst_auto_pi_sc_by_AY, "FST_EAXA_RES") + get_1st_py_cy(fst_auto_pi_lc_by_AY, "FST_EAXA_RES")
fst_auto_fb <- get_1st_py_cy(fst_auto_fb_by_AY, "FST_EAXA_RES") # FBはバイクなし
fst_auto_le <- get_1st_py_cy(fst_auto_le_by_AY, "FST_EAXA_RES")
fst_auto_axap <- get_1st_py_cy(fst_auto_axap_by_AY, "FST_EAXA_RES") # AXA+はバイクなし
fst_auto_eq <- get_1st_py_cy(fst_auto_eq_by_AY, "FST_EAXA_RES") # EQはバイクなし
# Bike
fst_bike_bil <- get_1st_py_cy(fst_bike_bil_sc_by_AY, "FST_EAXA_RES") +　get_1st_py_cy(fst_bike_bil_lc_by_AY, "FST_EAXA_RES")
fst_bike_pdl <- get_1st_py_cy(fst_bike_pdl_by_AY, "FST_EAXA_RES")
fst_bike_ppa <- get_1st_py_cy(fst_bike_ppa_by_AY, "FST_EAXA_RES")
fst_bike_pi <- get_1st_py_cy(fst_bike_pi_sc_by_AY, "FST_EAXA_RES") +　get_1st_py_cy(fst_bike_pi_lc_by_AY, "FST_EAXA_RES")
fst_bike_le <- get_1st_py_cy(fst_bike_le_by_AY, "FST_EAXA_RES")
# Health and Pet
fst_health <- get_1st_py_cy(fst_health_total_by_AY, "FST_EAXA_RES")
fst_pet <- get_1st_py_cy(fst_pet_total_by_AY, "FST_EAXA_RES")
#
# 1st opinion summary PY and CY for Flash Report
#
# Motor, Health, Pet summary
fst_motor <- as.data.frame(rbind(fst_motor_bil, fst_motor_pdl, fst_motor_ppa, fst_motor_odpe_attr_nat, fst_motor_pi, fst_motor_fb, fst_motor_le, fst_motor_axap, fst_motor_eq))
fst_motor_tot <- fst_motor %>%
  summarize_if(is.numeric, sum, na.rm=TRUE)
fst_all_lob <- rbind(fst_motor_tot, fst_motor, fst_health, fst_pet)
all_covs <- c("Motor", "BIL", "PDL", "PPA", "ODPE", "PI",  "FB", "LX", "AXA+", "EQ", "Health", "Pet")
fst_all_lob <- cbind(all_covs, fst_all_lob)
# Auto
fst_auto <- as.data.frame(rbind(fst_auto_bil, fst_auto_pdl, fst_auto_ppa, fst_auto_odpe_attr_nat, fst_auto_pi, fst_auto_fb, fst_auto_le, fst_auto_axap, fst_auto_eq))
fst_auto_tot <- fst_auto %>%
  summarize_if(is.numeric, sum, na.rm=TRUE)
fst_auto_all <- rbind(fst_auto_tot, fst_auto)
auto_covs <- c("Auto", "BIL", "PDL", "PPA", "ODPE", "PI",  "FB", "LX", "AXA+", "EQ")
fst_auto_lob <- cbind(auto_covs, fst_auto_all)
# Bike
fst_bike <- as.data.frame(rbind(fst_bike_bil, fst_bike_pdl, fst_bike_ppa, fst_bike_pi, fst_bike_le))
fst_bike_tot <- fst_bike %>%
  summarize_if(is.numeric, sum, na.rm=TRUE)
fst_bike_all <- rbind(fst_bike_tot, fst_bike)
bike_covs <- c("Bike", "BIL", "PDL", "PPA", "PI", "LX")
fst_bike_lob <- cbind(bike_covs, fst_bike_all)
#
# Export to Excel -------------------------------
#
# 1st and 2nd opinion summary
#
# ------Export summary tables to Excel
l <- list(
  "motor_tot" = fst_motor_total_by_AY,
  "motor_bil" = fst_motor_bil_by_AY,
  "motor_pdl" = fst_motor_pdl_by_AY,
  "motor_ppa" = fst_motor_ppa_by_AY,
  "motor_pi" = fst_motor_pi_by_AY,
  "motor_le" = fst_motor_le_by_AY,
  "auto_bil_sc" = fst_auto_bil_sc_by_AY,
  "auto_bil_lc" = fst_auto_bil_lc_by_AY,
  "auto_pdl" = fst_auto_pdl_by_AY,
  "auto_ppa" = fst_auto_ppa_by_AY,
  "auto_odpe_attr" = fst_auto_odpe_attr_by_AY,
  "auto_odpe_nat" = fst_auto_odpe_nat_by_AY,
  "auto_pi_sc" = fst_auto_pi_sc_by_AY,
  "auto_pi_lc" = fst_auto_pi_lc_by_AY,
  "auto_axap" = fst_auto_axap_by_AY,
  "auto_fb" = fst_auto_fb_by_AY,
  "auto_le" = fst_auto_le_by_AY,
  "auto_eq" = fst_auto_eq_by_AY,
  "bike_bil_sc" = fst_bike_bil_sc_by_AY,
  "bike_bil_lc" = fst_bike_bil_lc_by_AY,
  "bike_pdl" = fst_bike_pdl_by_AY,
  "bike_ppa" = fst_bike_ppa_by_AY,
  "bike_pi_sc" = fst_bike_pi_sc_by_AY,
  "bike_pi_lc" = fst_bike_pi_lc_by_AY,
  "bike_le" = fst_bike_le_by_AY,
  "health_total" = fst_health_total_by_AY,
  "pet_total" = fst_pet_total_by_AY,
  "all_py_cy" = df_fst_sec_metrics,
  "motor_tax" = motor_tax
)
setwd(current_work_wd)
write_xlsx(l, "Apr24_summary_1st_2nd_v2.xlsx")
#
#
#
#
# -----------------------------------------------------------
