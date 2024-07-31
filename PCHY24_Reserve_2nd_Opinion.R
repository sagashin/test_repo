# Create working directories

mainDir <- "C:/Users/shin.sagara/OneDrive - AXA/02. 保険引受リスク（P&C)/43.リザービング/2024/PCHY"

#subDir_1 <- "00_R_Program"
subDir_2 <- "01_Triangle_Data"
subDir_3 <- "02 Flash_Report"
subDir_4 <- "03 Communication"
subDir_5 <- "99 work"

#dir.create(file.path(mainDir, subDir_1), showWarnings = FALSE)
dir.create(file.path(mainDir, subDir_2), showWarnings = FALSE)
dir.create(file.path(mainDir, subDir_3), showWarnings = FALSE)
dir.create(file.path(mainDir, subDir_4), showWarnings = FALSE)
dir.create(file.path(mainDir, subDir_5), showWarnings = FALSE)

triangle_wd <- file.path(mainDir, subDir_2)
workfile_wd <- file.path(mainDir, subDir_5)
next_process_working_wd <- "C:/Users/shin.sagara/OneDrive - AXA/02. 保険引受リスク（P&C)/43.リザービング/2024/PCFY/99 work"

setwd(triangle_wd)

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
library(writexl)
library(res2ndopi)
library(lubridate)

end_date <- "2024/4/01" #日は月末の1日としている（プログラムの都合上）

# Read Excel --------------------------------------------------------------
# Motor
df <- read_excel("Triangle_Motor_202404.xlsx", sheet = "output_202404") 
# If the 'COV_NAME' of a row is neither 'OD' nor 'OD_EQ' and 'CL_Large' is 'Evente', 
# then we change 'CL_Large' to 'Attrit' for that row.
df$CL_Large <- ifelse(!df$COV_NAME %in% c("OD", "OD_EQ") & df$CL_Large == "Evente", "Attrit", df$CL_Large)
df <- df %>% 
  group_by(BIZ_LINE, PRODUCTS, COV_NAME, CL_Large, YM_EV_OCC, Dev_Mths) %>%
  summarise(paid=sum(paid), os=sum(os), inc=sum(inc), claims=sum(claims))


df <- df %>%
  rename(
    CL_Large2 = CL_Large) # CL_Largeのヘッダー名をCL_Large2へ修正
df$Dev_Mths <- df['Dev_Mths'][[1]] + 1 #数理のデータがMotorだけDevMonthが０から始まっているから調整
df$CL_Large2[df$CL_Large2 == "Evente"] <- "Evented" #HY2020 数理データの"Evented"が"Evente"に変更されていたので対応処理
# Non Motor
df_nonmotor <- read_excel("Triangle_NonMotor_202404.xlsx", sheet = "Data")
df_nonmotor$CL_Large2 <- "Attrit" #CL_Large2フィールドがNonMotorにはないので追加

df <- df[!(df$PRODUCTS=="BIKE" & df$COV_NAME=="PI" & df$CL_Large2=="Large" & df$YM_EV_OCC<=201207),] # イレギュラーなデータを削除
df_nonmotor <- df_nonmotor[!(df_nonmotor$YM_EV_OCC==200104),]

# ----------EP and Exposure-------------------------------
fst_auto_filename <- "01_IBNR_Auto_Linkout202404Finance.xlsx"
fst_bike_filename <- "01_IBNR_Bike_Linkout202404Finance.xlsx"
fst_pet_filename <- "01_IBNR_Pet_new_Linkout202404Finance.xlsx"
fst_health_filename <- "01_IBNR_Health_new_Linkout202404Finance.xlsx"
# set up dates
origin_date <- data.frame(ORIGIN_DATE = seq(ymd(19990601), ymd(end_date), "months")) # old >> read_excel("origin_ym_199906_202104.xlsx")
# --------------------------------------------------------

get_ep_exposure <- function(filename, sheetname, excl_row1, excl_row2, excl_row3, col_exp, col_ep, biz, prod, covname) {

  df <- read_excel(filename, sheet = sheetname)
  df <- df[-c(1:excl_row1, excl_row2:10000), c(col_exp, col_ep)]
  df <- cbind(origin_date, df)
  df$BIZ_LINE <- biz
  df$PRODUCTS <- prod
  df$COV_NAME <- covname
  colnames(df)[2] <- "EXPOSURE"
  colnames(df)[3] <-"EARNINGS"
  df$EXPOSURE <- as.numeric(df['EXPOSURE'][[1]])
  df$EARNINGS <- as.numeric(df['EARNINGS'][[1]])
  df <- df[, c(4,5,6,1,2,3)]
  return(df)

}
get_tax <- function(filename, sheetname, excl_row1, excl_row2, excl_row3, col_tax, biz, prod, covname) {
  
  df <- read_excel(filename, sheet = sheetname)
  df <- df[-c(1:excl_row1, excl_row2:10000), c(col_tax)]
  df <- cbind(origin_date, df)
  df$BIZ_LINE <- biz
  df$PRODUCTS <- prod
  df$COV_NAME <- covname
  colnames(df)[2] <- "TAX"
  df$TAX <- as.numeric(df['TAX'][[1]])
  df$TAX[is.na(df$TAX)] <- 0
  #df <- df[, c(4,5,6,1,2,3)]
  df_sum <- df %>% 
    group_by(PRODUCTS, COV_NAME) %>% 
    summarise(TAX = sum(TAX))
  return(df_sum)
  
}

#update manually every time  
excl_rows2_input <- 316 #as of April 2024
excl_rows2_input_health_and_total <- 308 #as of April 2024

auto_bil <- get_ep_exposure(fst_auto_filename, sheetname="BIL_attr", excl_row1=16, excl_row2=excl_rows2_input, col_exp=9, col_ep=10, biz="Motor", prod="AUTO", covname="BIL")
auto_pdl <- get_ep_exposure(fst_auto_filename, sheetname="PDL_attr", excl_row1=16, excl_row2=excl_rows2_input, col_exp=9, col_ep=10, biz="Motor", prod="AUTO", covname="PDL")
auto_ppa <- get_ep_exposure(fst_auto_filename, sheetname="PPA_attr", excl_row1=16, excl_row2=excl_rows2_input, col_exp=9, col_ep=10, biz="Motor", prod="AUTO", covname="PPA")
auto_od <- get_ep_exposure(fst_auto_filename, sheetname="OD_attr", excl_row1=16, excl_row2=excl_rows2_input, col_exp=9, col_ep=10, biz="Motor", prod="AUTO", covname="OD")
auto_pe <- get_ep_exposure(fst_auto_filename, sheetname="PE_attr", excl_row1=16, excl_row2=excl_rows2_input, col_exp=9, col_ep=10, biz="Motor", prod="AUTO", covname="PE")
auto_pi <- get_ep_exposure(fst_auto_filename, sheetname="PI_attr", excl_row1=16, excl_row2=excl_rows2_input, col_exp=9, col_ep=10, biz="Motor", prod="AUTO", covname="PI")
auto_axap <- get_ep_exposure(fst_auto_filename, sheetname="AXA+_attr", excl_row1=16, excl_row2=excl_rows2_input, col_exp=9, col_ep=10, biz="Motor", prod="AUTO", covname="AXA+")
auto_fb <- get_ep_exposure(fst_auto_filename, sheetname="FB_attr", excl_row1=16, excl_row2=excl_rows2_input, col_exp=9, col_ep=10, biz="Motor", prod="AUTO", covname="FamilyBIKE")
auto_le <- get_ep_exposure(fst_auto_filename, sheetname="LX_attr", excl_row1=16, excl_row2=excl_rows2_input, col_exp=9, col_ep=10, biz="Motor", prod="AUTO", covname="LawyerE")
auto_eq <- get_ep_exposure(fst_auto_filename, sheetname="EQ_attr", excl_row1=16, excl_row2=excl_rows2_input, col_exp=9, col_ep=10, biz="Motor", prod="AUTO", covname="OD_EQ")[-c(1:164), ] #EQ special treatment
# Added July23
auto_rc <- get_ep_exposure(fst_auto_filename, sheetname="RC_attr", excl_row1=16, excl_row2=excl_rows2_input, col_exp=9, col_ep=10, biz="Motor", prod="AUTO", covname="RC")
auto_sk <- get_ep_exposure(fst_auto_filename, sheetname="SK", excl_row1=16, excl_row2=excl_rows2_input, col_exp=9, col_ep=10, biz="Motor", prod="AUTO", covname="SK")

bike_bil <- get_ep_exposure(fst_bike_filename, sheetname="BIL_attr", excl_row1=16, excl_row2=excl_rows2_input, col_exp=9, col_ep=10, biz="Motor", prod="BIKE", covname="BIL")[-c(1:71), ]
bike_pdl <- get_ep_exposure(fst_bike_filename, sheetname="PDL_attr", excl_row1=16, excl_row2=excl_rows2_input, col_exp=9, col_ep=10, biz="Motor", prod="BIKE", covname="PDL")[-c(1:71), ]
bike_ppa <- get_ep_exposure(fst_bike_filename, sheetname="PPA_attr", excl_row1=16, excl_row2=excl_rows2_input, col_exp=9, col_ep=10, biz="Motor", prod="BIKE", covname="PPA")[-c(1:71), ]
bike_pi <- get_ep_exposure(fst_bike_filename, sheetname="PI_attr", excl_row1=16, excl_row2=excl_rows2_input, col_exp=9, col_ep=10, biz="Motor", prod="BIKE", covname="PI")[-c(1:71), ]
bike_le <- get_ep_exposure(fst_bike_filename, sheetname="LX_attr", excl_row1=16, excl_row2=excl_rows2_input, col_exp=9, col_ep=10, biz="Motor", prod="BIKE", covname="LawyerE")[-c(1:71), ]

# PAの数理ファイルのフォーマット変更に対応 202103
pa_total <- get_ep_exposure(fst_health_filename, sheetname="PA_TOTAL", excl_row1=8, excl_row2=excl_rows2_input_health_and_total, col_exp=5, col_ep=6, biz="PA", prod="PA", covname="Total")[-c(1:23), ]
pet_total <- get_ep_exposure(fst_pet_filename, sheetname="PET", excl_row1=16, excl_row2=excl_rows2_input, col_exp=9, col_ep=10, biz="PET", prod="PET", covname="Total")[-c(1:127), ]

df_exp_ep <- rbind(
  auto_bil,
  auto_pdl,
  auto_ppa,
  auto_od,
  auto_pe,
  auto_pi,
  auto_axap,
  auto_fb,
  auto_le,
  auto_eq,
  #auto_rc, #added PCFY23
  #auto_sk,
  bike_bil,
  bike_pdl,
  bike_ppa,
  bike_pi,
  bike_le,
  pa_total,
  pet_total
)

df_exp_ep$PRODUCTS[df_exp_ep$PRODUCTS == "AUTO"] <- "Auto"
df_exp_ep[is.na(df_exp_ep['EXPOSURE']), 'EXPOSURE'] <- 0
df_exp_ep[is.na(df_exp_ep['EARNINGS']), 'EARNINGS'] <- 0

# Tax treatment
auto_bil_tax <- get_tax(fst_auto_filename, sheetname="BIL_attr", excl_row1=16, excl_row2=excl_rows2_input, col_tax=55, biz="Motor", prod="AUTO", covname="BIL")
auto_pdl_tax <- get_tax(fst_auto_filename, sheetname="PDL_attr", excl_row1=16, excl_row2=excl_rows2_input, col_tax=55, biz="Motor", prod="AUTO", covname="PDL")
auto_ppa_tax <- get_tax(fst_auto_filename, sheetname="PPA_attr", excl_row1=16, excl_row2=excl_rows2_input, col_tax=55, biz="Motor", prod="AUTO", covname="PPA")
auto_od_tax <- get_tax(fst_auto_filename, sheetname="OD_attr", excl_row1=16, excl_row2=excl_rows2_input, col_tax=55, biz="Motor", prod="AUTO", covname="OD")
auto_pe_tax <- get_tax(fst_auto_filename, sheetname="PE_attr", excl_row1=16, excl_row2=excl_rows2_input, col_tax=55, biz="Motor", prod="AUTO", covname="PE")
auto_pi_tax <- get_tax(fst_auto_filename, sheetname="PI_attr", excl_row1=16, excl_row2=excl_rows2_input, col_tax=55, biz="Motor", prod="AUTO", covname="PI")
auto_axap_tax <- get_tax(fst_auto_filename, sheetname="AXA+_attr", excl_row1=16, excl_row2=excl_rows2_input, col_tax=55, biz="Motor", prod="AUTO", covname="AXA+")
auto_fb_tax <- get_tax(fst_auto_filename, sheetname="FB_attr", excl_row1=16, excl_row2=excl_rows2_input, col_tax=55, biz="Motor", prod="AUTO", covname="FamilyBIKE")
auto_le_tax <- get_tax(fst_auto_filename, sheetname="LX_attr", excl_row1=16, excl_row2=excl_rows2_input, col_tax=55, biz="Motor", prod="AUTO", covname="LawyerE")
auto_eq_tax <- get_tax(fst_auto_filename, sheetname="EQ_attr", excl_row1=16, excl_row2=excl_rows2_input, col_tax=55, biz="Motor", prod="AUTO", covname="OD_EQ")

bike_bil_tax <- get_tax(fst_bike_filename, sheetname="BIL_attr", excl_row1=16, excl_row2=excl_rows2_input, col_tax=55, biz="Motor", prod="BIKE", covname="BIL")
bike_pdl_tax <- get_tax(fst_bike_filename, sheetname="PDL_attr", excl_row1=16, excl_row2=excl_rows2_input, col_tax=55, biz="Motor", prod="BIKE", covname="PDL")
bike_ppa_tax <- get_tax(fst_bike_filename, sheetname="PPA_attr", excl_row1=16, excl_row2=excl_rows2_input, col_tax=55, biz="Motor", prod="BIKE", covname="PPA")
bike_pi_tax <- get_tax(fst_bike_filename, sheetname="PI_attr", excl_row1=16, excl_row2=excl_rows2_input, excl_row3=406, col_tax=55, biz="Motor", prod="BIKE", covname="PI")
bike_le_tax <- get_tax(fst_bike_filename, sheetname="LX_attr", excl_row1=16, excl_row2=excl_rows2_input, col_tax=55, biz="Motor", prod="BIKE", covname="LawyerE")

auto_bike_tax <- rbind(auto_bil_tax, auto_pdl_tax, auto_ppa_tax, auto_od_tax, auto_pe_tax, auto_pi_tax, auto_axap_tax, auto_fb_tax, auto_le_tax, auto_eq_tax,
      bike_bil_tax, bike_pdl_tax, bike_ppa_tax, bike_pi_tax, bike_le_tax)

motor_tax <- auto_bike_tax %>% 
  group_by(COV_NAME) %>% 
  summarise(TAX = sum(TAX))

#----------------------------------

df_motor <- df %>%
  group_by(BIZ_LINE, YM_EV_OCC, Dev_Mths) %>%
  filter(PRODUCTS == "total") %>%
  summarise(paid=sum(paid), os=sum(os), inc=sum(inc), claims=sum(claims)) %>%
  mutate(PRODUCTS = "Motor", COV_NAME = "Total", CL_Large2="Total")


#Summary for total coverage Motor (attritional and atypical total)
df_total <- df %>%
  group_by(BIZ_LINE, PRODUCTS, COV_NAME, YM_EV_OCC, Dev_Mths) %>%
  summarise(paid=sum(paid), os=sum(os), inc=sum(inc), claims=sum(claims)) %>%
  mutate(CL_Large2="Total")
#ODPEを抽出
df_odpe <- df %>%
  filter(COV_NAME == "OD" | COV_NAME == "PE") %>%
  group_by(BIZ_LINE, PRODUCTS, CL_Large2, YM_EV_OCC, Dev_Mths) %>%
  summarise(paid=sum(paid), os=sum(os), inc=sum(inc), claims=sum(claims)) %>%
  mutate(COV_NAME="ODPE")
#BI, PIはLargeとSmallに分けるのでデータを分けて持つ
df_bi_pi <- df %>%
  filter(COV_NAME == "BIL" | COV_NAME == "PI")

#Non Motor
#NonMotor Totalを集計 coverage totalが必要 すべてAttrit
df_nm_total <- df_nonmotor %>%
  group_by(BIZ_LINE, CL_Large2, YM_EV_OCC, Dev_Mths) %>%
  summarise(paid=sum(paid), os=sum(os), inc=sum(inc), claims=sum(claims)) %>%
  mutate(PRODUCTS = "Total", COV_NAME="Total")
#PA Total
df_pa_total <- df_nonmotor %>%
  filter(BIZ_LINE=="PA") %>%
  group_by(BIZ_LINE, CL_Large2, YM_EV_OCC, Dev_Mths) %>%
  summarise(paid=sum(paid), os=sum(os), inc=sum(inc), claims=sum(claims)) %>%
  mutate(PRODUCTS = "PA", COV_NAME="Total")
#Pet Total
df_pet_total <- df_nonmotor %>%
  filter(BIZ_LINE=="PET") %>%
  group_by(BIZ_LINE, CL_Large2, YM_EV_OCC, Dev_Mths) %>%
  summarise(paid=sum(paid), os=sum(os), inc=sum(inc), claims=sum(claims)) %>%
  mutate(PRODUCTS = "PET", COV_NAME="Total")
#Death+Disabilityを抽出
df_nm_dd <- df_nonmotor %>%
  filter(COV_NAME == "DEATH" | COV_NAME == "DISBL") %>%
  group_by(BIZ_LINE, CL_Large2, YM_EV_OCC, Dev_Mths) %>%
  summarise(paid=sum(paid), os=sum(os), inc=sum(inc), claims=sum(claims)) %>%
  mutate(COV_NAME="DTH_DISBL", PRODUCTS="PA")
#PA other を抽出
df_nm_oth <- df_nonmotor %>%
  filter(BIZ_LINE == "PA" & (COV_NAME != "DEATH" & COV_NAME != "DISBL")) %>%
  group_by(BIZ_LINE, CL_Large2, YM_EV_OCC, Dev_Mths) %>%
  summarise(paid=sum(paid), os=sum(os), inc=sum(inc), claims=sum(claims)) %>%
  mutate(COV_NAME="PA_OTH", PRODUCTS="PA")
#Combine all
#df_with_total <- as.data.frame(bind_rows(df_total, df_odpe, df_bi_pi, df_pa_total, df_pet_total))
df_with_total <- as.data.frame(bind_rows(df_total, df_odpe, df_motor, df_bi_pi, df_pa_total, df_pet_total))
#Nest
df_for_tri <- df_with_total %>%
  group_by(BIZ_LINE, PRODUCTS, COV_NAME, CL_Large2) %>%
  nest()



# Exposure, EP ------------------------------------------------------------

#ODPEを抽出
df_exp_ep2 <- df_exp_ep %>%
  separate(ORIGIN_DATE, sep = "-", into= c("year", "month", "day")) %>%
  mutate(ym=paste(year, month, sep = "")) %>%
  dplyr::select(BIZ_LINE, PRODUCTS, COV_NAME, ym, EXPOSURE, EARNINGS)
df_odpe_exp <- df_exp_ep2 %>%
  filter(COV_NAME == "OD" | COV_NAME == "PE") %>%
  group_by(BIZ_LINE, PRODUCTS, ym) %>%
  summarise(EXPOSURE=sum(EXPOSURE)/2, EARNINGS=sum(EARNINGS)) %>%
  mutate(COV_NAME="ODPE")
#PA
df_pa_exp <- df_exp_ep2 %>%
  filter(BIZ_LINE=="PA" & COV_NAME=="Total") %>%
  group_by(BIZ_LINE, ym) %>%
  summarise(EXPOSURE=sum(EXPOSURE), EARNINGS=sum(EARNINGS)) %>%
  mutate(PRODUCTS="PA", COV_NAME="Total")

# Motor
df_motor_exp <- df_exp_ep2 %>%
  filter(BIZ_LINE=="Motor") %>%
  group_by(BIZ_LINE, ym) %>%
  summarise(EXPOSURE=sum(EXPOSURE), EARNINGS=sum(EARNINGS)) %>%
  mutate(PRODUCTS="Motor", COV_NAME="Total")

#Combine all exp ep
#df_exp_ep2 <- bind_rows(df_exp_ep2, df_odpe_exp, df_pa_exp)
df_exp_ep2 <- bind_rows(df_exp_ep2, df_odpe_exp, df_motor_exp, df_pa_exp)
#Nest
df_exp_ep_nest <- df_exp_ep2 %>%
  as.data.frame() %>%  mutate(ym = as.numeric(ym)) %>%
  group_by(BIZ_LINE, PRODUCTS, COV_NAME) %>%
  nest()
#exhibit preparation

df_exp_ep_nest <- df_exp_ep_nest %>%
  mutate(
    cont_exhibit = data %>% map(cont_exhibit)
  )

# Merge Claim and Exposure and EP -----------------------------------------

#Add key to claim
df_for_tri$key <- paste(df_for_tri$PRODUCTS, df_for_tri$COV_NAME)
#add key to cont
df_exp_ep_nest$key <- paste(df_exp_ep_nest$PRODUCTS, df_exp_ep_nest$COV_NAME)
df_exp_ep_nest <- df_exp_ep_nest %>%
  ungroup() %>% #added 20200601 seems new dplyr logic
  dplyr::select(data, cont_exhibit, key)

#Merge
df_claim_cont <- merge(x=df_for_tri, y=df_exp_ep_nest, by="key") %>%
  dplyr::select(-key) %>%
  dplyr::rename(data=data.x)


# Prepare reserving data --------------------------------------------------

Reserve_Data <- df_claim_cont %>%
  mutate(
    inc_tri   = data %>% map(make_inc_tri),
    paid_tri  = data %>% map(make_paid_tri),
    os_tri    = data %>% map(make_os_tri),
    clcnt_tri = data %>% map(make_clcnt_tri),
    lkr_tri   = inc_tri %>% map(make_lkr_tri),
    mthly_inc = inc_tri %>% map(get_mthly_inc),
    #linkratio without adjustment
    lkr_smpl = inc_tri %>% map(lkr_smpl_avg),
    lkr_vwtd = inc_tri %>% map(lkr_vwtd),
    lkr_excl_beg = map2(inc_tri, 10, lkr_vwtd_excl_beg),
    lkr12 = inc_tri %>% map(lkr_ex_maxmin12),
    lkr24 = inc_tri %>% map(lkr_ex_maxmin24),
    lkr36 = inc_tri %>% map(lkr_ex_maxmin36),
    lkr3 = inc_tri %>% map(lkr_ex_maxmin3),
    lkr5 = inc_tri %>% map(lkr_ex_maxmin5),
    lkr7 = inc_tri %>% map(lkr_ex_maxmin7),
    lkr_clcnt_smpl = clcnt_tri %>% map(lkr_smpl_avg),
    lkr_paid_smpl = paid_tri %>% map(lkr_smpl_avg),
    lkr12_paid = paid_tri %>% map(lkr_ex_maxmin12),
    #linkratio with adjustment
    lkr_smpl_adj = map(lkr_smpl, lkr_adjust),
    lkr_vwtd_adj = map(lkr_vwtd, lkr_adjust),
    lkr_excl_beg_adj = map(lkr_excl_beg, lkr_adjust),
    lkr12_adj = map(lkr12, lkr_adjust),
    lkr24_adj = map(lkr24, lkr_adjust),
    lkr36_adj = map(lkr36, lkr_adjust),
    lkr3_adj = map(lkr3, lkr_adjust),
    lkr5_adj = map(lkr5, lkr_adjust),
    lkr7_adj = map(lkr7, lkr_adjust),
    #linkratio cumulated
    lkr_smpl_cum = lkr_smpl %>% map(lkr_cumulate),
    lkr_vwtd_cum = lkr_vwtd %>% map(lkr_cumulate),
    lkr12_cum = lkr12 %>% map(lkr_cumulate),
    lkr24_cum = lkr24 %>% map(lkr_cumulate),
    lkr36_cum = lkr36 %>% map(lkr_cumulate),
    lkr3_cum = lkr3 %>% map(lkr_cumulate),
    lkr5_cum = lkr5 %>% map(lkr_cumulate),
    lkr7_cum = lkr7 %>% map(lkr_cumulate),
    lkr_clcnt_smpl_cum = lkr_clcnt_smpl %>% map(lkr_cumulate),

    lkr_smpl_adj_cum = lkr_smpl_adj %>% map(lkr_cumulate),
    lkr_vwtd_adj_cum = lkr_vwtd_adj %>% map(lkr_cumulate),
    lkr12_adj_cum = lkr12_adj %>% map(lkr_cumulate),
    lkr24_adj_cum = lkr24_adj %>% map(lkr_cumulate),
    lkr36_adj_cum = lkr36_adj %>% map(lkr_cumulate),
    lkr3_adj_cum = lkr3_adj %>% map(lkr_cumulate),
    lkr5_adj_cum = lkr5_adj %>% map(lkr_cumulate),
    lkr7_adj_cum = lkr7_adj %>% map(lkr_cumulate),
    #Exhibits
    exhibit_smpl = pmap(list(inc_tri, paid_tri, os_tri, clcnt_tri, lkr_smpl_cum, lkr_clcnt_smpl_cum, cont_exhibit), make_exhibit),
    exhibit_vwtd = pmap(list(inc_tri, paid_tri, os_tri, clcnt_tri, lkr_vwtd_cum, lkr_clcnt_smpl_cum, cont_exhibit), make_exhibit),
    exhibit3 = pmap(list(inc_tri, paid_tri, os_tri, clcnt_tri, lkr3_cum, lkr_clcnt_smpl_cum, cont_exhibit), make_exhibit),
    exhibit5 = pmap(list(inc_tri, paid_tri, os_tri, clcnt_tri, lkr5_cum, lkr_clcnt_smpl_cum, cont_exhibit), make_exhibit),
    exhibit7 = pmap(list(inc_tri, paid_tri, os_tri, clcnt_tri, lkr7_cum, lkr_clcnt_smpl_cum, cont_exhibit), make_exhibit),
    exhibit12 = pmap(list(inc_tri, paid_tri, os_tri, clcnt_tri, lkr12_cum, lkr_clcnt_smpl_cum, cont_exhibit), make_exhibit),
    exhibit24 = pmap(list(inc_tri, paid_tri, os_tri, clcnt_tri, lkr24_cum, lkr_clcnt_smpl_cum, cont_exhibit), make_exhibit),
    exhibit36 = pmap(list(inc_tri, paid_tri, os_tri, clcnt_tri, lkr36_cum, lkr_clcnt_smpl_cum, cont_exhibit), make_exhibit),


    exhibit_smpl_adj = pmap(list(inc_tri, paid_tri, os_tri, clcnt_tri, lkr_smpl_adj_cum, lkr_clcnt_smpl_cum, cont_exhibit), make_exhibit),
    exhibit_vwtd_adj = pmap(list(inc_tri, paid_tri, os_tri, clcnt_tri, lkr_vwtd_adj_cum, lkr_clcnt_smpl_cum, cont_exhibit), make_exhibit),
    exhibit3_adj = pmap(list(inc_tri, paid_tri, os_tri, clcnt_tri, lkr3_adj_cum, lkr_clcnt_smpl_cum, cont_exhibit), make_exhibit),
    exhibit5_adj = pmap(list(inc_tri, paid_tri, os_tri, clcnt_tri, lkr5_adj_cum, lkr_clcnt_smpl_cum, cont_exhibit), make_exhibit),
    exhibit7_adj = pmap(list(inc_tri, paid_tri, os_tri, clcnt_tri, lkr7_adj_cum, lkr_clcnt_smpl_cum, cont_exhibit), make_exhibit),
    exhibit12_adj = pmap(list(inc_tri, paid_tri, os_tri, clcnt_tri, lkr12_adj_cum, lkr_clcnt_smpl_cum, cont_exhibit), make_exhibit),
    exhibit24_adj = pmap(list(inc_tri, paid_tri, os_tri, clcnt_tri, lkr24_adj_cum, lkr_clcnt_smpl_cum, cont_exhibit), make_exhibit),
    exhibit36_adj = pmap(list(inc_tri, paid_tri, os_tri, clcnt_tri, lkr36_adj_cum, lkr_clcnt_smpl_cum, cont_exhibit), make_exhibit),


    #Plot lkr
    #lkr_plot = map2(inc_tri, lkr_tri, plot_lkr),
    #monthly eaxa
    mthly_eaxa12 = map2(mthly_inc, lkr12_cum, get_mthly_eaxa)

  )

# Projection --------------------------------------------------------------


# Plot and Exhibit ----------------------------------------------------------------

eaxa_res_by_model <- Reserve_Data %>%
  unnest(exhibit_vwtd, exhibit3, exhibit5, exhibit7, exhibit12, exhibit24, exhibit36, exhibit_vwtd_adj, exhibit3_adj, exhibit5_adj, exhibit7_adj, exhibit12_adj, exhibit24_adj, exhibit36_adj) %>%
  dplyr::filter(AY == "CY_Total" | AY == "PY_Total" | AY == "Total") %>%
  dplyr::filter(PRODUCTS == "Auto" | PRODUCTS =="BIKE" | PRODUCTS =="PA" | PRODUCTS == "PET") %>%
  dplyr::select(BIZ_LINE, PRODUCTS, COV_NAME, CL_Large2, AY, starts_with("EAXA_RES")) %>%
  dplyr::rename(M01_vwtd = EAXA_RES, M02_emm3 = EAXA_RES1, M03_emm5 = EAXA_RES2, M04_emm7 = EAXA_RES3, M05_emm12 = EAXA_RES4, M06_emm24 = EAXA_RES5, M07_emm36 = EAXA_RES6,
                M08_vwtd_adj = EAXA_RES7, M09_emm3_adj = EAXA_RES8, M10_emm5_adj = EAXA_RES9, M11_emm7_adj = EAXA_RES10, M12_emm12_adj = EAXA_RES11, M13_emm24_adj = EAXA_RES12, M14_emm36_adj = EAXA_RES13) %>%
  tidyr::gather(key = Model, value = EAXA_Res, M01_vwtd, M02_emm3, M03_emm5, M04_emm7, M05_emm12, M06_emm24, M07_emm36, M08_vwtd_adj, M09_emm3_adj, M10_emm5_adj, M11_emm7_adj, M12_emm12_adj, M13_emm24_adj, M14_emm36_adj)

eaxa_res_plot_all <- ggplot(eaxa_res_by_model,
                            aes(x = AY, y = EAXA_Res/10^6, fill = Model)) +
  geom_bar(stat = "identity", position = "dodge") +
  facet_wrap( ~ PRODUCTS+COV_NAME+CL_Large2, scales = "free") +
  theme(axis.text.x = element_text(angle = 90, hjust = 1)) +
  scale_y_continuous(name="EAXA Reserve (mJPY)", labels = scales::comma)
print(eaxa_res_plot_all)
#
#
#Auto BIL Small
product <- "Auto"
cov <- "BIL"
cl_type <- "Attrit"
mdl_type <- "exhibit36" # PCFY20 no chnage
lkr_type <- "lkr36"
tri_begin_ym <- "199909"
auto_bil_sc_ldf_plot <- Reserve_Data$lkr_plot[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_bil_sc_er_models_plot <- Eaxa_res_plot_by_model(eaxa_res_by_model, product = product, cov = cov, cl_type = cl_type)
auto_bil_sc_exhibit <- exhibit_by_cov(mdl_type, product = product, cov = cov, cl_type = cl_type) # keep emm36 HY20
auto_bil_sc_fulltri <- get_full_triangle(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
auto_bil_sc_lkr <- Reserve_Data[, lkr_type][Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_bil_sc_tri <- Reserve_Data$inc_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_bil_sc_lkr_tri <- Reserve_Data$lkr_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_bil_sc_clcnt_tri <- Reserve_Data$clcnt_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_bil_sc_clcnt_fulltri <- get_full_triangle_clcnt(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
auto_bil_sc_ac_tri <- cbind(auto_bil_sc_tri[, 1], auto_bil_sc_fulltri[, -1] / auto_bil_sc_clcnt_fulltri[, -1])
auto_bil_sc_freq_tri <- make_freq_tri(begin_ym = tri_begin_ym, clcnt_tri = auto_bil_sc_clcnt_fulltri)
auto_bil_sc_lr_tri <- make_lr_full_tri(begin_ym = tri_begin_ym, full_tri = auto_bil_sc_fulltri)
auto_bil_sc_bc_tri <- make_bc_tri(begin_ym = tri_begin_ym, inc_tri = auto_bil_sc_fulltri)
auto_bil_sc_fulltri_paid <- get_full_triangle_paid(lkr_type = "lkr12_paid", product = product, cov = cov, cl_type = cl_type)


#Auto BIL Large
product <- "Auto"
cov <- "BIL"
cl_type <- "Large"
mdl_type <- "exhibit36" # PCFY20 24, HY21 FC1 36へ変更必要か
lkr_type <- "lkr36"
tri_begin_ym <- "200101"
auto_bil_lc_ldf_plot <- Reserve_Data$lkr_plot[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_bil_lc_er_models_plot <- Eaxa_res_plot_by_model(eaxa_res_by_model, product = product, cov = cov, cl_type = cl_type)
auto_bil_lc_exhibit <- exhibit_by_cov(mdl_type, product = product, cov = cov, cl_type = cl_type) #mod emm7 to emm36
auto_bil_lc_fulltri <- get_full_triangle(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
auto_bil_lc_lkr <- Reserve_Data[, lkr_type][Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_bil_lc_tri <- Reserve_Data$inc_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_bil_lc_lkr_tri <- Reserve_Data$lkr_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_bil_lc_clcnt_tri <- Reserve_Data$clcnt_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_bil_lc_clcnt_fulltri <- get_full_triangle_clcnt(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
auto_bil_lc_ac_tri <- cbind(auto_bil_lc_tri[, 1], auto_bil_lc_fulltri[, -1] / auto_bil_lc_clcnt_fulltri[, -1])
auto_bil_lc_freq_tri <- make_freq_tri(begin_ym = tri_begin_ym, clcnt_tri = auto_bil_lc_clcnt_fulltri)
auto_bil_lc_lr_tri <- make_lr_full_tri(begin_ym = tri_begin_ym, full_tri = auto_bil_lc_fulltri)
auto_bil_lc_bc_tri <- make_bc_tri(begin_ym = tri_begin_ym, inc_tri = auto_bil_lc_fulltri)
auto_bil_lc_fulltri_paid <- get_full_triangle_paid(lkr_type = "lkr12_paid", product = product, cov = cov, cl_type = cl_type)


#Auto PDL
product <- "Auto"
cov <- "PDL"
cl_type <- "Total"
mdl_type <- "exhibit12" # PCFY20 no chnage
lkr_type <- "lkr12"
tri_begin_ym <- "199908"
auto_pdl_ldf_plot <- Reserve_Data$lkr_plot[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_pdl_er_models_plot <- Eaxa_res_plot_by_model(eaxa_res_by_model, product = product, cov = cov, cl_type = cl_type)
auto_pdl_exhibit <- exhibit_by_cov(mdl_type, product = product, cov = cov, cl_type = cl_type)
auto_pdl_fulltri <- get_full_triangle(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
auto_pdl_lkr <- Reserve_Data[, lkr_type][Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_pdl_tri <- Reserve_Data$inc_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_pdl_lkr_tri <- Reserve_Data$lkr_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_pdl_clcnt_tri <- Reserve_Data$clcnt_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_pdl_clcnt_fulltri <- get_full_triangle_clcnt(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
auto_pdl_ac_tri <- cbind(auto_pdl_tri[, 1], auto_pdl_fulltri[, -1] / auto_pdl_clcnt_fulltri[, -1])
auto_pdl_freq_tri <- make_freq_tri(begin_ym = tri_begin_ym, clcnt_tri = auto_pdl_clcnt_fulltri)
auto_pdl_lr_tri <- make_lr_full_tri(begin_ym = tri_begin_ym, full_tri = auto_pdl_fulltri)
auto_pdl_bc_tri <- make_bc_tri(begin_ym = tri_begin_ym, inc_tri = auto_pdl_fulltri)
auto_pdl_fulltri_paid <- get_full_triangle_paid(lkr_type = "lkr12_paid", product = product, cov = cov, cl_type = cl_type)

#Auto PI Small
product <- "Auto"
cov <- "PI"
cl_type <- "Attrit"
mdl_type <- "exhibit12_adj" # HY21 12から36を検討 
lkr_type <- "lkr12_adj"
tri_begin_ym <- "200104"
auto_pi_sc_ldf_plot <- Reserve_Data$lkr_plot[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_pi_sc_er_models_plot <- Eaxa_res_plot_by_model(eaxa_res_by_model, product = product, cov = cov, cl_type = cl_type)
auto_pi_sc_exhibit <- exhibit_by_cov(mdl_type, product = product, cov = cov, cl_type = cl_type)
auto_pi_sc_fulltri <- get_full_triangle(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)  # keep emm12 HY20
auto_pi_sc_lkr <- Reserve_Data[, lkr_type][Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_pi_sc_tri <- Reserve_Data$inc_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_pi_sc_lkr_tri <- Reserve_Data$lkr_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_pi_sc_clcnt_tri <- Reserve_Data$clcnt_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_pi_sc_clcnt_fulltri <- get_full_triangle_clcnt(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
auto_pi_sc_ac_tri <- cbind(auto_pi_sc_tri[, 1], auto_pi_sc_fulltri[, -1] / auto_pi_sc_clcnt_fulltri[, -1])
auto_pi_sc_freq_tri <- make_freq_tri(begin_ym = tri_begin_ym, clcnt_tri = auto_pi_sc_clcnt_fulltri)
auto_pi_sc_lr_tri <- make_lr_full_tri(begin_ym = tri_begin_ym, full_tri = auto_pi_sc_fulltri)
auto_pi_sc_bc_tri <- make_bc_tri(begin_ym = tri_begin_ym, inc_tri = auto_pi_sc_fulltri)
auto_pi_sc_fulltri_paid <- get_full_triangle_paid(lkr_type = "lkr12_paid", product = product, cov = cov, cl_type = cl_type)

#Auto PI Large
product <- "Auto"
cov <- "PI"
cl_type <- "Large"
mdl_type <- "exhibit12" # PCFY20 no chnage
lkr_type <- "lkr12"
tri_begin_ym <- "200205"
auto_pi_lc_ldf_plot <- Reserve_Data$lkr_plot[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_pi_lc_er_models_plot <- Eaxa_res_plot_by_model(eaxa_res_by_model, product = product, cov = cov, cl_type = cl_type)
auto_pi_lc_exhibit <- exhibit_by_cov(mdl_type, product = product, cov = cov, cl_type = cl_type)
auto_pi_lc_fulltri <- get_full_triangle(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
auto_pi_lc_lkr <- Reserve_Data[, lkr_type][Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_pi_lc_tri <- Reserve_Data$inc_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_pi_lc_lkr_tri <- Reserve_Data$lkr_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_pi_lc_clcnt_tri <- Reserve_Data$clcnt_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_pi_lc_clcnt_fulltri <- get_full_triangle_clcnt(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
auto_pi_lc_ac_tri <- cbind(auto_pi_lc_tri[, 1], auto_pi_lc_fulltri[, -1] / auto_pi_lc_clcnt_fulltri[, -1])
auto_pi_lc_freq_tri <- make_freq_tri(begin_ym = tri_begin_ym, clcnt_tri = auto_pi_lc_clcnt_fulltri)
auto_pi_lc_lr_tri <- make_lr_full_tri(begin_ym = tri_begin_ym, full_tri = auto_pi_lc_fulltri)
auto_pi_lc_bc_tri <- make_bc_tri(begin_ym = tri_begin_ym, inc_tri = auto_pi_lc_fulltri)
auto_pi_lc_fulltri_paid <- get_full_triangle_paid(lkr_type = "lkr12_paid", product = product, cov = cov, cl_type = cl_type)

#Auto PPA
product <- "Auto"
cov <- "PPA"
cl_type <- "Total"
mdl_type <- "exhibit12" # PCFY20 no chnage
lkr_type <- "lkr12"
tri_begin_ym <- "199909"
auto_ppa_ldf_plot <- Reserve_Data$lkr_plot[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_ppa_er_models_plot <- Eaxa_res_plot_by_model(eaxa_res_by_model, product = product, cov = cov, cl_type = cl_type)
auto_ppa_exhibit <- exhibit_by_cov(mdl_type, product = product, cov = cov, cl_type = cl_type)
auto_ppa_fulltri <- get_full_triangle(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
auto_ppa_lkr <- Reserve_Data[, lkr_type][Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_ppa_tri <- Reserve_Data$inc_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_ppa_lkr_tri <- Reserve_Data$lkr_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_ppa_clcnt_tri <- Reserve_Data$clcnt_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_ppa_clcnt_fulltri <- get_full_triangle_clcnt(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
auto_ppa_ac_tri <- cbind(auto_ppa_tri[, 1], auto_ppa_fulltri[, -1] / auto_ppa_clcnt_fulltri[, -1])
auto_ppa_freq_tri <- make_freq_tri(begin_ym = tri_begin_ym, clcnt_tri = auto_ppa_clcnt_fulltri)
auto_ppa_lr_tri <- make_lr_full_tri(begin_ym = tri_begin_ym, full_tri = auto_ppa_fulltri)
auto_ppa_bc_tri <- make_bc_tri(begin_ym = tri_begin_ym, inc_tri = auto_ppa_fulltri)
auto_ppa_fulltri_paid <- get_full_triangle_paid(lkr_type = "lkr12_paid", product = product, cov = cov, cl_type = cl_type)

#Auto ODPE
product <- "Auto"
cov <- "ODPE"
cl_type <- "Attrit"
mdl_type <- "exhibit7" # PCFY20 no chnage
lkr_type <- "lkr7"
tri_begin_ym <- "199908"
auto_odpe_ldf_plot <- Reserve_Data$lkr_plot[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_odpe_er_models_plot <- Eaxa_res_plot_by_model(eaxa_res_by_model, product = product, cov = cov, cl_type = cl_type)
auto_odpe_exhibit <- exhibit_by_cov(mdl_type, product = product, cov = cov, cl_type = cl_type)
auto_odpe_fulltri <- get_full_triangle(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)  #mod emm12 to emm7 HY20
auto_odpe_lkr <- Reserve_Data[, lkr_type][Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_odpe_tri <- Reserve_Data$inc_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_odpe_lkr_tri <- Reserve_Data$lkr_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_odpe_clcnt_tri <- Reserve_Data$clcnt_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_odpe_clcnt_fulltri <- get_full_triangle_clcnt(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
auto_odpe_ac_tri <- cbind(auto_odpe_tri[, 1], auto_odpe_fulltri[, -1] / auto_odpe_clcnt_fulltri[, -1])
auto_odpe_freq_tri <- make_freq_tri(begin_ym = tri_begin_ym, clcnt_tri = auto_odpe_clcnt_fulltri)
auto_odpe_lr_tri <- make_lr_full_tri(begin_ym = tri_begin_ym, full_tri = auto_odpe_fulltri)
auto_odpe_bc_tri <- make_bc_tri(begin_ym = tri_begin_ym, inc_tri = auto_odpe_fulltri)
auto_odpe_fulltri_paid <- get_full_triangle_paid(lkr_type = "lkr12_paid", product = product, cov = cov, cl_type = cl_type)

#Auto ODPE Event
product <- "Auto"
cov <- "ODPE"
cl_type <- "Evented"
mdl_type <- "exhibit12"  # PCFY20 no chnage
lkr_type <- "lkr12"
tri_begin_ym <- "200707"
auto_odpe_ev_ldf_plot <- Reserve_Data$lkr_plot[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_odpe_ev_er_models_plot <- Eaxa_res_plot_by_model(eaxa_res_by_model, product = product, cov = cov, cl_type = cl_type)
auto_odpe_ev_exhibit <- exhibit_by_cov(mdl_type, product = product, cov = cov, cl_type = cl_type)
auto_odpe_ev_fulltri <- get_full_triangle(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)  # keep emm12 HY20
auto_odpe_ev_lkr <- Reserve_Data[, lkr_type][Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_odpe_ev_tri <- Reserve_Data$inc_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_odpe_ev_lkr_tri <- Reserve_Data$lkr_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_odpe_ev_clcnt_tri <- Reserve_Data$clcnt_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_odpe_ev_clcnt_fulltri <- get_full_triangle_clcnt(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
auto_odpe_ev_ac_tri <- cbind(auto_odpe_ev_tri[, 1], auto_odpe_ev_fulltri[, -1] / auto_odpe_ev_clcnt_fulltri[, -1])
auto_odpe_ev_freq_tri <- make_freq_tri(begin_ym = tri_begin_ym, clcnt_tri = auto_odpe_ev_clcnt_fulltri)
auto_odpe_ev_lr_tri <- make_lr_full_tri(begin_ym = tri_begin_ym, full_tri = auto_odpe_ev_fulltri)
auto_odpe_ev_bc_tri <- make_bc_tri(begin_ym = tri_begin_ym, inc_tri = auto_odpe_ev_fulltri)
auto_odpe_ev_fulltri_paid <- get_full_triangle_paid(lkr_type = "lkr12_paid", product = product, cov = cov, cl_type = cl_type)

#Auto LE
product <- "Auto"
cov <- "LawyerE"
cl_type <- "Total"
mdl_type <- "exhibit36_adj" # PCFY20 24 HY21 FC1 36とするがまだ少ないか
lkr_type <- "lkr36_adj"
tri_begin_ym <- "200309"
auto_le_ldf_plot <- Reserve_Data$lkr_plot[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_le_er_models_plot <- Eaxa_res_plot_by_model(eaxa_res_by_model, product = product, cov = cov, cl_type = cl_type)
auto_le_exhibit <- exhibit_by_cov(mdl_type, product = product, cov = cov, cl_type = cl_type)
auto_le_fulltri <- get_full_triangle(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)  #mod emm12 to emm24
auto_le_lkr <- Reserve_Data[, lkr_type][Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_le_tri <- Reserve_Data$inc_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_le_lkr_tri <- Reserve_Data$lkr_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_le_clcnt_tri <- Reserve_Data$clcnt_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_le_clcnt_fulltri <- get_full_triangle_clcnt(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
auto_le_ac_tri <- cbind(auto_le_tri[, 1], auto_le_fulltri[, -1] / auto_le_clcnt_fulltri[, -1])
auto_le_freq_tri <- make_freq_tri(begin_ym = tri_begin_ym, clcnt_tri = auto_le_clcnt_fulltri)
auto_le_lr_tri <- make_lr_full_tri(begin_ym = tri_begin_ym, full_tri = auto_le_fulltri)
auto_le_bc_tri <- make_bc_tri(begin_ym = tri_begin_ym, inc_tri = auto_le_fulltri)
auto_le_fulltri_paid <- get_full_triangle_paid(lkr_type = "lkr12_paid", product = product, cov = cov, cl_type = cl_type)

#Auto FB
product <- "Auto"
cov <- "FamilyBIKE"
cl_type <- "Total"
mdl_type <- "exhibit36" # PCFY20 12
lkr_type <- "lkr36"
tri_begin_ym <- "200305"
auto_fb_ldf_plot <- Reserve_Data$lkr_plot[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_fb_er_models_plot <- Eaxa_res_plot_by_model(eaxa_res_by_model, product = product, cov = cov, cl_type = cl_type)
auto_fb_exhibit <- exhibit_by_cov(mdl_type, product = product, cov = cov, cl_type = cl_type)
auto_fb_fulltri <- get_full_triangle(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
auto_fb_lkr <- Reserve_Data[, lkr_type][Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_fb_tri <- Reserve_Data$inc_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_fb_lkr_tri <- Reserve_Data$lkr_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_fb_clcnt_tri <- Reserve_Data$clcnt_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_fb_clcnt_fulltri <- get_full_triangle_clcnt(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
auto_fb_ac_tri <- cbind(auto_fb_tri[, 1], auto_fb_fulltri[, -1] / auto_fb_clcnt_fulltri[, -1])
auto_fb_freq_tri <- make_freq_tri(begin_ym = tri_begin_ym, clcnt_tri = auto_fb_clcnt_fulltri)
auto_fb_lr_tri <- make_lr_full_tri(begin_ym = tri_begin_ym, full_tri = auto_fb_fulltri)
auto_fb_bc_tri <- make_bc_tri(begin_ym = tri_begin_ym, inc_tri = auto_fb_fulltri)
auto_fb_fulltri_paid <- get_full_triangle_paid(lkr_type = "lkr12_paid", product = product, cov = cov, cl_type = cl_type)

#Auto AXAP
product <- "Auto"
cov <- "AXA+"
cl_type <- "Total"
mdl_type <- "exhibit12" # HY21 12
lkr_type <- "lkr12"
tri_begin_ym <- "200501"
auto_axap_ldf_plot <- Reserve_Data$lkr_plot[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_axap_er_models_plot <- Eaxa_res_plot_by_model(eaxa_res_by_model, product = product, cov = cov, cl_type = cl_type)
auto_axap_exhibit <- exhibit_by_cov(mdl_type, product = product, cov = cov, cl_type = cl_type)
auto_axap_fulltri <- get_full_triangle(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
auto_axap_lkr <- Reserve_Data[, lkr_type][Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_axap_tri <- Reserve_Data$inc_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_axap_lkr_tri <- Reserve_Data$lkr_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_axap_clcnt_tri <- Reserve_Data$clcnt_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_axap_clcnt_fulltri <- get_full_triangle_clcnt(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
auto_axap_ac_tri <- cbind(auto_axap_tri[, 1], auto_axap_fulltri[, -1] / auto_axap_clcnt_fulltri[, -1])
auto_axap_freq_tri <- make_freq_tri(begin_ym = tri_begin_ym, clcnt_tri = auto_axap_clcnt_fulltri)
auto_axap_lr_tri <- make_lr_full_tri(begin_ym = tri_begin_ym, full_tri = auto_axap_fulltri)
auto_axap_bc_tri <- make_bc_tri(begin_ym = tri_begin_ym, inc_tri = auto_axap_fulltri)
auto_axap_fulltri_paid <- get_full_triangle_paid(lkr_type = "lkr12_paid", product = product, cov = cov, cl_type = cl_type)

#Auto EQ
product <- "Auto"
cov <- "OD_EQ"
cl_type <- "Total"
mdl_type <- "exhibit12" # PCFY20 no chnage
lkr_type <- "lkr12"
tri_begin_ym <- "201604"
auto_eq_ldf_plot <- Reserve_Data$lkr_plot[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_eq_er_models_plot <- Eaxa_res_plot_by_model(eaxa_res_by_model, product = product, cov = cov, cl_type = cl_type)
auto_eq_exhibit <- exhibit_by_cov(mdl_type, product = product, cov = cov, cl_type = cl_type)
auto_eq_exhibit <- eaxa_exhibit(c("1999", "2000", "2001", "2002", "2003", "2004", "2005", "2006",
                                  "2007", "2008", "2009", "2010", "2011", "2012"), auto_eq_exhibit)
auto_eq_fulltri <- get_full_triangle(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
auto_eq_lkr <- Reserve_Data[, lkr_type][Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_eq_tri <- Reserve_Data$inc_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_eq_lkr_tri <- Reserve_Data$lkr_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_eq_clcnt_tri <- Reserve_Data$clcnt_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
auto_eq_clcnt_fulltri <- get_full_triangle_clcnt(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
auto_eq_ac_tri <- cbind(auto_eq_tri[, 1], auto_eq_fulltri[, -1] / auto_eq_clcnt_fulltri[, -1])
auto_eq_freq_tri <- make_freq_tri(begin_ym = tri_begin_ym, clcnt_tri = auto_eq_clcnt_fulltri)
auto_eq_lr_tri <- make_lr_full_tri(begin_ym = tri_begin_ym, full_tri = auto_eq_fulltri)
auto_eq_bc_tri <- make_bc_tri(begin_ym = tri_begin_ym, inc_tri = auto_eq_fulltri)
auto_eq_fulltri_paid <- get_full_triangle_paid(lkr_type = "lkr12_paid", product = product, cov = cov, cl_type = cl_type)

#Bike BIL Small
product <- "BIKE"
cov <- "BIL"
cl_type <- "Attrit"
mdl_type <- "exhibit12" #PCFY23 12_adjから12に変更 
lkr_type <- "lkr12"
tri_begin_ym <- "200505"
bike_bil_sc_ldf_plot <- Reserve_Data$lkr_plot[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_bil_sc_er_models_plot <- Eaxa_res_plot_by_model(eaxa_res_by_model, product = product, cov = cov, cl_type = cl_type)
bike_bil_sc_exhibit <- exhibit_by_cov(mdl_type, product = product, cov = cov, cl_type = cl_type)
bike_bil_sc_fulltri <- get_full_triangle(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
bike_bil_sc_lkr <- Reserve_Data[, lkr_type][Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_bil_sc_tri <- Reserve_Data$inc_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_bil_sc_lkr_tri <- Reserve_Data$lkr_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_bil_sc_clcnt_tri <- Reserve_Data$clcnt_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_bil_sc_clcnt_fulltri <- get_full_triangle_clcnt(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
bike_bil_sc_ac_tri <- cbind(bike_bil_sc_tri[, 1], bike_bil_sc_fulltri[, -1] / bike_bil_sc_clcnt_fulltri[, -1])
bike_bil_sc_freq_tri <- make_freq_tri(begin_ym = tri_begin_ym, clcnt_tri = bike_bil_sc_clcnt_fulltri)
bike_bil_sc_lr_tri <- make_lr_full_tri(begin_ym = tri_begin_ym, full_tri = bike_bil_sc_fulltri)
bike_bil_sc_bc_tri <- make_bc_tri(begin_ym = tri_begin_ym, inc_tri = bike_bil_sc_fulltri)
bike_bil_sc_fulltri_paid <- get_full_triangle_paid(lkr_type = "lkr12_paid", product = product, cov = cov, cl_type = cl_type)

#Bike BIL Large
product <- "BIKE"
cov <- "BIL"
cl_type <- "Large"
mdl_type <- "exhibit_vwtd_adj" # PCFY20 36
lkr_type <- "lkr_vwtd_adj"
tri_begin_ym <- "200612"
bike_bil_lc_ldf_plot <- Reserve_Data$lkr_plot[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_bil_lc_er_models_plot <- Eaxa_res_plot_by_model(eaxa_res_by_model, product = product, cov = cov, cl_type = cl_type)
bike_bil_lc_exhibit <- exhibit_by_cov(mdl_type, product = product, cov = cov, cl_type = cl_type)
bike_bil_lc_fulltri <- get_full_triangle(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
bike_bil_lc_lkr <- Reserve_Data[, lkr_type][Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_bil_lc_tri <- Reserve_Data$inc_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_bil_lc_lkr_tri <- Reserve_Data$lkr_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_bil_lc_clcnt_tri <- Reserve_Data$clcnt_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_bil_lc_clcnt_fulltri <- get_full_triangle_clcnt(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
bike_bil_lc_ac_tri <- cbind(bike_bil_lc_tri[, 1], bike_bil_lc_fulltri[, -1] / bike_bil_lc_clcnt_fulltri[, -1])
bike_bil_lc_freq_tri <- make_freq_tri(begin_ym = tri_begin_ym, clcnt_tri = bike_bil_lc_clcnt_fulltri)
bike_bil_lc_lr_tri <- make_lr_full_tri(begin_ym = tri_begin_ym, full_tri = bike_bil_lc_fulltri)
bike_bil_lc_bc_tri <- make_bc_tri(begin_ym = tri_begin_ym, inc_tri = bike_bil_lc_fulltri)
bike_bil_lc_fulltri_paid <- get_full_triangle_paid(lkr_type = "lkr12_paid", product = product, cov = cov, cl_type = cl_type)

#Bike PDL
product <- "BIKE"
cov <- "PDL"
cl_type <- "Total"
mdl_type <- "exhibit12" # PCFY20 no chnage
lkr_type <- "lkr12"
tri_begin_ym <- "200505"
bike_pdl_ldf_plot <- Reserve_Data$lkr_plot[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_pdl_er_models_plot <- Eaxa_res_plot_by_model(eaxa_res_by_model, product = product, cov = cov, cl_type = cl_type)
bike_pdl_exhibit <- exhibit_by_cov(mdl_type, product = product, cov = cov, cl_type = cl_type)
bike_pdl_fulltri <- get_full_triangle(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
bike_pdl_lkr <- Reserve_Data[, lkr_type][Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_pdl_tri <- Reserve_Data$inc_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_pdl_lkr_tri <- Reserve_Data$lkr_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_pdl_clcnt_tri <- Reserve_Data$clcnt_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_pdl_clcnt_fulltri <- get_full_triangle_clcnt(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
bike_pdl_ac_tri <- cbind(bike_pdl_tri[, 1], bike_pdl_fulltri[, -1] / bike_pdl_clcnt_fulltri[, -1])
bike_pdl_freq_tri <- make_freq_tri(begin_ym = tri_begin_ym, clcnt_tri = bike_pdl_clcnt_fulltri)
bike_pdl_lr_tri <- make_lr_full_tri(begin_ym = tri_begin_ym, full_tri = bike_pdl_fulltri)
bike_pdl_bc_tri <- make_bc_tri(begin_ym = tri_begin_ym, inc_tri = bike_pdl_fulltri)
bike_pdl_fulltri_paid <- get_full_triangle_paid(lkr_type = "lkr12_paid", product = product, cov = cov, cl_type = cl_type)

#Bike PI Small
product <- "BIKE"
cov <- "PI"
cl_type <- "Attrit"
mdl_type <- "exhibit24" # HY21 emm24に変更
lkr_type <- "lkr24"
tri_begin_ym <- "200811"
bike_pi_sc_ldf_plot <- Reserve_Data$lkr_plot[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_pi_sc_er_models_plot <- Eaxa_res_plot_by_model(eaxa_res_by_model, product = product, cov = cov, cl_type = cl_type)
bike_pi_sc_exhibit <- exhibit_by_cov(mdl_type, product = product, cov = cov, cl_type = cl_type)
bike_pi_sc_fulltri <- get_full_triangle(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
bike_pi_sc_lkr <- Reserve_Data[, lkr_type][Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_pi_sc_lkr <- Reserve_Data[, lkr_type][Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_pi_sc_tri <- Reserve_Data$inc_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_pi_sc_lkr_tri <- Reserve_Data$lkr_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_pi_sc_clcnt_tri <- Reserve_Data$clcnt_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_pi_sc_clcnt_fulltri <- get_full_triangle_clcnt(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
bike_pi_sc_ac_tri <- cbind(bike_pi_sc_tri[, 1], bike_pi_sc_fulltri[, -1] / bike_pi_sc_clcnt_fulltri[, -1])
bike_pi_sc_freq_tri <- make_freq_tri(begin_ym = tri_begin_ym, clcnt_tri = bike_pi_sc_clcnt_fulltri)
bike_pi_sc_lr_tri <- make_lr_full_tri(begin_ym = tri_begin_ym, full_tri = bike_pi_sc_fulltri)
bike_pi_sc_bc_tri <- make_bc_tri(begin_ym = tri_begin_ym, inc_tri = bike_pi_sc_fulltri)
bike_pi_sc_fulltri_paid <- get_full_triangle_paid(lkr_type = "lkr12_paid", product = product, cov = cov, cl_type = cl_type)

#Bike PI Large
product <- "BIKE"
cov <- "PI"
cl_type <- "Large"
mdl_type <- "exhibit_vwtd_adj" # PCFY20 12
lkr_type <- "lkr_vwtd_adj"
tri_begin_ym <- "201306" # mod 202103
bike_pi_lc_ldf_plot <- Reserve_Data$lkr_plot[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_pi_lc_er_models_plot <- Eaxa_res_plot_by_model(eaxa_res_by_model, product = product, cov = cov, cl_type = cl_type)
bike_pi_lc_exhibit <- exhibit_by_cov(mdl_type, product = product, cov = cov, cl_type = cl_type)
bike_pi_lc_fulltri <- get_full_triangle(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
bike_pi_lc_lkr <- Reserve_Data[, lkr_type][Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_pi_lc_tri <- Reserve_Data$inc_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_pi_lc_lkr_tri <- Reserve_Data$lkr_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_pi_lc_clcnt_tri <- Reserve_Data$clcnt_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_pi_lc_clcnt_fulltri <- get_full_triangle_clcnt(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
bike_pi_lc_ac_tri <- cbind(bike_pi_lc_tri[, 1], bike_pi_lc_fulltri[, -1] / bike_pi_lc_clcnt_fulltri[, -1])
bike_pi_lc_freq_tri <- make_freq_tri(begin_ym = tri_begin_ym, clcnt_tri = bike_pi_lc_clcnt_fulltri)
bike_pi_lc_lr_tri <- make_lr_full_tri(begin_ym = tri_begin_ym, full_tri = bike_pi_lc_fulltri)
bike_pi_lc_bc_tri <- make_bc_tri(begin_ym = tri_begin_ym, inc_tri = bike_pi_lc_fulltri)
bike_pi_lc_fulltri_paid <- get_full_triangle_paid(lkr_type = "lkr12_paid", product = product, cov = cov, cl_type = cl_type)

#Bike PPA
product <- "BIKE"
cov <- "PPA"
cl_type <- "Total"
mdl_type <- "exhibit36" # PCFY20 12
lkr_type <- "lkr36"
tri_begin_ym <- "200505"
bike_ppa_ldf_plot <- Reserve_Data$lkr_plot[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_ppa_er_models_plot <- Eaxa_res_plot_by_model(eaxa_res_by_model, product = product, cov = cov, cl_type = cl_type)
bike_ppa_exhibit <- exhibit_by_cov(mdl_type, product = product, cov = cov, cl_type = cl_type)
bike_ppa_fulltri <- get_full_triangle(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
bike_ppa_lkr <- Reserve_Data[, lkr_type][Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_ppa_tri <- Reserve_Data$inc_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_ppa_lkr_tri <- Reserve_Data$lkr_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_ppa_clcnt_tri <- Reserve_Data$clcnt_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_ppa_clcnt_fulltri <- get_full_triangle_clcnt(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
bike_ppa_ac_tri <- cbind(bike_ppa_tri[, 1], bike_ppa_fulltri[, -1] / bike_ppa_clcnt_fulltri[, -1])
bike_ppa_freq_tri <- make_freq_tri(begin_ym = tri_begin_ym, clcnt_tri = bike_ppa_clcnt_fulltri)
bike_ppa_lr_tri <- make_lr_full_tri(begin_ym = tri_begin_ym, full_tri = bike_ppa_fulltri)
bike_ppa_bc_tri <- make_bc_tri(begin_ym = tri_begin_ym, inc_tri = bike_ppa_fulltri)
bike_ppa_fulltri_paid <- get_full_triangle_paid(lkr_type = "lkr12_paid", product = product, cov = cov, cl_type = cl_type)

#Bike LE
product <- "BIKE"
cov <- "LawyerE"
cl_type <- "Total"
mdl_type <- "exhibit36" #PCFY 12
lkr_type <- "lkr36"
tri_begin_ym <- "200507"
bike_le_ldf_plot <- Reserve_Data$lkr_plot[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_le_er_models_plot <- Eaxa_res_plot_by_model(eaxa_res_by_model, product = product, cov = cov, cl_type = cl_type)
bike_le_exhibit <- exhibit_by_cov(mdl_type, product = product, cov = cov, cl_type = cl_type)
bike_le_fulltri <- get_full_triangle(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
bike_le_lkr <- Reserve_Data[, lkr_type][Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_le_tri <- Reserve_Data$inc_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_le_lkr_tri <- Reserve_Data$lkr_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_le_clcnt_tri <- Reserve_Data$clcnt_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
bike_le_clcnt_fulltri <- get_full_triangle_clcnt(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
bike_le_ac_tri <- cbind(bike_le_tri[, 1], bike_le_fulltri[, -1] / bike_le_clcnt_fulltri[, -1])
bike_le_freq_tri <- make_freq_tri(begin_ym = tri_begin_ym, clcnt_tri = bike_le_clcnt_fulltri)
bike_le_lr_tri <- make_lr_full_tri(begin_ym = tri_begin_ym, full_tri = bike_le_fulltri)
bike_le_bc_tri <- make_bc_tri(begin_ym = tri_begin_ym, inc_tri = bike_le_fulltri)
bike_le_fulltri_paid <- get_full_triangle_paid(lkr_type = "lkr12_paid", product = product, cov = cov, cl_type = cl_type)

#PA
product <- "PA"
cov <- "Total"
cl_type <- "Attrit"
mdl_type <- "exhibit12_adj" # PCFY23 lkr12に変更
lkr_type <- "lkr12_adj"
tri_begin_ym <- "200105"
pa_ldf_plot <- Reserve_Data$lkr_plot[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
pa_er_models_plot <- Eaxa_res_plot_by_model(eaxa_res_by_model, product = product, cov = cov, cl_type = cl_type)
pa_exhibit <- exhibit_by_cov(mdl_type, product = product, cov = cov, cl_type = cl_type)
pa_exhibit <- eaxa_exhibit(c("1999", "2000"), pa_exhibit)
pa_exhibit <- format_exhibit(pa_exhibit)
pa_fulltri <- get_full_triangle(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
pa_lkr <- Reserve_Data[, lkr_type][Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
pa_tri <- Reserve_Data$inc_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
pa_lkr_tri <- Reserve_Data$lkr_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
pa_clcnt_tri <- Reserve_Data$clcnt_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
pa_clcnt_fulltri <- get_full_triangle_clcnt(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
pa_ac_tri <- cbind(pa_tri[, 1], pa_fulltri[, -1] / pa_clcnt_fulltri[, -1])
pa_freq_tri <- make_freq_tri(begin_ym = tri_begin_ym, clcnt_tri = pa_clcnt_fulltri)
pa_lr_tri <- make_lr_full_tri(begin_ym = tri_begin_ym, full_tri = pa_fulltri)
pa_bc_tri <- make_bc_tri(begin_ym = tri_begin_ym, inc_tri = pa_fulltri)
pa_fulltri_paid <- get_full_triangle_paid(lkr_type = "lkr12_paid", product = product, cov = cov, cl_type = cl_type)

#Pet
product <- "PET"
cov <- "Total"
cl_type <- "Attrit"
mdl_type <- "exhibit12" # PCFY20 no chnage
lkr_type <- "lkr12"
tri_begin_ym <- "200604"
pet_ldf_plot <- Reserve_Data$lkr_plot[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
pet_er_models_plot <- Eaxa_res_plot_by_model(eaxa_res_by_model, product = product, cov = cov, cl_type = cl_type)
pet_exhibit <- exhibit_by_cov(mdl_type, product = product, cov = cov, cl_type = cl_type)
pet_exhibit <- eaxa_exhibit(c("1999", "2000", "2001", "2002", "2003", "2004", "2005", "2006", "2007", "2008", "2009"), pet_exhibit)
pet_exhibit <- format_exhibit(pet_exhibit)
pet_fulltri <- get_full_triangle(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
pet_lkr <- Reserve_Data[, lkr_type][Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
pet_tri <- Reserve_Data$inc_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
pet_lkr_tri <- Reserve_Data$lkr_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
pet_clcnt_tri <- Reserve_Data$clcnt_tri[Reserve_Data$PRODUCTS==product & Reserve_Data$COV_NAME==cov & Reserve_Data$CL_Large2==cl_type][[1]]
pet_clcnt_fulltri <- get_full_triangle_clcnt(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type)
pet_ac_tri <- cbind(pet_tri[, 1], pet_fulltri[, -1] / pet_clcnt_fulltri[, -1])
pet_freq_tri <- make_freq_tri_pet(begin_ym = tri_begin_ym, clcnt_tri = pet_clcnt_fulltri) # *Pet specific function
pet_lr_tri <- make_lr_full_tri_pet(begin_ym = tri_begin_ym, full_tri = pet_fulltri) # *Pet specific function
pet_bc_tri <- make_bc_tri_pet(begin_ym = tri_begin_ym, inc_tri = pet_fulltri) # *Pet specific function
pet_fulltri_paid <- get_full_triangle_paid(lkr_type = "lkr12_paid", product = product, cov = cov, cl_type = cl_type)
#
# End of Projection section ----------------------------------------------------
#

# Aggregate --------------------------------------------------------------------

#Auto BIL
auto_bil_exhibit <- add_sc_lc(auto_bil_sc_exhibit, auto_bil_lc_exhibit)

#Auto PI
auto_pi_exhibit <- add_sc_lc(auto_pi_sc_exhibit, auto_pi_lc_exhibit)

#Auto ODPE Total
auto_od_pe_ev_exhibit <- add_sc_lc(auto_odpe_exhibit, auto_odpe_ev_exhibit)

#Auto Total
drops <- c("AY", "EXPOSURE", "FF_LR", "EAXA_LR", "FREQ", "AVG_COST", "AVG_EP", "EAXA_FREQ", "EAXA_AVG_COST")
auto_exhibit <- cbind(
  AY=auto_bil_exhibit[, "AY"],
  (
    auto_bil_exhibit[, !(names(auto_bil_exhibit) %in% drops)] +
      auto_pi_exhibit[, !(names(auto_pi_exhibit) %in% drops)] +
      auto_pdl_exhibit[, !(names(auto_pdl_exhibit) %in% drops)] +
      auto_ppa_exhibit[, !(names(auto_ppa_exhibit) %in% drops)] +
      auto_od_pe_ev_exhibit[, !(names(auto_od_pe_ev_exhibit) %in% drops)] +
      auto_fb_exhibit[, !(names(auto_fb_exhibit) %in% drops)] +
      auto_le_exhibit[, !(names(auto_le_exhibit) %in% drops)] +
      auto_axap_exhibit[, !(names(auto_axap_exhibit) %in% drops)] +
      auto_eq_exhibit[, !(names(auto_eq_exhibit) %in% drops)]
  ),
  EXPOSURE = auto_bil_exhibit$EXPOSURE)
auto_exhibit <- stats_func(auto_exhibit)
auto_exhibit <- format_exhibit(auto_exhibit)

#Bike###

#Bike BIL
bike_bil_exhibit <- add_sc_lc(bike_bil_sc_exhibit, bike_bil_lc_exhibit)

#Bike PI
bike_pi_exhibit <- add_sc_lc(bike_pi_sc_exhibit, bike_pi_lc_exhibit)

#Bike Total
drops <- c("AY", "EXPOSURE", "FF_LR", "EAXA_LR", "FREQ", "AVG_COST","AVG_EP", "EAXA_FREQ", "EAXA_AVG_COST")
bike_exhibit <- cbind(
  AY=bike_bil_exhibit[, "AY"],
  (
    bike_bil_exhibit[, !(names(bike_bil_exhibit) %in% drops)] +
      bike_pi_exhibit[, !(names(bike_pi_exhibit) %in% drops)] +
      bike_pdl_exhibit[, !(names(bike_pdl_exhibit) %in% drops)] +
      bike_ppa_exhibit[, !(names(bike_ppa_exhibit) %in% drops)] +
      bike_le_exhibit[, !(names(bike_le_exhibit) %in% drops)]
  ),
  EXPOSURE = bike_bil_exhibit$EXPOSURE)
bike_exhibit <- stats_func(bike_exhibit)
bike_exhibit <- format_exhibit(bike_exhibit)

#Motor (Auto + Bike)

#Motor BIL Small
motor_bil_sc_exhibit <- add_auto_bike_exhibit(auto_bil_sc_exhibit, bike_bil_sc_exhibit)
#Motor BIL Large
motor_bil_lc_exhibit <- add_auto_bike_exhibit(auto_bil_lc_exhibit, bike_bil_lc_exhibit)
#Motor BIL
motor_bil_exhibit <- add_sc_lc(motor_bil_sc_exhibit, motor_bil_lc_exhibit)
#Motor PDL
motor_pdl_exhibit <- add_auto_bike_exhibit(auto_pdl_exhibit, bike_pdl_exhibit)
#Motor PI Small
motor_pi_sc_exhibit <- add_auto_bike_exhibit(auto_pi_sc_exhibit, bike_pi_sc_exhibit)
#Motor PI Large
motor_pi_lc_exhibit <- add_auto_bike_exhibit(auto_pi_lc_exhibit, bike_pi_lc_exhibit)
#Motor PI
motor_pi_exhibit <- add_sc_lc(motor_pi_sc_exhibit, motor_pi_lc_exhibit)
#Motor PPA
motor_ppa_exhibit <- add_auto_bike_exhibit(auto_ppa_exhibit, bike_ppa_exhibit)
#Motor LawyerE
motor_le_exhibit <- add_auto_bike_exhibit(auto_le_exhibit, bike_le_exhibit)
#Motor Total
drops <- c("AY", "EXPOSURE", "FF_LR", "EAXA_LR", "FREQ", "AVG_COST", "AVG_EP", "EAXA_FREQ", "EAXA_AVG_COST")
motor_exhibit <- cbind(
  AY=motor_bil_exhibit[, "AY"],
  (
    motor_bil_exhibit[, !(names(motor_bil_exhibit) %in% drops)] +
      motor_pi_exhibit[, !(names(motor_pi_exhibit) %in% drops)] +
      motor_pdl_exhibit[, !(names(motor_pdl_exhibit) %in% drops)] +
      motor_ppa_exhibit[, !(names(motor_ppa_exhibit) %in% drops)] +
      auto_od_pe_ev_exhibit[, !(names(auto_od_pe_ev_exhibit) %in% drops)] +
      auto_fb_exhibit[, !(names(auto_fb_exhibit) %in% drops)] +
      motor_le_exhibit[, !(names(motor_le_exhibit) %in% drops)] +
      auto_axap_exhibit[, !(names(auto_axap_exhibit) %in% drops)] +
      auto_eq_exhibit[, !(names(auto_eq_exhibit) %in% drops)]
  ),
  EXPOSURE = motor_bil_exhibit$EXPOSURE)
motor_exhibit <- stats_func(motor_exhibit)
motor_exhibit <- format_exhibit(motor_exhibit)

#Company Total###

company_exhibit <- add_motor_pet_pa_exhibit(motor_exhibit, pet_exhibit, pa_exhibit)


# For Summary -------------------------------------------------------------

#Make dataframe for graph--------------------

#For EAXA Plot---

lst <- list(auto_bil_lc_exhibit, auto_bil_sc_exhibit, auto_bil_exhibit, auto_pi_lc_exhibit, auto_pi_sc_exhibit, auto_pi_exhibit, auto_pdl_exhibit, auto_ppa_exhibit, auto_od_pe_ev_exhibit,
            auto_fb_exhibit, auto_le_exhibit, auto_axap_exhibit, auto_eq_exhibit, auto_exhibit,
            bike_bil_lc_exhibit, bike_bil_sc_exhibit, bike_bil_exhibit, bike_pi_lc_exhibit, bike_pi_sc_exhibit,  bike_pi_exhibit, bike_pdl_exhibit, bike_ppa_exhibit, bike_le_exhibit, bike_exhibit,
            motor_bil_lc_exhibit, motor_bil_sc_exhibit, motor_bil_exhibit, motor_pi_lc_exhibit, motor_pi_sc_exhibit, motor_pi_exhibit, motor_pdl_exhibit, motor_ppa_exhibit, auto_od_pe_ev_exhibit,
            auto_fb_exhibit, motor_le_exhibit, auto_axap_exhibit, auto_eq_exhibit, motor_exhibit,
            pa_exhibit, pet_exhibit, company_exhibit)
lnm <- list("Auto_BIL_LC", "Auto_BIL_SC", "Auto_BIL", "Auto_PI_LC", "Auto_PI_SC", "Auto_PI", "Auto_PDL", "Auto_PPA", "Auto_ODPE",
            "Auto_FB", "Auto_LE", "Auto_Axa+", "Auto_EQ", "Auto",
            "Bike_BIL_LC", "Bike_BIL_SC", "Bike_BIL", "Bike_PI_LC", "Bike_PI_SC", "Bike_PI", "Bike_PDL", "Bike_PPA", "Bike_LE", "Bike",
            "Motor_BIL_LC", "Motor_BIL_SC", "Motor_BIL", "Motor_PI_LC", "Motor_PI_SC", "Motor_PI", "Motor_PDL", "Motor_PPA", "Motor_ODPE",
            "Motor_FB", "Motor_LE", "Motor_Axa+", "Motor_EQ", "Motor",
            "PA", "Pet", "Company")
all_cov <- c()
for(i in seq_along(lst)) {
  df <- lst[[i]] %>%
    dplyr::mutate(cov=lnm[[i]])
  all_cov <- rbind(all_cov, df)
}
all_exhibit <- all_cov %>%
  group_by(cov) %>%
  nest() %>%
  mutate(
    EAXA = data %>% map(calc_eaxa),
    EAXA_RES = data %>% map(get_eres),
    EAXA_LR = data %>% map(get_EAXA_LR),
    FF_LR = data %>% map(get_FF_LR),
    EAXA_FREQ = data %>% map(get_EAXA_FREQ),
    EAXA_AVG_COST = data %>% map(get_EAXA_AVG_COST)
  )

# summary_all table used for Flash report summary
summary_all <- all_exhibit %>%
  unnest(EAXA_RES, EAXA_LR, FF_LR, EAXA_FREQ, EAXA_AVG_COST) %>%
  dplyr::select(-data, -EAXA) %>%
  dplyr::rename(CY_EAXA_RES = CY_Total, PY_EAXA_RES = PY_Total, Tot_EAXA_RES = Total,
                CY_EAXA_LR = CY_Total1, PY_EAXA_LR = PY_Total1, Tot_EAXA_LR = Total1,
                CY_FF_LR = CY_Total2, PY_FF_LR = PY_Total2, Tot_FF_LR = Total2,
                CY_EAXA_FREQ = CY_Total3, PY_EAXA_FREQ = PY_Total3, Tot_EAXA_FREQ = Total3,
                CY_EAXA_AVG_COST = CY_Total4, PY_EAXA_AVG_COST = PY_Total4, Tot_EAXA_AVG_COST = Total4
  )
#summary_all[, 2:4] <- summary_all[, 2:4] %>% map(accounting, 0)
#summary_all[, 5:10] <- summary_all[, 5:10] %>% map(percent, 2)
# ---

df_auto_graph_all <- all_exhibit %>%
  unnest(EAXA) %>%
  dplyr::filter(str_detect(cov,"Auto"))
df_bike_graph_all <- all_exhibit %>%
  unnest(EAXA) %>%
  dplyr::filter(str_detect(cov,"Bike"))
df_motor_graph_all <- all_exhibit %>%
  unnest(EAXA) %>%
  dplyr::filter(str_detect(cov,"Motor"))
df_pa_pet_graph_all <- all_exhibit %>%
  unnest(EAXA) %>%
  dplyr::filter(cov=="PA" | cov=="Pet")
df_company_graph_all <- all_exhibit %>%
  unnest(EAXA) %>%
  dplyr::filter(cov=="Motor" | cov=="Auto" | cov=="Bike" | cov=="PA" | cov=="Pet" | cov=="Company")
df_company_graph_all$cov <- factor(df_company_graph_all$cov,levels = c("Motor", "Auto", "Bike", "PA", "Pet", "Company"))

#For stats plot---

#Auto---
auto_g_stats <- all_exhibit %>%
  dplyr::select(cov, data) %>% ## need to check
  unnest(data) %>%
  dplyr::filter(str_detect(cov,"Auto")) %>%
  dplyr::select(cov, AY, EAXA_LR, FREQ, AVG_COST, AVG_EP)

#Bike---

bike_g_stats <- all_exhibit %>%
  dplyr::select(cov, data) %>% ## need to check
  unnest(data) %>%
  dplyr::filter(str_detect(cov,"Bike")) %>%
  dplyr::select(cov, AY, EAXA_LR, FREQ, AVG_COST, AVG_EP)

#Motor---

motor_g_stats <- all_exhibit %>%
  dplyr::select(cov, data) %>% ## need to check
  unnest(data) %>%
  dplyr::filter(str_detect(cov,"Motor")) %>%
  dplyr::select(cov, AY, EAXA_LR, FREQ, AVG_COST, AVG_EP)

#PA & Pet---

pa_pet_g_stats <- all_exhibit %>%
  dplyr::select(cov, data) %>% ## need to check
  unnest(data) %>%
  dplyr::filter(cov == "PA" | cov == "Pet") %>%
  dplyr::select(cov, AY, EAXA_LR, FREQ, AVG_COST, AVG_EP)

#Company Total---
#Auto, Bike, Motor, PA, Pet, Company

co_g_stats <- all_exhibit %>%
  dplyr::select(cov, data) %>% ## need to check
  unnest(data) %>%
  dplyr::filter(cov == "Motor" | cov == "Auto" | cov == "Bike" | cov == "PA" | cov == "Pet" | cov == "Company") %>%
  dplyr::select(cov, AY, EAXA_LR, FREQ, AVG_COST, AVG_EP)

# Eaxa plot ---------------------------------------------------------------

#Auto
Eaxa_plot_func_comScale(df_auto_graph_all, "Auto", "Auto EAXA (Common Scale)")
Eaxa_plot_func_freeScale(df_auto_graph_all, "Auto EAXA (Free Scale)")
#Bike
Eaxa_plot_func_comScale(df_bike_graph_all, "Bike", "Bike EAXA (Common Scale)")
Eaxa_plot_func_freeScale(df_bike_graph_all, "Bike EAXA (Free Scale)")
#Motor
Eaxa_plot_func_comScale(df_motor_graph_all, "Motor", "Motor EAXA (Common Scale)")
Eaxa_plot_func_freeScale(df_motor_graph_all, "Motor EAXA (Free Scale)")
#PA Pet
Eaxa_plot_func_comScale(df_pa_pet_graph_all, "", "PA & Pet EAXA (Common Scale)")
Eaxa_plot_func_freeScale(df_pa_pet_graph_all, "PA & Pet EAXA (Free Scale)")
#Company
Eaxa_plot_func_comScale(df_company_graph_all, "", "Company EAXA (Common Scale)")
Eaxa_plot_func_freeScale(df_company_graph_all, "Company EAXA (Free Scale)")

#Stats Plot---

#EAXA LR

Eaxa_LR_plot_func(auto_g_stats, "Auto EAXA LR")
Eaxa_LR_plot_func(bike_g_stats, "Bike EAXA LR")
Eaxa_LR_plot_func(motor_g_stats, "Motor EAXA LR")
Eaxa_LR_plot_func(pa_pet_g_stats, "PA & Pet EAXA LR")
Eaxa_LR_plot_func(co_g_stats, "Company EAXA LR")


#Frequency
Freq_plot_func(auto_g_stats, "Auto F/F Freq")
Freq_plot_func(bike_g_stats, "Bike F/F Freq")
Freq_plot_func(motor_g_stats, "Motor F/F Freq")
Freq_plot_func(pa_pet_g_stats, "PA & Pet F/F Freq")
Freq_plot_func(co_g_stats, "Company F/F Freq")

#AVG Cost
#Non Large
AC_sc_plot_func(auto_g_stats, "Auto F/F AVG Cost")
AC_sc_plot_func(bike_g_stats, "Bike F/F AVG Cost")
AC_sc_plot_func(motor_g_stats, "Motor F/F AVG Cost")
AC_sc_plot_func(pa_pet_g_stats, "PA & Pet F/F AVG Cost")
AC_sc_plot_func(co_g_stats, "Company F/F AVG Cost")

#BI and PI Large
AC_lc_plot_func(auto_g_stats, "Auto F/F AVG Cost")
AC_lc_plot_func(bike_g_stats, "Bike F/F AVG Cost")
AC_lc_plot_func(motor_g_stats, "Motor F/F AVG Cost")

#Average EP
Avg_ep_plot_func(auto_g_stats, "Auto AVG EP")
Avg_ep_plot_func(bike_g_stats, "Bike AVG EP")
Avg_ep_plot_func(motor_g_stats, "Motor AVG EP")
Avg_ep_plot_func(pa_pet_g_stats, "PA & Pet AVG EP")
Avg_ep_plot_func(co_g_stats, "Company AVG EP")
#
#
#
#
#
#
# End of report -----------------------------------------------------------
#
#
#
#
#
#
# For Summary -------------------------------------------------------------
#
# set working directory
setwd(workfile_wd)

all_g_stats <- all_exhibit %>%
  dplyr::select(cov, data) %>% ## need to check
  unnest(data) %>%
  dplyr::select(cov, AY, EAXA_LR, FREQ, AVG_COST, AVG_EP, EAXA_RES, EAXA_FREQ, EAXA_AVG_COST, FF_LR)


a <- all_g_stats %>%
  dplyr::select(AY, cov, EAXA_LR) %>%
  spread(key = cov, value = EAXA_LR)

a[, 2:10] <- a[,2:10] %>% map(percent, 2)

sec_res_summary <- all_g_stats %>%
  dplyr::select(AY, cov, EAXA_RES) %>%
  spread(key = cov, value = EAXA_RES)

sec_eres_py_cy <- all_g_stats %>%
  dplyr::select(cov, AY, EAXA_RES) %>%
  dplyr::filter(
                (cov == "Motor" | cov == "Motor_BIL" | cov == "Motor_PDL" | cov == "Motor_PPA" | cov == "Motor_ODPE"
                 | cov == "Motor_PI" | cov == "Motor_FB" | cov == "Motor_LE" | cov == "Motor_Axa+" | cov == "Motor_EQ"
                 | cov == "PA" | cov == "Pet") &
                  (AY == "PY_Total" | AY == "CY_Total")
                ) %>%
  spread(key = AY, value = EAXA_RES) %>%
  dplyr::select(cov, PY_Total, CY_Total)
sec_eres_py_cy <- sec_eres_py_cy[c(1,3,8,10,7,9,5,6,2,4,11,12),]


summary_metrics <- all_g_stats %>%
  dplyr::filter(
    (cov == "Motor" |cov == "Pet" | cov == "PA" | cov == "Company") &
      (AY == "CY_Total" | AY == "PY_Total" | AY == "Total")
    )

# Export linkratio
selected_ldf <- cbind.fill(
  # Auto
  data.table(auto_bil_sc_lkr),
  data.table(auto_bil_lc_lkr),
  data.table(auto_pdl_lkr),
  data.table(auto_ppa_lkr),
  data.table(auto_odpe_lkr),
  data.table(auto_odpe_ev_lkr),
  data.table(auto_pi_sc_lkr),
  data.table(auto_pi_lc_lkr),
  data.table(auto_fb_lkr),
  data.table(auto_le_lkr),
  data.table(auto_axap_lkr),
  data.table(auto_eq_lkr),
  # Bike
  data.table(bike_bil_sc_lkr),
  data.table(bike_bil_lc_lkr),
  data.table(bike_pdl_lkr),
  data.table(bike_ppa_lkr),
  data.table(bike_pi_sc_lkr),
  data.table(bike_pi_lc_lkr),
  data.table(bike_le_lkr),
  # Health
  data.table(pa_lkr),
  # Pet
  data.table(pet_lkr),
  fill = 1.0
)
id <- rownames(selected_ldf)
selected_ldf <- cbind(id, selected_ldf)
l <- list(
          "sec_res_sum" = data.frame(sec_res_summary),
          "sec_eres_pycy" = data.frame(sec_eres_py_cy),
          "sec_LDF" = data.frame(selected_ldf),
          "graph_data" = data.frame(all_g_stats),
          "Exposure_EP" = df_exp_ep
)
write_xlsx(l, "second_opinion_summaries.xlsx")
#write.csv(df_exp_ep, "exp_ep.csv")
selected_ldf.long2 <- selected_ldf %>%
  #dplyr::select(id, auto_pdl_lkr, "auto_ppa_lkr") %>%
  pivot_longer(-id, names_to = "variable", values_to = "value")

#ggplot(selected_ldf.long2, aes(id, value, colour = variable)) + geom_line()


#
# ------Export EAXA Incurred Full Triangles to Excel
#
l <- list(
  "auto_bil_sc"=as.data.frame(auto_bil_sc_fulltri),
  "auto_bil_lc"=as.data.frame(auto_bil_lc_fulltri),
  "auto_pdl"=as.data.frame(auto_pdl_fulltri),
  "auto_ppa"=as.data.frame(auto_ppa_fulltri),
  "auto_odpe"=as.data.frame(auto_odpe_fulltri),
  "auto_odpe_ev"=as.data.frame(auto_odpe_ev_fulltri),
  "auto_pi_sc"=as.data.frame(auto_pi_sc_fulltri),
  "auto_pi_lc"=as.data.frame(auto_pi_lc_fulltri),
  "auto_axap"=as.data.frame(auto_axap_fulltri),
  "auto_fb"=as.data.frame(auto_fb_fulltri),
  "auto_le"=as.data.frame(auto_le_fulltri),
  "auto_eq"=as.data.frame(auto_eq_fulltri),
  "bike_bil_sc"=as.data.frame(bike_bil_sc_fulltri),
  "bike_bil_lc"=as.data.frame(bike_bil_lc_fulltri),
  "bike_pdl"=as.data.frame(bike_pdl_fulltri),
  "bike_ppa"=as.data.frame(bike_ppa_fulltri),
  "bike_pi_sc"=as.data.frame(bike_pi_sc_fulltri),
  "bike_pi_lc"=as.data.frame(bike_pi_lc_fulltri),
  "bike_le"=as.data.frame(bike_le_fulltri),
  "health_total"=as.data.frame(pa_fulltri),
  "pet_total"=as.data.frame(pet_fulltri)
)
write_xlsx(l, "full_triangles.xlsx")
setwd(next_process_working_wd) # 次回のプロセスでActual vs Expectedを作成するためトライアングルをExport
write_xlsx(l, "Previous_full_triangles.xlsx")
setwd(workfile_wd) # 今期のWDに戻す

#
# ------Export linkratio traiangles to Excel
#
l <- list(
  "auto_bil_sc"=as.data.frame(auto_bil_sc_lkr_tri),
  "auto_bil_lc"=as.data.frame(auto_bil_lc_lkr_tri),
  "auto_pdl"=as.data.frame(auto_pdl_lkr_tri),
  "auto_ppa"=as.data.frame(auto_ppa_lkr_tri),
  "auto_odpe"=as.data.frame(auto_odpe_lkr_tri),
  "auto_pi_sc"=as.data.frame(auto_pi_sc_lkr_tri),
  "auto_pi_lc"=as.data.frame(auto_pi_lc_lkr_tri),
  "auto_axap"=as.data.frame(auto_axap_lkr_tri),
  "auto_fb"=as.data.frame(auto_fb_lkr_tri),
  "auto_le"=as.data.frame(auto_le_lkr_tri),
  "auto_eq"=as.data.frame(auto_eq_lkr_tri),
  "bike_bil_sc"=as.data.frame(bike_bil_sc_lkr_tri),
  "bike_bil_lc"=as.data.frame(bike_bil_lc_lkr_tri),
  "bike_pdl"=as.data.frame(bike_pdl_lkr_tri),
  "bike_ppa"=as.data.frame(bike_ppa_lkr_tri),
  "bike_pi_sc"=as.data.frame(bike_pi_sc_lkr_tri),
  "bike_pi_lc"=as.data.frame(bike_pi_lc_lkr_tri),
  "bike_le"=as.data.frame(bike_le_lkr_tri),
  "health_total"=as.data.frame(pa_lkr_tri),
  "pet_total"=as.data.frame(pet_lkr_tri)
)
write_xlsx(l, "linkratio_triangles.xlsx")
#
# ------Export claim count traiangles to Excel
#
l <- list(
  "auto_bil_sc"=as.data.frame(auto_bil_sc_clcnt_fulltri),
  "auto_bil_lc"=as.data.frame(auto_bil_lc_clcnt_fulltri),
  "auto_pdl"=as.data.frame(auto_pdl_clcnt_fulltri),
  "auto_ppa"=as.data.frame(auto_ppa_clcnt_fulltri),
  "auto_odpe"=as.data.frame(auto_odpe_clcnt_fulltri),
  "auto_pi_sc"=as.data.frame(auto_pi_sc_clcnt_fulltri),
  "auto_pi_lc"=as.data.frame(auto_pi_lc_clcnt_fulltri),
  "auto_axap"=as.data.frame(auto_axap_clcnt_fulltri),
  "auto_fb"=as.data.frame(auto_fb_clcnt_fulltri),
  "auto_le"=as.data.frame(auto_le_clcnt_fulltri),
  "auto_eq"=as.data.frame(auto_eq_clcnt_fulltri),
  "bike_bil_sc"=as.data.frame(bike_bil_sc_clcnt_fulltri),
  "bike_bil_lc"=as.data.frame(bike_bil_lc_clcnt_fulltri),
  "bike_pdl"=as.data.frame(bike_pdl_clcnt_fulltri),
  "bike_ppa"=as.data.frame(bike_ppa_clcnt_fulltri),
  "bike_pi_sc"=as.data.frame(bike_pi_sc_clcnt_fulltri),
  "bike_pi_lc"=as.data.frame(bike_pi_lc_clcnt_fulltri),
  "bike_le"=as.data.frame(bike_le_clcnt_fulltri),
  "health_total"=as.data.frame(pa_clcnt_fulltri),
  "pet_total"=as.data.frame(pet_clcnt_fulltri)
)
write_xlsx(l, "claim_count_triangles.xlsx")
#
# ------Export LR traiangles to Excel
#
l <- list(
  "auto_bil_sc"=as.data.frame(auto_bil_sc_lr_tri),
  "auto_bil_lc"=as.data.frame(auto_bil_lc_lr_tri),
  "auto_pdl"=as.data.frame(auto_pdl_lr_tri),
  "auto_ppa"=as.data.frame(auto_ppa_lr_tri),
  "auto_odpe"=as.data.frame(auto_odpe_lr_tri),
  "auto_pi_sc"=as.data.frame(auto_pi_sc_lr_tri),
  "auto_pi_lc"=as.data.frame(auto_pi_lc_lr_tri),
  "auto_axap"=as.data.frame(auto_axap_lr_tri),
  "auto_fb"=as.data.frame(auto_fb_lr_tri),
  "auto_le"=as.data.frame(auto_le_lr_tri),
  "auto_eq"=as.data.frame(auto_eq_lr_tri),
  "bike_bil_sc"=as.data.frame(bike_bil_sc_lr_tri),
  "bike_bil_lc"=as.data.frame(bike_bil_lc_lr_tri),
  "bike_pdl"=as.data.frame(bike_pdl_lr_tri),
  "bike_ppa"=as.data.frame(bike_ppa_lr_tri),
  "bike_pi_sc"=as.data.frame(bike_pi_sc_lr_tri),
  "bike_pi_lc"=as.data.frame(bike_pi_lc_lr_tri),
  "bike_le"=as.data.frame(bike_le_lr_tri),
  "health_total"=as.data.frame(pa_lr_tri),
  "pet_total"=as.data.frame(pet_lr_tri)
)
write_xlsx(l, "LR_triangles.xlsx")
#
# ------Export Frequency traiangles to Excel
#
l <- list(
  "auto_bil_sc"=as.data.frame(auto_bil_sc_freq_tri),
  "auto_bil_lc"=as.data.frame(auto_bil_lc_freq_tri),
  "auto_pdl"=as.data.frame(auto_pdl_freq_tri),
  "auto_ppa"=as.data.frame(auto_ppa_freq_tri),
  "auto_odpe"=as.data.frame(auto_odpe_freq_tri),
  "auto_pi_sc"=as.data.frame(auto_pi_sc_freq_tri),
  "auto_pi_lc"=as.data.frame(auto_pi_lc_freq_tri),
  "auto_axap"=as.data.frame(auto_axap_freq_tri),
  "auto_fb"=as.data.frame(auto_fb_freq_tri),
  "auto_le"=as.data.frame(auto_le_freq_tri),
  "auto_eq"=as.data.frame(auto_eq_freq_tri),
  "bike_bil_sc"=as.data.frame(bike_bil_sc_freq_tri),
  "bike_bil_lc"=as.data.frame(bike_bil_lc_freq_tri),
  "bike_pdl"=as.data.frame(bike_pdl_freq_tri),
  "bike_ppa"=as.data.frame(bike_ppa_freq_tri),
  "bike_pi_sc"=as.data.frame(bike_pi_sc_freq_tri),
  "bike_pi_lc"=as.data.frame(bike_pi_lc_freq_tri),
  "bike_le"=as.data.frame(bike_le_freq_tri),
  "health_total"=as.data.frame(pa_freq_tri),
  "pet_total"=as.data.frame(pet_freq_tri)
)
write_xlsx(l, "freq_triangles.xlsx")
#
# ------Export Severity traiangles to Excel
#
l <- list(
  "auto_bil_sc"=as.data.frame(auto_bil_sc_ac_tri),
  "auto_bil_lc"=as.data.frame(auto_bil_lc_ac_tri),
  "auto_pdl"=as.data.frame(auto_pdl_ac_tri),
  "auto_ppa"=as.data.frame(auto_ppa_ac_tri),
  "auto_odpe"=as.data.frame(auto_odpe_ac_tri),
  "auto_pi_sc"=as.data.frame(auto_pi_sc_ac_tri),
  "auto_pi_lc"=as.data.frame(auto_pi_lc_ac_tri),
  "auto_axap"=as.data.frame(auto_axap_ac_tri),
  "auto_fb"=as.data.frame(auto_fb_ac_tri),
  "auto_le"=as.data.frame(auto_le_ac_tri),
  "auto_eq"=as.data.frame(auto_eq_ac_tri),
  "bike_bil_sc"=as.data.frame(bike_bil_sc_ac_tri),
  "bike_bil_lc"=as.data.frame(bike_bil_lc_ac_tri),
  "bike_pdl"=as.data.frame(bike_pdl_ac_tri),
  "bike_ppa"=as.data.frame(bike_ppa_ac_tri),
  "bike_pi_sc"=as.data.frame(bike_pi_sc_ac_tri),
  "bike_pi_lc"=as.data.frame(bike_pi_lc_ac_tri),
  "bike_le"=as.data.frame(bike_le_ac_tri),
  "health_total"=as.data.frame(pa_ac_tri),
  "pet_total"=as.data.frame(pet_ac_tri)
)
write_xlsx(l, "severity_triangles.xlsx")
#
# ------Export BC traiangles to Excel
#
l <- list(
  "auto_bil_sc"=as.data.frame(auto_bil_sc_bc_tri),
  "auto_bil_lc"=as.data.frame(auto_bil_lc_bc_tri),
  "auto_pdl"=as.data.frame(auto_pdl_bc_tri),
  "auto_ppa"=as.data.frame(auto_ppa_bc_tri),
  "auto_odpe"=as.data.frame(auto_odpe_bc_tri),
  "auto_pi_sc"=as.data.frame(auto_pi_sc_bc_tri),
  "auto_pi_lc"=as.data.frame(auto_pi_lc_bc_tri),
  "auto_axap"=as.data.frame(auto_axap_bc_tri),
  "auto_fb"=as.data.frame(auto_fb_bc_tri),
  "auto_le"=as.data.frame(auto_le_bc_tri),
  "auto_eq"=as.data.frame(auto_eq_bc_tri),
  "bike_bil_sc"=as.data.frame(bike_bil_sc_bc_tri),
  "bike_bil_lc"=as.data.frame(bike_bil_lc_bc_tri),
  "bike_pdl"=as.data.frame(bike_pdl_bc_tri),
  "bike_ppa"=as.data.frame(bike_ppa_bc_tri),
  "bike_pi_sc"=as.data.frame(bike_pi_sc_bc_tri),
  "bike_pi_lc"=as.data.frame(bike_pi_lc_bc_tri),
  "bike_le"=as.data.frame(bike_le_bc_tri),
  "health_total"=as.data.frame(pa_bc_tri),
  "pet_total"=as.data.frame(pet_bc_tri)
)
write_xlsx(l, "BurningCost_triangles.xlsx")
#
# Actual vs Expected
#
devmth <- 7 # 202306 ~ 202309 4 months
fulltrifile <- "Previous_full_triangles.xlsx" # full triangle from previous 2nd opinion
calc_actual_vs_expected <- function(sheetname, inc_tri) {
  expected <- read_excel(fulltrifile, sheet = sheetname)
  ay <- expected[, 1]
  expected <- expected[, -c(1, dim(expected)[2])]
  actual <- inc_tri[-c((dim(inc_tri)[1]-devmth+1):dim(inc_tri)[1]), -c(1, (dim(inc_tri)[2]-devmth+1):dim(inc_tri)[2])] #HY
  act_vs_exp <- actual - expected
  act_vs_exp <- cbind(ay, act_vs_exp)
  return(act_vs_exp)
}
# Auto
ave_auto_bil_sc <- calc_actual_vs_expected(sheetname = "auto_bil_sc", auto_bil_sc_tri)
ave_auto_bil_lc <- calc_actual_vs_expected(sheetname = "auto_bil_lc", auto_bil_lc_tri)
ave_auto_pdl <- calc_actual_vs_expected(sheetname = "auto_pdl", auto_pdl_tri)
ave_auto_ppa <- calc_actual_vs_expected(sheetname = "auto_ppa", auto_ppa_tri)
ave_auto_pi_sc <- calc_actual_vs_expected(sheetname = "auto_pi_sc", auto_pi_sc_tri)
ave_auto_pi_lc <- calc_actual_vs_expected(sheetname = "auto_pi_lc", auto_pi_lc_tri)
ave_auto_odpe <- calc_actual_vs_expected(sheetname = "auto_odpe", auto_odpe_tri)
ave_auto_le <- calc_actual_vs_expected(sheetname = "auto_le", auto_le_tri)
ave_auto_fb <- calc_actual_vs_expected(sheetname = "auto_fb", auto_fb_tri)
ave_auto_axap <- calc_actual_vs_expected(sheetname = "auto_axap", auto_axap_tri)
# Bike
ave_bike_bil_sc <- calc_actual_vs_expected(sheetname = "bike_bil_sc", bike_bil_sc_tri)
ave_bike_bil_lc <- calc_actual_vs_expected(sheetname = "bike_bil_lc", bike_bil_lc_tri)
ave_bike_pdl <- calc_actual_vs_expected(sheetname = "bike_pdl", bike_pdl_tri)
ave_bike_ppa <- calc_actual_vs_expected(sheetname = "bike_ppa", bike_ppa_tri)
ave_bike_pi_sc <- calc_actual_vs_expected(sheetname = "bike_pi_sc", bike_pi_sc_tri)
ave_bike_pi_lc <- calc_actual_vs_expected(sheetname = "bike_pi_lc", bike_pi_lc_tri)
ave_bike_le <- calc_actual_vs_expected(sheetname = "bike_le", bike_le_tri) # Bike LE under construction
# Health
ave_health <- calc_actual_vs_expected(sheetname = "health_total", pa_tri)
# Pet
ave_pet <- calc_actual_vs_expected(sheetname = "pet_total", pet_tri)
#
# Export to Excel
l <- list(
  "auto_bil_sc" = ave_auto_bil_sc,
  "auto_bil_lc" = ave_auto_bil_lc,
  "auto_pdl" = ave_auto_pdl,
  "auto_ppa" = ave_auto_ppa,
  "auto_pi_sc" = ave_auto_pi_sc,
  "auto_pi_lc" = ave_auto_pi_lc,
  "auto_odpe" = ave_auto_odpe,
  "auto_le" = ave_auto_le,
  "auto_fb" = ave_auto_fb,
  "auto_axap" = ave_auto_axap,
  "bike_bil_sc" = ave_bike_bil_sc,
  "bike_bil_lc" = ave_bike_bil_lc,
  "bike_pdl" = ave_bike_pdl,
  "bike_ppa" = ave_bike_ppa,
  "bike_pi_sc" = ave_bike_pi_sc,
  "bike_pi_lc" = ave_bike_pi_lc,
  "bike_le" = ave_bike_le,
  "health" = ave_health,
  "pet" = ave_pet
)
write_xlsx(l, "actual_vs_expected.xlsx")

sum(getLatestCumulative(as.matrix(ave_auto_bil_sc)))

# EP grwoth rate　年間ベースで比較
cy_ym_beg <- "202305" # 直近1年前
py_ym_beg <- "202210" # 前期のFlash時点から1年前
py_ym_end <- "202309"

# 直近1年のEP
motor_ep_current <- df_motor_exp %>%
  filter(ym >= cy_ym_beg) %>%
  summarise(EP = sum(EARNINGS))
# 前期のFlash時点から1年前
motor_ep_prev <- df_motor_exp %>%
  filter(ym >= py_ym_beg & ym <= py_ym_end) %>%
  summarise(EP = sum(EARNINGS))

motor_ep_growth_rate <- motor_ep_current$EP / motor_ep_prev$EP

# Health
pa_ep_current <- df_pa_exp %>%
  filter(ym >= cy_ym_beg) %>%
  summarise(EP = sum(EARNINGS))

pa_ep_prev <- df_pa_exp %>%
  filter(ym >= py_ym_beg & ym <= py_ym_end) %>%
  summarise(EP = sum(EARNINGS))
# Pet
pet_ep_current <- df_exp_ep2 %>%
  filter(BIZ_LINE == "PET") %>%
  filter(ym >= cy_ym_beg) %>%
  summarise(EP = sum(EARNINGS))

pet_ep_prev <- df_exp_ep2 %>%
  filter(BIZ_LINE == "PET") %>%
  filter(ym >= py_ym_beg & ym <= py_ym_end) %>%
  summarise(EP = sum(EARNINGS))

ep_grwoth_rate <- ((motor_ep_current$EP + pa_ep_current$EP + pet_ep_current) / (motor_ep_prev$EP + pa_ep_prev$EP + pet_ep_prev)) -1
write_xlsx(ep_grwoth_rate, "ep_growth_rate.xlsx")


# ---------------------

# 
#
# Projection
#
#

## Functions for Future AY ##

# Paid future
get_paid_future_select_dev_mth <- function(paid_fulltri, dev_mth) {
  paid_fulltri <- replace(paid_fulltri, is.na(paid_fulltri), 0)
  paid_cy <- sum(paid_fulltri[c((dim(paid_fulltri)[1]-2):dim(paid_fulltri)[1]), c(2)])
  paid_py <- sum(paid_fulltri[c((dim(paid_fulltri)[1]-14):(dim(paid_fulltri)[1]-12)), c(2)])
  paid_growth_rate <- paid_cy / paid_py
  paid_growth_rate <- replace(paid_growth_rate, is.na(paid_growth_rate), 1)
  if (dev_mth == 3) {
  paid_future <- paid_fulltri[c((dim(paid_fulltri)[1]-11):(dim(paid_fulltri)[1]-9)), c(2:4)] * paid_growth_rate
  paid_projected <- c(paid_future[1, 3], paid_future[2, 2], paid_future[3, 1])
  }
  else if (dev_mth == 2) {
    paid_future <- paid_fulltri[c((dim(paid_fulltri)[1]-11):(dim(paid_fulltri)[1]-10)), c(2:3)] * paid_growth_rate
    paid_projected <- c(paid_future[1, 2], paid_future[2, 1])
  }
  else {
  paid_future <- paid_fulltri[c((dim(paid_fulltri)[1]-12)), c(2)] * paid_growth_rate
  paid_projected <- paid_future
  }
  return(paid_projected)
}
# Eaxa future
get_eaxa_future_select_dev_mth <- function(fulltri, dev_mth) {
  fulltri <- replace(fulltri, is.na(fulltri), 0)
  eaxa_by_mth <- data.frame(fulltri)[, names(data.frame(fulltri)) %in% c('V1', 'Ult')]
  eaxa_cy <- sum(eaxa_by_mth[c((dim(eaxa_by_mth)[1]-2):dim(eaxa_by_mth)[1]), c(2)])
  eaxa_py <- sum(eaxa_by_mth[c((dim(eaxa_by_mth)[1]-14):(dim(eaxa_by_mth)[1]-12)), c(2)])
  eaxa_growth_rate <- eaxa_cy / eaxa_py
  eaxa_growth_rate <- replace(eaxa_growth_rate, is.na(eaxa_growth_rate), 1)
  if (dev_mth == 3) {
  eaxa_future <- eaxa_by_mth[c((dim(eaxa_by_mth)[1]-11):(dim(eaxa_by_mth)[1]-9)), 2] * eaxa_growth_rate
  }
  else if (dev_mth == 2) {
    eaxa_future <- eaxa_by_mth[c((dim(eaxa_by_mth)[1]-10):(dim(eaxa_by_mth)[1]-9)), 2] * eaxa_growth_rate
  }
  else {
  eaxa_future <- eaxa_by_mth[c((dim(eaxa_by_mth)[1]-11)), 2] * eaxa_growth_rate
  }
  return(eaxa_future)
}

## Runoff ##

# Runoff paid
get_runoff_paid_developed_select_dev_mth <- function(paid_fulltri, dev_mth) {
  if (dev_mth == 3) {
  paid2 <- paid_fulltri[-c(1:3), !names(paid_fulltri) %in% c('V1', 'X1', 'X2', 'X3', 'Ult')]
  paid_dev <- c()
  for(i in 1:dim(paid2)[[1]]) {
    bal <- paid2[dim(paid2)[[1]]+1-i, i]
    paid_dev <- rbind(paid_dev, bal)
  }
  paid_dev_runoff <- c(paid_fulltri[c(1:3), 'Ult'], rev(paid_dev))
  }
  else if (dev_mth == 2) {
    paid2 <- paid_fulltri[-c(1:2), !names(paid_fulltri) %in% c('V1', 'X1', 'X2', 'Ult')]
    paid_dev <- c()
    for(i in 1:dim(paid2)[[1]]) {
      bal <- paid2[dim(paid2)[[1]]+1-i, i]
      paid_dev <- rbind(paid_dev, bal)
    }
    paid_dev_runoff <- c(paid_fulltri[c(1:2), 'Ult'], rev(paid_dev))
  }
  else {
    paid2 <- paid_fulltri[-c(1), !names(paid_fulltri) %in% c('V1', 'X1', 'Ult')]
    paid_dev <- c()
    for(i in 1:dim(paid2)[[1]]) {
      bal <- paid2[dim(paid2)[[1]]+1-i, i]
      paid_dev <- rbind(paid_dev, bal)
    }
    paid_dev_runoff <- c(paid_fulltri[c(1), 'Ult'], rev(paid_dev))  
  }
  paid_dev_runoff <- data.frame(cbind(paid_fulltri[, 1], paid_dev_runoff))
  colnames(paid_dev_runoff) <- c("ym", 'PAID')
  
  
  return (paid_dev_runoff)
}
# Runoff paid group by AY
get_paid_developed_runoff_by_ay <- function(paid_dev_runoff) {
  
  paid_by_ay <- paid_dev_runoff %>%
    mutate(AY=substring(ym, 1, 4)) %>%
    group_by(AY) %>%
    summarise(PAID=sum(PAID))
  
  return(paid_by_ay)
}

# Runoff + Future
get_eaxa_res_for_projection <- function(eaxa_future, paid_projected, eaxares_pcfy_runoff) {
  eaxa_res_future <- eaxa_future - paid_projected
  eaxa_res_pcfy <- sum(eaxares_pcfy_runoff) + sum(eaxa_res_future)
  
  return(eaxa_res_pcfy)
}

# EAXA Future
get_eaxa_future_select_dev_mth_v1 <- function(eaxa_by_mth, eaxa_growth_rate, dev_mth) {
  if (dev_mth == 3) {
    eaxa_growth_rate <- replace(eaxa_growth_rate, is.na(eaxa_growth_rate), 1)
    eaxa_future <- eaxa_by_mth[c((dim(eaxa_by_mth)[1]-11):(dim(eaxa_by_mth)[1]-9)), 2] * eaxa_growth_rate
  }
  else if (dev_mth == 2) {
    eaxa_growth_rate <- replace(eaxa_growth_rate, is.na(eaxa_growth_rate), 1)
    eaxa_future <- eaxa_by_mth[c((dim(eaxa_by_mth)[1]-11):(dim(eaxa_by_mth)[1]-10)), 2] * eaxa_growth_rate
  }
  else {
    eaxa_growth_rate <- replace(eaxa_growth_rate, is.na(eaxa_growth_rate), 1)
    eaxa_future <- eaxa_by_mth[c((dim(eaxa_by_mth)[1]-12)), 2] * eaxa_growth_rate
  }
  return(eaxa_future)
}

# Motor --------------------

# Motor total ---------------------
product <- "Motor"
cov <- "Total"
cl_type <- "Total"
lkr_type <- "lkr_paid_smpl"
develop_month <- 2
# make paid full triangle
motor_paid_fulltri <- data.frame(get_full_triangle_paid(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type))
#write.csv(motor_paid_fulltri, "motor_paid_fulltri.csv")
# Get paid developed runoff
motor_paid_runoff_developed <- get_runoff_paid_developed_select_dev_mth(motor_paid_fulltri, dev_mth = develop_month)
# Paid summarize by AY
tbl <- get_paid_developed_runoff_by_ay(motor_paid_runoff_developed)
# EAXA Reserve Runoff
motor_eaxares_runoff <- motor_exhibit[-c((dim(motor_exhibit)[1]-2):dim(motor_exhibit)[1]), "EAXA"] - tbl[, "PAID"][[1]] # calc check OK
# Paid future
motor_paid_projected <- get_paid_future_select_dev_mth(motor_paid_fulltri, dev_mth = develop_month)
# Motor EAXA by Month for projection
eaxa_lst <- list(auto_bil_lc_fulltri, auto_bil_sc_fulltri, auto_pi_lc_fulltri, auto_pi_sc_fulltri, auto_pdl_fulltri, auto_ppa_fulltri, auto_odpe_fulltri, auto_odpe_ev_fulltri,
                 auto_fb_fulltri, auto_le_fulltri, auto_axap_fulltri, auto_eq_fulltri,
                 bike_bil_lc_fulltri, bike_bil_sc_fulltri, bike_pi_lc_fulltri, bike_pi_sc_fulltri, bike_pdl_fulltri, bike_ppa_fulltri, bike_le_fulltri,
                 pa_fulltri, pet_fulltri)
eaxa_lnm <- list("Auto_BIL_LC", "Auto_BIL_SC", "Auto_PI_LC", "Auto_PI_SC", "Auto_PDL", "Auto_PPA", "Auto_ODPE", "Auto_ODPE_CAT",
                 "Auto_FB", "Auto_LE", "Auto_Axa+", "Auto_EQ",
                 "Bike_BIL_LC", "Bike_BIL_SC",  "Bike_PI_LC", "Bike_PI_SC", "Bike_PDL", "Bike_PPA", "Bike_LE",
                 "PA", "Pet")
eaxa_all_cov <- c()
for(i in seq_along(eaxa_lst)) {
  df <- data.frame(eaxa_lst[[i]])[, names(data.frame(eaxa_lst[[i]])) %in% c('V1', 'Ult')] %>%
    dplyr::mutate(cov=eaxa_lnm[[i]])
  eaxa_all_cov <- rbind(eaxa_all_cov, df)
}
eaxa_all_cov[is.na(eaxa_all_cov)] <- 0
motor_eaxa <- eaxa_all_cov[eaxa_all_cov["cov"] != "PA" & eaxa_all_cov["cov"] != "Pet", ]
# Motor coverage別のEAXAを合算
motor_eaxa_by_mth <- motor_eaxa %>%
  group_by(V1) %>%
  summarise(Ult = sum(Ult))
write.csv(motor_eaxa_by_mth, "motor_eaxa_by_mth.csv")
# growth rateの準備
# CYの直近3か月 (PCHY24 : Apr 2024のHailの影響を外すため4月を除いてPFを作成)
motor_eaxa_cy <- sum(motor_eaxa_by_mth[c((dim(motor_eaxa_by_mth)[1]-2):dim(motor_eaxa_by_mth)[1]-1), c(2)])
# PYの直近３か月
motor_eaxa_py <- sum(motor_eaxa_by_mth[c((dim(motor_eaxa_by_mth)[1]-15):(dim(motor_eaxa_by_mth)[1]-13)), c(2)])
# Growth Rate
motor_eaxa_growth_rate <- max(motor_eaxa_cy / motor_eaxa_py, 0)
# EAXAにGrowth Rateをかけて、CYの予測をする
motor_eaxa_future <- get_eaxa_future_select_dev_mth_v1(motor_eaxa_by_mth, 
                                                       motor_eaxa_growth_rate, 
                                                       dev_mth = develop_month)
# 予測したEAXAから予測したPaidを引いてリザーブを計算
motor_eaxa_res_future <- motor_eaxa_future - motor_paid_projected
# Motor EAXA Reserveの総額
motor_eaxa_res <- sum(motor_eaxares_runoff) + sum(motor_eaxa_res_future)

motor_eaxa_res_by_ay_runoff_future <- data.frame(AY = motor_exhibit[-c((dim(motor_exhibit)[1]-2):dim(motor_exhibit)[1]),'AY'], eres = motor_eaxares_runoff)
motor_eaxa_res_by_ay_runoff_future[motor_eaxa_res_by_ay_runoff_future$AY == 2024, 'eres'] <- 
  motor_eaxa_res_by_ay_runoff_future[motor_eaxa_res_by_ay_runoff_future$AY == 2024, 'eres'] + sum(motor_eaxa_res_future)
write_xlsx(motor_eaxa_res_by_ay_runoff_future, "motor_eaxa_res_by_ay_runoff_future.xlsx")

motor_cy_paid <- tbl[, "PAID"][[1]][dim(tbl[, "PAID"])[1]] + motor_paid_projected


# PA ---------------------
product <- "PA"
cov <- "Total"
cl_type <- "Attrit"
lkr_type <- "lkr12_paid"
pa_paid_fulltri <- data.frame(get_full_triangle_paid(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type))
# Get paid developed
pa_paid_runoff_developed <- get_runoff_paid_developed_select_dev_mth(pa_paid_fulltri, dev_mth = develop_month)
# Paid summarize by AY
pa_tbl <- get_paid_developed_runoff_by_ay(pa_paid_runoff_developed)
# EAXA Reserve Runoff
pa_eaxares_runoff <- pa_exhibit[-c(1:2, (dim(pa_exhibit)[1]-2):dim(pa_exhibit)[1]), "EAXA"] - pa_tbl[, "PAID"][[1]] # calc check OK
# Paid future
pa_paid_projected <- get_paid_future_select_dev_mth(pa_paid_fulltri, dev_mth = develop_month)
# Eaxa future
pa_eaxa_future <- get_eaxa_future_select_dev_mth(pa_fulltri, dev_mth = develop_month)
# EAXA reserve runoff + future
pa_eaxa_res <- get_eaxa_res_for_projection(pa_eaxa_future, pa_paid_projected, pa_eaxares_runoff)

pa_eaxa_res_by_ay_runoff_future <- data.frame(AY = pa_exhibit[-c(1:2,(dim(pa_exhibit)[1]-2):dim(pa_exhibit)[1]),'AY'], eres = pa_eaxares_runoff)
pa_eaxa_res_by_ay_runoff_future[pa_eaxa_res_by_ay_runoff_future$AY == 2024, 2] <- 
  pa_eaxa_res_by_ay_runoff_future[pa_eaxa_res_by_ay_runoff_future$AY == 2024, 2] + sum(pa_eaxa_future - pa_paid_projected)
write_xlsx(pa_eaxa_res_by_ay_runoff_future, "pa_eaxa_res_by_ay_runoff_future.xlsx")

# Pet ---------------------
product <- "PET"
cov <- "Total"
cl_type <- "Attrit"
lkr_type <- "lkr12_paid"
pet_paid_fulltri <- data.frame(get_full_triangle_paid(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type))
# Get paid developed
pet_paid_runoff_developed <- get_runoff_paid_developed_select_dev_mth(pet_paid_fulltri, dev_mth = develop_month)
# Paid summarize by AY
pet_tbl <- get_paid_developed_runoff_by_ay(pet_paid_runoff_developed)
# PCFY Dec EAXA Reserve (EAXA at Sep - Paid at Dec) Runoff
pet_eaxares_runoff <- pet_exhibit[-c(1:7, (dim(pet_exhibit)[1]-2):dim(pet_exhibit)[1]), "EAXA"] - pet_tbl[, "PAID"][[1]] # calc check OK
# Paid future
pet_paid_projected <- get_paid_future_select_dev_mth(pet_paid_fulltri, dev_mth = develop_month)
# Eaxa future
pet_eaxa_future <- get_eaxa_future_select_dev_mth(pet_fulltri, dev_mth = develop_month)
# EAXA reserve runoff + future
pet_eaxa_res <- get_eaxa_res_for_projection(pet_eaxa_future, pet_paid_projected, pet_eaxares_runoff[-c(1:4)]) #Petは2009年以前を除外

pet_eaxa_res_by_ay_runoff_future <- data.frame(AY = pet_exhibit[-c(1:11,(dim(pet_exhibit)[1]-2):dim(pet_exhibit)[1]),'AY'], eres = pet_eaxares_runoff[-c(1:4)])
pet_eaxa_res_by_ay_runoff_future[pet_eaxa_res_by_ay_runoff_future$AY == 2024, 2] <- 
  pet_eaxa_res_by_ay_runoff_future[pet_eaxa_res_by_ay_runoff_future$AY == 2024, 2] + sum(pet_eaxa_future - pet_paid_projected)
write_xlsx(pet_eaxa_res_by_ay_runoff_future, "pet_eaxa_res_by_ay_runoff_future.xlsx")

c(motor_eaxa_res, pa_eaxa_res, pet_eaxa_res, sum(motor_eaxa_res, pa_eaxa_res, pet_eaxa_res))


# Export to Excel
setwd(workfile_wd)
l <- list(
  "motor" = motor_paid_fulltri,
  "health" = pa_paid_fulltri,
  "pet" = pet_paid_fulltri
)
write_xlsx(l, "paid full triangles.xlsx")



# ----- (not completed yet) mod PCFY21 --------------------
develop_month <- 1

# Calculate EAXA Reserve by accident months and make a table
get_eres_by_cov <- function(fulltri_paid, fulltri, LoB, products, cov_name, exhibit, cl_type, mega_LoB) {
  #eres runoff
  #runoff_by_mth <- get_eaxa_res_runoff(fulltri_paid, develop_month, fulltri, LoB)
  runoff_paid_developed <- get_runoff_paid_developed_select_dev_mth(data.frame(fulltri_paid), develop_month)[, 'PAID']
  runoff_paid_developed <- replace(runoff_paid_developed, is.na(runoff_paid_developed), 0)
  runoff_eaxa <- replace(fulltri[, 'Ult'], is.na(fulltri[, 'Ult']), 0)
  runoff_by_mth <- runoff_eaxa - runoff_paid_developed
  runoff_by_mth <- data.frame(runoff_by_mth) %>% 
    mutate(cov_name = cov_name, ym = fulltri[, 1], eaxa = runoff_eaxa, paid = runoff_paid_developed) %>% 
    rename(eres = runoff_by_mth)
  
  eres_runoff <- runoff_by_mth[, c("ym", "eres")]
  
  #eres future
  #ep projection
  df_ep <- df_exp_ep2 %>% 
    filter(PRODUCTS == products & COV_NAME == cov_name) %>% 
    mutate(rownum=row_number())
  #df_rownum <- data.frame(rownum = (dim(df_ep)[1]+1):(dim(df_ep)[1]+2))
  if (develop_month == 1) {
    df_rownum <- data.frame(rownum = (dim(df_ep)[1]+1))  
  } else {
    df_rownum <- data.frame(rownum = (dim(df_ep)[1]+1):(dim(df_ep)[1]+2))
  }
  df_ep <-df_ep %>% 
    filter(rownum >= (dim(df_ep)[1]-11))
  #EP linear model
  lm_ep <- lm(EARNINGS ~ rownum, data = df_ep)
  #EP linear projection
  ep_future <- predict(lm_ep, df_rownum)
  #get CY EAXA LR
  eaxa_lr <- exhibit %>% 
    filter(AY == "CY_Total") %>% 
    dplyr::select(EAXA_LR)
  #calc EAXA Future
  eaxa_future <- ep_future * as.numeric(eaxa_lr)
  #calc Paid future
  #paid_future <- get_paid_future_select_dev_mth(fulltri_paid, dev_mth = develop_month)
  paid_tri <- Reserve_Data$paid_tri[Reserve_Data$PRODUCTS==products & Reserve_Data$COV_NAME==cov_name & Reserve_Data$CL_Large2==cl_type][[1]]
  df_paid1 <- data.frame(getLatestCumulative(paid_tri[,-1]))
  df_ratio <- df_paid1[[1]] / runoff_by_mth['eaxa']
  df_ratio <- replace(df_ratio, is.na(df_ratio), 0)
  if (develop_month == 1) {
    paid_future <- eaxa_future * df_ratio[dim(df_ratio)[1], ]
  } else {
    paid_future <- eaxa_future * df_ratio[(dim(df_ratio)[1]-1):dim(df_ratio)[1], ]
  }
  #calc EAXA reserve future
  #eres_future <- data.frame(ym =(eres_runoff[dim(eres_runoff)[1], "ym"] + 1):(eres_runoff[dim(eres_runoff)[1], "ym"] + 2), eres = eaxa_future - paid_future)
  if (develop_month == 1) {
    eres_future <- data.frame(ym =(eres_runoff[dim(eres_runoff)[1], "ym"] + 1), eres = eaxa_future - paid_future)
  } else {
    eres_future <- data.frame(ym =(eres_runoff[dim(eres_runoff)[1], "ym"] + 1):(eres_runoff[dim(eres_runoff)[1], "ym"] + 2), eres = eaxa_future - paid_future)
  }
  #bind future and runoff
  eres_runoff_future <- rbind(eres_runoff, eres_future)
  #calc AY
  eres_runoff_future$ay <- substring(as.character(eres_runoff_future[, "ym"]), 1, 4)
  #summarize by AY
  eres_runoff_future_by_ay <- eres_runoff_future %>% 
    group_by(ay) %>% 
    summarise(eres = sum(eres))
  #py cy summary
  eres_py <- eres_runoff_future_by_ay %>% 
    filter(ay < as.character(eres_runoff_future_by_ay[dim(eres_runoff_future_by_ay)[1], "ay"])) %>% 
    summarise(eres = sum(eres))
  eres_cy <- eres_runoff_future_by_ay %>% 
    filter(ay == as.character(eres_runoff_future_by_ay[dim(eres_runoff_future_by_ay)[1], "ay"])) %>% 
    summarise(eres = sum(eres))
  eres_all_yr <- eres_runoff_future_by_ay %>% 
    summarise(eres = sum(eres))
  #
  eres_py <- data.frame(ay = "PY_Total", eres = eres_py)
  eres_cy <- data.frame(ay = "CY_Total", eres = eres_cy)
  eres_all_yr <- data.frame(ay = "Total", eres = eres_all_yr)
  eres_runoff_future_by_ay <- rbind(eres_runoff_future_by_ay, eres_py, eres_cy, eres_all_yr)
  eres_runoff_future_by_ay <- eres_runoff_future_by_ay %>% 
    mutate(product = products, LoB = LoB, cov = cov_name, cl_type = cl_type, mega_LoB = mega_LoB)
  
  return(eres_runoff_future_by_ay)
  #return(eaxa_future)
  
}

eres_runoff_future_by_ay_auto_bil_lc <- get_eres_by_cov(auto_bil_lc_fulltri_paid, auto_bil_lc_fulltri, "Auto_BIL_LC", "Auto", "BIL", auto_bil_lc_exhibit, "Large", "Motor")
eres_runoff_future_by_ay_auto_bil_sc <- get_eres_by_cov(auto_bil_sc_fulltri_paid, auto_bil_sc_fulltri, "Auto_BIL_SC", "Auto", "BIL", auto_bil_sc_exhibit, "Attrit", "Motor")
eres_runoff_future_by_ay_auto_pi_lc <- get_eres_by_cov(auto_pi_lc_fulltri_paid, auto_pi_lc_fulltri, "Auto_PI_LC", "Auto", "PI", auto_pi_lc_exhibit, "Large", "Motor")
eres_runoff_future_by_ay_auto_pi_sc <- get_eres_by_cov(auto_pi_sc_fulltri_paid, auto_pi_sc_fulltri, "Auto_PI_SC", "Auto", "PI", auto_pi_sc_exhibit, "Attrit", "Motor")
eres_runoff_future_by_ay_auto_pdl <- get_eres_by_cov(auto_pdl_fulltri_paid, auto_pdl_fulltri, "Auto_PDL", "Auto", "PDL", auto_pdl_exhibit, "Total", "Motor")
eres_runoff_future_by_ay_auto_ppa <- get_eres_by_cov(auto_ppa_fulltri_paid, auto_ppa_fulltri, "Auto_PPA", "Auto", "PPA", auto_ppa_exhibit, "Total", "Motor")
eres_runoff_future_by_ay_auto_odpe <- get_eres_by_cov(auto_odpe_fulltri_paid, auto_odpe_fulltri, "Auto_ODPE", "Auto", "ODPE", auto_odpe_exhibit, "Attrit", "Motor")
eres_runoff_future_by_ay_auto_odpe_ev <- get_eres_by_cov(auto_odpe_ev_fulltri_paid, auto_odpe_ev_fulltri, "Auto_ODPE_EV", "Auto", "ODPE", auto_odpe_ev_exhibit, "Evented", "Motor")
eres_runoff_future_by_ay_auto_fb <- get_eres_by_cov(auto_fb_fulltri_paid, auto_fb_fulltri, "Auto_FB", "Auto", "FamilyBIKE", auto_fb_exhibit, "Total", "Motor")
eres_runoff_future_by_ay_auto_le <- get_eres_by_cov(auto_le_fulltri_paid, auto_le_fulltri, "Auto_LE", "Auto", "LawyerE", auto_le_exhibit, "Total", "Motor")
eres_runoff_future_by_ay_auto_axap <- get_eres_by_cov(auto_axap_fulltri_paid, auto_axap_fulltri, "Auto_AXAP", "Auto", "AXA+", auto_axap_exhibit, "Total", "Motor")
eres_runoff_future_by_ay_auto_eq <- get_eres_by_cov(auto_eq_fulltri_paid, auto_eq_fulltri, "Auto_EQ", "Auto", "OD_EQ", auto_eq_exhibit, "Total", "Motor")
#bike
eres_runoff_future_by_ay_bike_bil_lc <- get_eres_by_cov(bike_bil_lc_fulltri_paid, bike_bil_lc_fulltri, "Bike_BIL_LC", "BIKE", "BIL", bike_bil_lc_exhibit, "Large", "Motor")
eres_runoff_future_by_ay_bike_bil_sc <- get_eres_by_cov(bike_bil_sc_fulltri_paid, bike_bil_sc_fulltri, "Bike_BIL_SC", "BIKE", "BIL", bike_bil_sc_exhibit, "Attrit", "Motor")
eres_runoff_future_by_ay_bike_pi_lc <- get_eres_by_cov(bike_pi_lc_fulltri_paid, bike_pi_lc_fulltri, "Bike_PI_LC", "BIKE", "PI", bike_pi_lc_exhibit, "Large", "Motor")
eres_runoff_future_by_ay_bike_pi_sc <- get_eres_by_cov(bike_pi_sc_fulltri_paid, bike_pi_sc_fulltri, "Bike_PI_SC", "BIKE", "PI", bike_pi_sc_exhibit, "Attrit", "Motor")
eres_runoff_future_by_ay_bike_pdl <- get_eres_by_cov(bike_pdl_fulltri_paid, bike_pdl_fulltri, "Bike_PDL", "BIKE", "PDL", bike_pdl_exhibit, "Total", "Motor")
eres_runoff_future_by_ay_bike_ppa <- get_eres_by_cov(bike_ppa_fulltri_paid, bike_ppa_fulltri, "Bike_PPA", "BIKE", "PPA", bike_ppa_exhibit, "Total", "Motor")
eres_runoff_future_by_ay_bike_le <- get_eres_by_cov(bike_le_fulltri_paid, bike_le_fulltri, "Bike_LE", "BIKE", "LawyerE", bike_le_exhibit, "Total", "Motor")
#pa
eres_runoff_future_by_ay_pa <- get_eres_by_cov(pa_fulltri_paid, pa_fulltri, "PA", "PA", "Total", pa_exhibit, "Attrit", "PA")
#pet
eres_runoff_future_by_ay_pet <- get_eres_by_cov(pet_fulltri_paid, pet_fulltri, "Pet", "PET", "Total", pet_exhibit, "Attrit", "Pet")

df_eres <- rbind(
  eres_runoff_future_by_ay_auto_bil_lc,
  eres_runoff_future_by_ay_auto_bil_sc,
  eres_runoff_future_by_ay_auto_pi_lc,
  eres_runoff_future_by_ay_auto_pi_sc,
  eres_runoff_future_by_ay_auto_pdl,
  eres_runoff_future_by_ay_auto_ppa,
  eres_runoff_future_by_ay_auto_odpe,
  eres_runoff_future_by_ay_auto_odpe_ev,
  eres_runoff_future_by_ay_auto_fb,
  eres_runoff_future_by_ay_auto_le,
  eres_runoff_future_by_ay_auto_axap,
  eres_runoff_future_by_ay_auto_eq,
  eres_runoff_future_by_ay_bike_bil_lc,
  eres_runoff_future_by_ay_bike_bil_sc,
  eres_runoff_future_by_ay_bike_pi_lc,
  eres_runoff_future_by_ay_bike_pi_sc,
  eres_runoff_future_by_ay_bike_pdl,
  eres_runoff_future_by_ay_bike_ppa,
  eres_runoff_future_by_ay_bike_le,
  eres_runoff_future_by_ay_pa,
  eres_runoff_future_by_ay_pet
)
write_xlsx(df_eres, "df_eres_runoff_future_by_cov1_v2.xlsx")



# old Motor total ---------------------

# Runoff Eaxa

get_eaxa_runoff_by_AY <- function(exhibit) {
  eaxa_runoff <- exhibit[-c((dim(exhibit)[1]-2):dim(exhibit)[1]), "EAXA"]
  
  return(eaxa_runoff)
}
motor_eaxa_runoff <- get_eaxa_runoff(motor_exhibit)


motor_eaxa_runoff <- motor_exhibit[-c((dim(motor_exhibit)[1]-2):dim(motor_exhibit)[1]), "EAXA"]
auto_bil_lc_exhibit[-c((dim(auto_bil_lc_exhibit)[1]-2):dim(auto_bil_lc_exhibit)[1]), "EAXA"]


# Runoff Paid
# make paid full triangle
product <- "Motor"
cov <- "Total"
cl_type <- "Total"
lkr_type <- "lkr12_paid"
motor_paid_fulltri <- data.frame(get_full_triangle_paid(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type))
# make runoff paid developed
motor_paid_dev_runoff <- get_paid_developed(motor_paid_fulltri)
# Runoff Paid summarize by AY
motor_paid_developed_runoff_by_ay <- get_paid_developed_runoff_by_ay(motor_paid_dev_runoff)
# Runoff EAXA Reserve (EAXA runoff - Paid developed)
motor_eaxares_runoff <- motor_eaxa_runoff - motor_paid_developed_runoff_by_ay[, "PAID"][[1]]


## Future ## 

# Future EAXA
auto_bil_lc_fulltri[is.na(auto_bil_lc_fulltri)] <- 0
get_eaxa_future(auto_bil_lc_fulltri)

auto_bil_lc_paid_projected <- get_paid_future_select_dev_mth(auto_bil_lc_fulltri_paid, 3)
auto_bil_lc_eaxa_future <- get_eaxa_future_select_dev_mth(auto_bil_lc_fulltri, 3)
auto_bil_lc_eaxa_future - auto_bil_lc_paid_projected

# ------------
# Get paid developed 3 months
motor_paid_pcfY_dec <- get_paid_developed_3mth(motor_paid_fulltri)[[1]]
# Paid summarize by AY
tbl <- get_paid_developed_3mth(motor_paid_fulltri)[[2]]
# PCFY Dec EAXA Reserve (EAXA at Sep - Paid at Dec) Runoff
motor_eaxares_pcfy_runoff <- motor_exhibit[-c((dim(motor_exhibit)[1]-2):dim(motor_exhibit)[1]), "EAXA"] - tbl[, "PAID"][[1]] # calc check OK



# PCFY Dec EAXA Reserve future
motor_paid_projected <- get_paid_future(motor_paid_fulltri)
# Motor EAXA by Month for projection
eaxa_lst <- list(auto_bil_lc_fulltri, auto_bil_sc_fulltri, auto_pi_lc_fulltri, auto_pi_sc_fulltri, auto_pdl_fulltri, auto_ppa_fulltri, auto_odpe_fulltri, auto_odpe_ev_fulltri,
                 auto_fb_fulltri, auto_le_fulltri, auto_axap_fulltri, auto_eq_fulltri,
                 bike_bil_lc_fulltri, bike_bil_sc_fulltri, bike_pi_lc_fulltri, bike_pi_sc_fulltri, bike_pdl_fulltri, bike_ppa_fulltri, bike_le_fulltri,
                 pa_fulltri, pet_fulltri)
eaxa_lnm <- list("Auto_BIL_LC", "Auto_BIL_SC", "Auto_PI_LC", "Auto_PI_SC", "Auto_PDL", "Auto_PPA", "Auto_ODPE", "Auto_ODPE_CAT",
                 "Auto_FB", "Auto_LE", "Auto_Axa+", "Auto_EQ",
                 "Bike_BIL_LC", "Bike_BIL_SC",  "Bike_PI_LC", "Bike_PI_SC", "Bike_PDL", "Bike_PPA", "Bike_LE",
                 "PA", "Pet")
eaxa_all_cov <- c()
for(i in seq_along(eaxa_lst)) {
  df <- data.frame(eaxa_lst[[i]])[, names(data.frame(eaxa_lst[[i]])) %in% c('V1', 'Ult')] %>%
    dplyr::mutate(cov=eaxa_lnm[[i]])
  eaxa_all_cov <- rbind(eaxa_all_cov, df)
}
eaxa_all_cov[is.na(eaxa_all_cov)] <- 0
motor_eaxa <- eaxa_all_cov[eaxa_all_cov["cov"] != "PA" & eaxa_all_cov["cov"] != "Pet", ]
# Motor coverage別のEAXAを合算
motor_eaxa_by_mth <- motor_eaxa %>%
  group_by(V1) %>%
  summarise(Ult = sum(Ult))
# EAXA future
motor_eaxa_future <- get_eaxa_future(motor_eaxa_by_mth)
# 予測したEAXAから予測したPaidを引いてリザーブを計算
motor_eaxa_res_future <- motor_eaxa_future - motor_paid_projected
# Motor EAXA Reserveの総額
motor_eaxa_res_pcfy <- sum(motor_eaxares_pcfy_runoff) + sum(motor_eaxa_res_future)


# PA ---------------------
product <- "PA"
cov <- "Total"
cl_type <- "Attrit"
lkr_type <- "lkr12_paid"
pa_paid_fulltri <- data.frame(get_full_triangle_paid(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type))
# Get paid developed 3 months
pa_paid_pcfY_dec <- get_paid_developed_3mth(pa_paid_fulltri)[[1]]
# Paid summarize by AY
pa_tbl <- get_paid_developed_3mth(pa_paid_fulltri)[[2]]
# PCFY Dec EAXA Reserve (EAXA at Sep - Paid at Dec) Runoff
pa_eaxares_pcfy_runoff <- pa_exhibit[-c(1:2, (dim(pa_exhibit)[1]-2):dim(pa_exhibit)[1]), "EAXA"] - pa_tbl[, "PAID"][[1]] # calc check OK
# PCFY Dec EAXA Reserve future
pa_paid_projected <- get_paid_future(pa_paid_fulltri)
pa_eaxa_future <- get_eaxa_future(pa_fulltri)
pa_eaxa_res_pcfy <- get_eaxa_res_pcfy(pa_eaxa_future, pa_paid_projected, pa_eaxares_pcfy_runoff)

# Pet ---------------------
product <- "PET"
cov <- "Total"
cl_type <- "Attrit"
lkr_type <- "lkr12_paid"
pet_paid_fulltri <- data.frame(get_full_triangle_paid(lkr_type = lkr_type, product = product, cov = cov, cl_type = cl_type))
# Get paid developed 3 months
pet_paid_pcfY_dec <- get_paid_developed_3mth(pet_paid_fulltri)[[1]]
# Paid summarize by AY
pet_tbl <- get_paid_developed_3mth(pet_paid_fulltri)[[2]]
# PCFY Dec EAXA Reserve (EAXA at Sep - Paid at Dec) Runoff
pet_eaxares_pcfy_runoff <- pet_exhibit[-c(1:7, (dim(pet_exhibit)[1]-2):dim(pet_exhibit)[1]), "EAXA"] - pet_tbl[, "PAID"][[1]] # calc check OK
# PCFY Dec EAXA Reserve future
pet_paid_projected <- get_paid_future(pet_paid_fulltri)
pet_eaxa_future <- get_eaxa_future(pet_fulltri)
pet_eaxa_res_pcfy <- get_eaxa_res_pcfy(pet_eaxa_future, pet_paid_projected, pet_eaxares_pcfy_runoff[-c(1:4)]) #Petは2009年以前を除外

# Eaxa Reserve by LoB PCFY --------------------------
eaxa_res_pcfy_by_lob <- t(data.frame(motor_eaxa_res_pcfy, pa_eaxa_res_pcfy, pet_eaxa_res_pcfy))
eaxa_res_pcfy_by_lob <- tibble::rownames_to_column(data.frame(eaxa_res_pcfy_by_lob), "LoB")

# -------------------------------
workfile_wd <- "P:/AXADJDivision/RiskManagement/05.Insurance（保険引受リスク）/43.リザービング/2020/Second Opinion/PCFY_December_Projection_test/99 work"
setwd(workfile_wd)

# Paid Full triangleのExport
write_xlsx(motor_paid_fulltri, "motor_paid_fulltri.xlsx")
write_xlsx(pa_paid_fulltri, "pa_paid_fulltri.xlsx")
write_xlsx(pet_paid_fulltri, "pet_paid_fulltri.xlsx")
# Motor Full Triangle
write_xlsx(data.frame(motor_fulltri), "motor_fulltri.xlsx")

write_xlsx(eaxa_res_pcfy_by_lob, "eaxa_res_pcfy_by_lob.xlsx")



# below not used  ------------------------------

# function old version ------
get_paid_developed_3mth <- function(paid_fulltri) {
  paid2 <- paid_fulltri[-c(1:3), !names(paid_fulltri) %in% c('V1', 'X1', 'X2', 'X3', 'Ult')]
  paid_dec <- c()
  for(i in 1:dim(paid2)[[1]]) {
    bal <- paid2[dim(paid2)[[1]]+1-i, i]
    paid_dec <- rbind(paid_dec, bal)
  }
  paid_pcfY_dec <- c(paid_fulltri[c(1:3), 'Ult'], rev(paid_dec))
  paid_pcfY_dec <- data.frame(cbind(paid_fulltri[, 1], paid_pcfY_dec))
  colnames(paid_pcfY_dec) <- c("ym", 'PAID')
  
  tbl <- paid_pcfY_dec %>%
    mutate(AY=substring(ym, 1, 4)) %>%
    group_by(AY) %>%
    summarise(PAID=sum(PAID))
  
  return(list(paid_pcfY_dec, tbl))
}


get_paid_future <- function(paid_fulltri) {
  paid_cy <- sum(paid_fulltri[c((dim(paid_fulltri)[1]-2):dim(paid_fulltri)[1]), c(2)])
  paid_py <- sum(paid_fulltri[c((dim(paid_fulltri)[1]-14):(dim(paid_fulltri)[1]-12)), c(2)])
  paid_growth_rate <- paid_cy / paid_py
  paid_future <- paid_fulltri[c((dim(paid_fulltri)[1]-11):(dim(paid_fulltri)[1]-9)), c(2:4)] * paid_growth_rate
  paid_projected <- c(paid_future[1, 3], paid_future[2, 2], paid_future[3, 1])
  
  return(paid_projected)
}

get_eaxa_future <- function(fulltri) {
  eaxa_by_mth <- data.frame(fulltri)[, names(data.frame(fulltri)) %in% c('V1', 'Ult')]
  eaxa_cy <- sum(eaxa_by_mth[c((dim(eaxa_by_mth)[1]-2):dim(eaxa_by_mth)[1]), c(2)])
  eaxa_py <- sum(eaxa_by_mth[c((dim(eaxa_by_mth)[1]-14):(dim(eaxa_by_mth)[1]-12)), c(2)])
  eaxa_growth_rate <- eaxa_cy / eaxa_py
  eaxa_future <- eaxa_by_mth[c((dim(eaxa_by_mth)[1]-11):(dim(eaxa_by_mth)[1]-9)), 2] * eaxa_growth_rate
  
  return(eaxa_future)
}

get_paid_developed <- function(paid_fulltri) {
  paid2 <- paid_fulltri[-c(1:3), !names(paid_fulltri) %in% c('V1', 'X1', 'X2', 'X3', 'Ult')]
  paid_dev <- c()
  for(i in 1:dim(paid2)[[1]]) {
    bal <- paid2[dim(paid2)[[1]]+1-i, i]
    paid_dev <- rbind(paid_dev, bal)
  }
  paid_dev_runoff <- c(paid_fulltri[c(1:3), 'Ult'], rev(paid_dev))
  paid_dev_runoff <- data.frame(cbind(paid_fulltri[, 1], paid_dev_runoff))
  colnames(paid_dev_runoff) <- c("ym", 'PAID')
  
  
  return (paid_dev_runoff)
}

get_eaxa_res_runoff <- function(fulltri_paid, dev_mth, fulltri_inc, cov_name) {
  runoff_paid_developed <- get_runoff_paid_developed_select_dev_mth(data.frame(fulltri_paid), dev_mth)[, 'PAID']
  runoff_paid_developed <- replace(runoff_paid_developed, is.na(runoff_paid_developed), 0)
  runoff_eaxa <- replace(fulltri_inc[, 'Ult'], is.na(fulltri_inc[, 'Ult']), 0)
  runoff_eres_developed <- runoff_eaxa - runoff_paid_developed
  runoff_eres_developed <- data.frame(runoff_eres_developed) %>% 
    mutate(cov_name = cov_name, ym = fulltri_inc[, 1], eaxa = runoff_eaxa, paid = runoff_paid_developed) %>% 
    rename(eres = runoff_eres_developed)
  
  return(runoff_eres_developed)
}

