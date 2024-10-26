library(readxl)
library(tidyverse)
library(tidyr)
library(data.table)
library(writexl)
library(DBI)
library(odbc)
library(writexl)
library(xlsx)

setwd("//pnjkncifs01/202_Common2/ECM/JapanModel/2024/Jun2024/Actual/10_Data")

##########
con <- dbConnect(odbc(),
                 Driver = "SQL Server",
                 Server = "DWAWDB245401,60101",
                 Database = "ICMDEV_5_3_JAPAN",
                 Trusted_Connections = "Yes"
)
qry_capital = "select [RunID],[Group],[Allocation],[Entity],[RiskLevel],[Risk],[Value] 
from [ICM].[Outputs_Capital_GroupAllocationByEntityRisk] 
where Allocation IN ('Standalone VaR','Standalone TVaR','One-Year Co-TVaR')"
df <- dbGetQuery(con, qry_capital)

qry_opr_oth = "SELECT * FROM [ICM].[Global_UserDefinedDistns_Names$DIM] a 
Inner join [ICM].[Global_UserDefinedDistns_Parameters$DATA] b 
on a.ID = b.Global_UserDefinedDistns_Names$ID"
df_opr_oth <- dbGetQuery(con, qry_opr_oth)

qry_data_group_versions <- "select [Version], [RunID], [DataGroupID] FROM [ICMDEV_5_3_JAPAN].[WTW].[DATAGROUP_VERSIONS]"
df_data_group_versions <- dbGetQuery(con, qry_data_group_versions)

qry_cat <- "select * from [ICM].[Outputs_Cat] where OriginPeriodBasis = 'Accident'"
df_cat <- dbGetQuery(con, qry_cat)

df_corr_matrix <- read.csv("Correlation_Matrix.csv")
#########
# Read SQL data
#df <- read.csv("June_Actual_Igloo_output_241023.csv", sep = "\t")
#df <- read.csv("Outputs_Capital_GroupAllocationByEntityRisk_241024.csv")
#df_opr_othold <- read.csv("opr_other.csv")
#df_data_group_versions <- read.csv("Data_Group_Versions.csv")
#df_cat <- read.csv("CAT_241025.csv")
##########

df_cat$RunID <- as.numeric(df_cat$RunID)
df_cat2 <- df_cat %>% 
  filter(RunID > 10300 & 
           Class == "Total Insurance" & 
           OriginPeriodBasis == "Accident" & 
           Stat == "P99.5" ) %>% 
  mutate(Group = "AJH", Allocation = "Standalone VaR") %>% 
  rename(RiskLevel = Peril, Risk = Category) %>% 
  select(RunID, Group, Allocation, Entity, RiskLevel, Risk, Value)
df_cat2$Value <- as.numeric(df_cat2$Value)



df_data_group_versions <- df_data_group_versions %>% 
  filter(DataGroupID == 3)


colnames(df_opr_oth)[5] <- "Version2"
colnames(df_opr_oth)[9] <- "Value2"


df_merge <- merge(x = df_opr_oth, y = df_data_group_versions, by.x = "Version", by.y = "Version", all.x = T) 

df_merge2 <- df_merge %>% mutate(paramID = str_sub(`Const_Ins_UDParams$ID`, start = -1, end = -1)) %>% 
  filter(RunID > 10300 & paramID == 1 & Version == Version2 & Position != 10) %>% 
  select(RunID, ID, Position, Value, Value2) %>% 
  arrange(RunID, Position)

df_merge3 <- df_merge2 %>% mutate(
  Entity = str_sub(Value, start = -3, end = -1),
  Group = "AJH",
  Allocation = "Standalone VaR",
  RiskLevel = "5"
              ) %>% 
  rename(Risk = Value, Value =Value2) %>% 
  select(RunID, Group, Allocation, Entity, RiskLevel, Risk, Value) %>% 
  filter(Entity != "isk" & Entity != "sks")
df_merge3$Entity[df_merge3$Entity == "npo"] <- "Sonpo"
df_merge3$Entity[df_merge3$Entity == "AJH"] <- "Total"
df_merge3$Value <- as.numeric(df_merge3$Value)

# Outputs_Capital_GroupAllocationByEntityRisk
df2 <- df %>% 
  filter(
    RunID > 10300 & 
    Allocation == "Standalone VaR" &
    ((Risk == "Insurance Risk" & RiskLevel == 4) | 
       Risk == "Premium Risk Non-Cat" | 
       (Risk == "Reserve Risk" & RiskLevel == 3) |       
      (Risk == "Catastrophe Risk" & RiskLevel == 2) |
       (Risk == "Credit Risk - RI Default" & RiskLevel == 3) |
       (Risk == "Market Risk" & RiskLevel == 4) |
       (Risk == "Interest Rate" & RiskLevel == 1) |
       (Risk == "Equity" & RiskLevel == 1) |
       (Risk == "Spread" & RiskLevel == 1) |
       (Risk == "Credit" & RiskLevel == 1) |
       (Risk == "Non-Insurance Currency Risk" & RiskLevel == 1) |
       (Risk == "Liquidity Risk" & RiskLevel == 1) |
       (Risk == "Tax" & RiskLevel == 1) |
       (Risk == "Total" & RiskLevel == 5) 
     )
           )

# Outputs_Capital_GroupAllocationByEntityRiskとOperational, other assets & liabilitiesを統合
df2$RunID <- as.character(df2$RunID)
df_merge3$RunID <- as.character(df_merge3$RunID)

df3 <- rbind(df2, df_merge3)

# IglooデータからTotal Capital Requirementを計算

# MOdeled Risks, Operational risk, Other assets & liabilityを抽出
df_agg <- df3 %>% 
  filter(RunID > 10300 & Risk == "Total" | str_detect(Risk,"Operational Risk") | str_detect(Risk,"Other"))
df_agg$RunID <- as.character(df_agg$RunID)

# RunIDごとにTotal Capital Requirementを計算
calc_tcr <- function(runid, entity) {
  df_runid <- df_agg %>% 
    filter(RunID == runid & Entity == entity)
  
  TCR <- sqrt(t(df_runid$Value) %*% as.matrix(df_corr_matrix[,-1]) %*% df_runid$Value)
  df_tcr <- data.frame(RunID = runid, 
                       Group = "AJH",
                       Allocation = "Standalone VaR",
                       Entity = entity, 
                       RiskLevel = "6",
                       Risk = "Total Capital Requirement",
                       Value = TCR)
  return(df_tcr)
}

df_tcr_list <- list()
count <- 1
for (i in unique(df_agg$RunID)) {
  for (j in unique(df_agg$Entity)) {
      
      skip_to_next <- FALSE
      tryCatch({
        tcr <- calc_tcr(i, j)
        df_tcr_list[[count]] <- tcr
        count <- count +1
      }, error = function(e) {
        skip_to_next <- TRUE
      })
      if(skip_to_next) {next}
  }
}
df_tcr <- do.call(rbind, df_tcr_list)


df_with_tcr <- rbind(df3, df_tcr)
df_with_tcr_cat <- rbind(df_with_tcr, df_cat2)

# Wideにする
wide_test <- pivot_wider(df_with_tcr_cat,
                         names_from = RunID,
                         values_from = Value)

write_xlsx(wide_test, "Outputs_Capital_GroupAllocationByEntityRisk_241025_v3.xlsx")
