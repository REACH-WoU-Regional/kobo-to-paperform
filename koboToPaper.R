if (!require("pacman")) install.packages("pacman")
pacman::p_load(sf,tidyverse, readxl, cowplot, DT, dplyr, utils, rlang,expss, ggplot2, data.table, openxlsx)

# Set directory
setwd(dirname(rstudioapi::getActiveDocumentContext()$path))

rm(list = ls())

##Read questions
koboQues <- read_excel("data/hsosKoboTool.xlsx", sheet = "survey")

##Read Choices
koboChoices <- read_excel("data/hsosKoboTool.xlsx", sheet = "choices")

##Creating table 
exTableEnglish <- koboQues %>% 
  select(relevant,type,`label::english`,`hint::english`,`constraint_message::english`) %>% 
  rename("Type" = type, "Name" = `label::english`, "Hint" = `hint::english`, "Constraint" = `constraint_message::english`, "Relevancy" = relevant) %>% 
  filter(!(is.na(Name))) %>% 
  filter(!(Name %in% c("start","end","today"))) %>% 
  mutate(Type = ifelse(str_detect(Type, "select_one") == T, gsub("select_one ","",Type),
                       ifelse(str_detect(Type, "select_multiple") == T, gsub("select_multiple ","",Type),Type)))


##Choices table
exChoicesEnglish <- koboChoices %>% 
  select(list_name, `label::english`) %>% 
  filter(!(is.na(list_name))) %>% 
  group_by(list_name) %>% 
  summarise_all(funs(toString(na.omit(.)))) %>% 
  rename("Type" = list_name, "Choices English" = `label::english`)

##combined table
paperFormEnglish <- exTableEnglish %>% 
  left_join(exChoicesEnglish, by= "Type") %>% 
  filter(!(Type %in% c("end_group")))


##Writing to excel and styling
#Creating the workbook
wb <- createWorkbook()

#Adding the worksheet
addWorksheet(wb, sheetName = "paperEnglish")


##Adding data Headers
for (i in 1:length(colnames(paperFormEnglish))){
  writeData(wb, "paperEnglish", x = colnames(paperFormEnglish)[i], startRow = 1, startCol = i)  
}

##Adding data table
for (i in 1:nrow(paperFormEnglish)){
  writeData(wb, "paperEnglish", x = paperFormEnglish[i,], startRow = 1 + i, col.names = F)
}


#Styling
columnName <- createStyle(fontName = "Arial Narrow", fontSize = 12,
                          wrapText = T, valign = "top")

addStyle(wb, "paperEnglish", columnName, cols = 2, rows = 2:sum(nrow(paperFormEnglish) + 1), gridExpand = T)



##Column Width and Row width
setColWidths(wb, sheet = "paperEnglish", cols = 2, widths = 75)

saveWorkbook(wb, "output.xlsx", overwrite = T)
