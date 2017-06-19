##TEXAS RAIL ROAD COMMISSION DATA##
rm(list=ls())
library(XLConnect)
library(readxl)
library(plyr)
wb<-loadWorkbook("texas.xlsx")
sheet_names<-getSheets(wb)
cum<-data.frame()
for ( i in 1:length(sheet_names))
{
  name <- paste("Texas_",trimws(sheet_names[i]))
  abc<-read_xlsx("texas.xlsx", sheet=sheet_names[i])
  
  if(ncol(abc) > 9)
  {
    abc<-abc[,c(1:9)]
  } else {
    abc<-abc
  }
  abc<-abc[!is.na(abc),]
  
  colnames(abc)<-c("Material", "Code","Storage", "Receipts", "Input", "Products", "Fuel_Used", "Deliveries", "Storage_EOM")

  abc$Month<-''
  abc$Refinery<-sheet_names[i]
  
  for(j in 2:nrow(abc))
  {
    if(grepl('[a-zA-Z]',abc$Deliveries[j]))
    {
      abc$Month[j]<-abc$Deliveries[j]
    } else {
      abc$Month[j]<-abc$Month[j-1]
    }
  }
  
  abc<-abc[!is.na(abc$Code),]
  
  for( k in 1:nrow(abc))
  {
    if(grepl('[a-zA-Z]',abc$Code[k]))
    {
      abc<-abc[-k,]
    }
  }
  
  cum<-rbind(abc,cum)
  
  assign(name, abc)
}

write.csv(cum,"Texas_Cum.csv")

summary_texas<-ddply(cum, c("Material", "Refinery"), summarise,sum(Storage),sum(Receipts), sum( Input), sum(Products))

