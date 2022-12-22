#setting the directory to working file
library(rstudioapi) 
setwd(dirname(getActiveDocumentContext()$path))  

#loading necessary packages/libraries
library(dplyr)
library(ggplot2)
library(lubridate) #for dealing with dates and times
library(janitor) #for cleaning names
library(scales) #extending for ggplot2
library(plotly) #for interactive graphs
library(readxl) #for reading excel file   
library(data.table)
#library(rJava)
#library(xlsx)
library(openxlsx)
library(readxl)
library(writexl)

index1<-c(6:11)
n=2
#load Data into R 
E1 <- data.frame(read_excel("data/E1.xlsx", sheet = "BenchMark",skip = n))
E1 <- E1 %>% 
  mutate_at(index1, as.numeric)

E2 <- data.frame(read_excel("data/E2.xlsx", sheet = "BenchMark",skip = n))

E2 <- E2 %>% 
  mutate_at(index1, as.numeric)

E3 <- data.frame(read_excel("data/E3.xlsx", sheet = "BenchMark",skip = n))
E3 <- E3 %>% 
  mutate_at(index1, as.numeric)

E4 <- data.frame(read_excel("data/E4.xlsx", sheet = "BenchMark",skip = n))
E4 <- E4 %>% 
  mutate_at(index1, as.numeric)
E5 <- data.frame(read_excel("data/E5.xlsx", sheet = "BenchMark",skip = n))
E5 <- E5 %>% 
  mutate_at(index1, as.numeric)

E6 <- data.frame(read_excel("data/E6.xlsx", sheet = "BenchMark",skip = n))
E6 <- E6 %>% 
  mutate_at(index1, as.numeric)

E7 <- data.frame(read_excel("data/E7.xlsx", sheet = "BenchMark",skip = n))

E7 <- E7 %>% 
  mutate_at(index1, as.numeric)


E8 <- data.frame(read_excel("data/E8.xlsx", sheet = "BenchMark",skip = n))
E8 <- E8 %>% 
  mutate_at(index1, as.numeric)


E9 <- data.frame(read_excel("data/E9.xlsx", sheet = "BenchMark",skip = n))
E9 <- E9 %>% 
  mutate_at(index1, as.numeric)

E10 <- data.frame(read_excel("data/E10.xlsx", sheet = "BenchMark",skip = n))
E10 <- E10 %>% 
  mutate_at(index1, as.numeric)

#set the base file
B1 <-E1%>%
  select(ID,Level.1,Level.2,Level.3,Level.4)%>%
  mutate(B1.E1=E1$B1,
         B1.E2=E2$B1,
         B1.E3=E3$B1,
         B1.E4=E4$B1,
         B1.E5=E5$B1,
         B1.E6=E6$B1,
         B1.E7=E7$B1,
         B1.E8=E8$B1,
         B1.E9=E9$B1,
         B1.E10=E10$B1)



wb <- loadWorkbook("data/Main.xlsx")
writeData(wb, sheet = "E1", E1)
writeData(wb, sheet = "E2", E2)
writeData(wb, sheet = "E3", E3)
writeData(wb, sheet = "E4", E4)
writeData(wb, sheet = "E5", E5)
writeData(wb, sheet = "E6", E6)
writeData(wb, sheet = "E7", E7)
writeData(wb, sheet = "E8", E8)
writeData(wb, sheet = "E9", E9)
writeData(wb, sheet = "E10", E10)
saveWorkbook(wb,"data/Main.xlsx",overwrite = T)