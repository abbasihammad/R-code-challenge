install.packages("dplyr")
install.packages(c("readxl","writexl")) 
install.packages("openxlsx")
library(dplyr)
library(xlsx)
library(readxl)
library(writexl)
library(openxlsx)

setwd("C:/Users/Hammad/Desktop/UM/New folder")
file <- "sessionCounts.csv"
file2 <-"addsToCart.csv"

sessioncnt <- read.table(file = file, header= TRUE, sep = ",")
addsToCart <- read.table(file = file2, header= TRUE, sep = ",")
sessioncnt$dim_date<- as.numeric(strftime(sessioncnt$dim_date, "%m"))

output1 <- sessioncnt %>% 
  group_by(Month = dim_date, "Device Category" = dim_deviceCategory) %>% 
  summarise(Sessions=sum(sessions), Transactions = sum(transactions), QTY = sum(QTY)) %>%
  mutate(ECR = Transactions/Sessions) 

output2 <- sessioncnt %>% 
  group_by(Month = dim_date) %>% 
  summarise(Sessions=sum(sessions), Transactions = sum(transactions), QTY = sum(QTY)) %>%
  mutate(ECR = Transactions/Sessions)

output2[["Device Category"]] <- c("Overall")
output3 <- output2[c(1,6,2,3,4,5)]
output4 <- dplyr::left_join(output3, addsToCart, by=c("Month" = "dim_month"))
output5 <- output4[c(1,2,3,4,5,7,6)]
colnames(output5)[6] <- "Adds To Cart"


hs2<- createStyle(halign = "left",
        fgFill = "#E5E5E5", border="Bottom", borderColour = "#E5E5E5", textDecoration = "BOLD",
        wrapText = TRUE,
        borderStyle = getOption("openxlsx.borderStyle", "thin"))

wbss <- createWorkbook()
addWorksheet(wbss, 'Month x Device', gridLines = FALSE)
setColWidths(wbss, sheet ="Month x Device", cols = 1:6, widths = 11)
writeData(wbss, sheet = "Month x Device", x = output1, headerStyle = hs2) 

addWorksheet(wbss, "Month Summary", gridLines = FALSE)
setColWidths(wbss, sheet = "Month Summary", cols = 1:7, widths = 11) 
writeData(wbss, sheet = "Month Summary", x = output5, headerStyle = hs2)  
saveWorkbook(wbss, "output.xlsx",  overwrite = TRUE)  











          

