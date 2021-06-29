library(openxlsx)

if(!dir.exists("data/tidy")) dir.create("data/tidy") # make dir if doesn't exist

################################################################################
# tidying
# ids
ids           <- read.xlsx("src/Идентификаторы.xlsx", sheet = "Идентификаторы")
names(ids)    <- ids[1,]
ids           <- ids[-1,]
ids$`№`       <- NULL
write.xlsx(x = ids, file = "data/tidy/Идентификаторы.xlsx", sheetName = "Идентификаторы")

# demography
demography         <- read.xlsx("src/Демографические данные.xlsx", sheet = "Демографические данные")
names(demography)  <- demography[1,]
demography         <- demography[-1,]
demography$`№`     <- NULL
demography$Возраст <- as.integer(demography$Возраст)
write.xlsx(x = demography, file = "data/tidy/Демографические данные.xlsx", sheetName = "Демографические данные")

# anthropometry  
anthropometry             <- read.xlsx("src/Антропометрия.xlsx", sheet = "Антропометрические данные")
names(anthropometry)      <- anthropometry[2,]
anthropometry             <- anthropometry[-2:-1,]
anthropometry$`№`         <- NULL
anthropometry$Температура <- as.numeric(anthropometry$Температура)
anthropometry$Рост        <- as.numeric(anthropometry$Рост)
anthropometry$Вес         <- as.numeric(anthropometry$Вес)
write.xlsx(x = anthropometry, file = "data/tidy/Антропометрические данные.xlsx", sheetName = "Антропометрические данные")

# biochemistry
biochemistry          <- read.xlsx("src/Биохимический анализ крови.xlsx", sheet = "Биохимический анализ крови")
names(biochemistry)   <- biochemistry[3,]
biochemistry          <- biochemistry[-3:-1,]
biochemistry$`№`      <- NULL
biochemistry$Белок    <- as.numeric(biochemistry$Белок)
biochemistry$Глюкоза  <- as.numeric(biochemistry$Глюкоза)
biochemistry$Мочевина <- as.numeric(biochemistry$Мочевина)

wb <- createWorkbook()
addWorksheet(wb, sheetName = "Биохимический анализ крови")
writeData(wb, biochemistry, sheet = "Биохимический анализ крови")
addStyle(wb, sheet = "Биохимический анализ крови", style = createStyle(numFmt = "NUMBER"), rows = 1:nrow(biochemistry) + 1, cols = 2:4, gridExpand = TRUE)
saveWorkbook(wb, file = "data/tidy/Биохимический анализ крови.xlsx", overwrite = TRUE)

write.xlsx(x = biochemistry, file = "data/tidy/Биохимический анализ крови.xlsx", sheetName = "Биохимический анализ крови")

################################################################################
# consolidation
options(digits=10) # to correctly format numbers wo extra digits

# read tidy data
ids           <- read.xlsx("data/tidy/Идентификаторы.xlsx",             sheet = "Идентификаторы")
demography    <- read.xlsx("data/tidy/Демографические данные.xlsx",     sheet = "Демографические данные")
anthropometry <- read.xlsx("data/tidy/Антропометрические данные.xlsx",  sheet = "Антропометрические данные")
biochemistry  <- read.xlsx("data/tidy/Биохимический анализ крови.xlsx", sheet = "Биохимический анализ крови")

# prepare workbook
db.intermediate <- createWorkbook()
addWorksheet(db.intermediate, sheetName = "ids")           : writeData(db.intermediate, sheet = "ids",           ids)     
addWorksheet(db.intermediate, sheetName = "demography")    : writeData(db.intermediate, sheet = "demography",    demography)
addWorksheet(db.intermediate, sheetName = "anthropometry") : writeData(db.intermediate, sheet = "anthropometry", anthropometry)
addWorksheet(db.intermediate, sheetName = "biochemistry")  : writeData(db.intermediate, sheet = "biochemistry",  biochemistry)
saveWorkbook(db.intermediate, "data/db.intermediate.xlsx", overwrite = TRUE)

# joining
d <- dplyr::left_join(ids, demography,    by = c("Демографические.данные"      = "Пациент"))
d <- dplyr::left_join(d,   anthropometry, by = c("Антропометрические.данные"   = "Пациент"))
d <- dplyr::left_join(d,   biochemistry,  by = c("Биохимический.анализ.крови"  = "Пациент"))

# saving of working dataset

d$Демографические.данные     <- NULL
d$Антропометрические.данные  <- NULL
d$Биохимический.анализ.крови <- NULL

names(d) <- c("patient.id", "sex", "age", "group", "temp", "height", "weight", "protein", "glucose", "urea")

write.xlsx(x = d, file = "data/working.dataset.xlsx", sheetName = "working.dataset")
