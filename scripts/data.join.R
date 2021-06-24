library(openxlsx)

# read source data
ids           <- read.xlsx("src/Идентификаторы.xlsx",             sheet = "Идентификаторы")
demography    <- read.xlsx("src/Демографические данные.xlsx",     sheet = "Демографические данные")
anthropometry <- read.xlsx("src/Антропометрия.xlsx",              sheet = "Антропометрические данные")
biochemistry  <- read.xlsx("src/Биохимический анализ крови.xlsx", sheet = "Биохимический анализ крови")

# consolidation
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
