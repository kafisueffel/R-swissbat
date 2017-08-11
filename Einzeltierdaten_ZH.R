#First Test
setwd("/Users/christoph_user/Dropbox/SwissBat/Z?rich/Swissbat/5_separate_Einzeldaten")
setwd("C:/Users/kkr/Dropbox_CT/Dropbox/SwissBat/Zürich/Swissbat/5_separate_Einzeldaten")
require(XLConnect)
wb_1 <- loadWorkbook("IMPORT_Fledermaus_v1.xls")
d.d_1 <- readWorksheet(wb_1, 1)
head(d.d_1)

wb_2 <- loadWorkbook("IMPORT_Fledermaeuse2_v1.xls")
d.d_2 <- readWorksheet(wb_2, 1)
head(d.d_2)

wb_3 <- loadWorkbook("SammlungStutz.xlsx")
d.d_3 <- readWorksheet(wb_3, 1)
head(d.d_3)

d.m <- merge(d.d_1, d.d_2, by = c("Catalogue.number", "Alt.Cat.Number"), all = TRUE)
head(d.m)

d.m <- within(d.m, Alt.Cat.Number <- gsub("Colnu ", "", Alt.Cat.Number))

d.m <- merge(d.m, d.d_3, by.x = "Alt.Cat.Number", by.y = "Colnu", all = TRUE)
d.m <- within(d.m, {
  SPENU <- paste(Genus.x, Species.x)
  SPENU2 <- paste(Genus.y, Species.y)
  })
head(d.m)


subset(d.m, Family.x != )