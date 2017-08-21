#Überprüfe und importiere Einzeltierdaten aus ZH
setwd("/Users/christoph_user/Dropbox/SwissBat/Zürich/Swissbat/5_separate_Einzeldaten")
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

head(d.m)


#Morphologiedaten
d.m <- within(d.m, {
  Col.Obj.Attribut.Remarks.x <- gsub("\\s", "", Col.Obj.Attribut.Remarks.x)
  Col.Obj.Attribut.Remarks.y <- gsub("\\s", "", Col.Obj.Attribut.Remarks.y)
})

d.morph.d_1 <- as.data.frame(do.call(rbind,strsplit(d.m$Col.Obj.Attribut.Remarks.x, "/")))
d.morph.d_2 <- as.data.frame(do.call(rbind,strsplit(d.m$Col.Obj.Attribut.Remarks.y, "/")))
colnames(d.morph.d_1) <- c("FOREA", "Digi1", "Digi3", "Digi5")
colnames(d.morph.d_2) <- c("Digi1", "Digi3", "Digi5", "FOREA")
head(d.morph.d_1)
head(d.morph.d_2)

#Überprüfe ob Reihenfolge jeweils korrekt
unique(as.data.frame(do.call(rbind,strsplit(as.character(d.morph.d_1$FOREA), ",")))[,1])
unique(as.data.frame(do.call(rbind,strsplit(as.character(d.morph.d_1$Digi1), ",")))[,1])
unique(as.data.frame(do.call(rbind,strsplit(as.character(d.morph.d_1$Digi3), ",")))[,1])
unique(as.data.frame(do.call(rbind,strsplit(as.character(d.morph.d_1$Digi5), ",")))[,1])

unique(as.data.frame(do.call(rbind,strsplit(as.character(d.morph.d_2$FOREA), ",")))[,1])
unique(as.data.frame(do.call(rbind,strsplit(as.character(d.morph.d_2$Digi1), ",")))[,1])
unique(as.data.frame(do.call(rbind,strsplit(as.character(d.morph.d_2$Digi3), ",")))[,1])
unique(as.data.frame(do.call(rbind,strsplit(as.character(d.morph.d_2$Digi5), ",")))[,1])


#ersetze Text
d.morph.d_1 <- within(d.morph.d_1, {
  FOREA <- gsub("FOREA\\(Unterarmlänge\\),", "", as.character(FOREA))
  Digi1 <- gsub("Digi1\\(fingerbone1\\),", "", as.character(Digi1))
  Digi3 <- gsub("Digi1\\(fingerbone3\\),", "", as.character(Digi3))
  Digi5 <- gsub("Digi1\\(fingerbone5\\),", "", as.character(Digi5))
})

d.morph.d_2 <- within(d.morph.d_2, {
  FOREA <- gsub("FOREA\\(Unterarmlänge\\)", "", as.character(FOREA))
  Digi1 <- gsub("Digi1\\(fingerbone1\\)", "", as.character(Digi1))
  Digi3 <- gsub("Digi3\\(fingerbone3\\)", "", as.character(Digi3))
  Digi5 <- gsub("Digi5\\(fingerbone5\\)", "", as.character(Digi5))
})


#konvertiere zu numerisch
d.morph.d_1 <- within(d.morph.d_1, {
  FOREA <- as.numeric(FOREA)
  Digi1 <- as.numeric(Digi1)
  Digi3 <- as.numeric(Digi3)
  Digi5 <- as.numeric(Digi5)
})

d.morph.d_2 <- within(d.morph.d_2, {
  FOREA <- as.numeric(FOREA)
  Digi1 <- as.numeric(Digi1)
  Digi3 <- as.numeric(Digi3)
  Digi5 <- as.numeric(Digi5)
})


#füge Datensätze d_1 und d_2 zusammen
#prüfe ob keine mehrfachen Werte in d_1 und d_2 
unique(subset(d.morph.d_2, !is.na(d.morph.d_1$FOREA))["FOREA"])
unique(subset(d.morph.d_2, !is.na(d.morph.d_1$Digi1))["Digi1"])
unique(subset(d.morph.d_2, !is.na(d.morph.d_1$Digi3))["Digi3"])
unique(subset(d.morph.d_2, !is.na(d.morph.d_1$Digi5))["Digi5"])

unique(subset(d.morph.d_1, !is.na(d.morph.d_2$FOREA))["FOREA"])
unique(subset(d.morph.d_1, !is.na(d.morph.d_2$Digi1))["Digi1"])
unique(subset(d.morph.d_1, !is.na(d.morph.d_2$Digi3))["Digi3"])
unique(subset(d.morph.d_1, !is.na(d.morph.d_2$Digi5))["Digi5"])


d.morph.d_1$FOREA[is.na(d.morph.d_1$FOREA)] <- d.morph.d_2$FOREA[is.na(d.morph.d_1$FOREA)]
d.morph.d_1$Digi1[is.na(d.morph.d_1$Digi1)] <- d.morph.d_2$Digi1[is.na(d.morph.d_1$Digi1)]
d.morph.d_1$Digi3[is.na(d.morph.d_1$Digi3)] <- d.morph.d_2$Digi3[is.na(d.morph.d_1$Digi3)]
d.morph.d_1$Digi5[is.na(d.morph.d_1$Digi5)] <- d.morph.d_2$Digi5[is.na(d.morph.d_1$Digi5)]



#füge Datensätze d.m und d_1 zusammen
#prüfe Übereinstimmungen zwischen d.m und d_1; falls keine Übereinstimmung, FALSE Werte
unique(d.m$FOREA == d.morph.d_1$FOREA)
unique(d.m$Digi1 == d.morph.d_1$Digi1)
unique(d.m$Digi3 == d.morph.d_1$Digi3)
unique(d.m$Digi5 == d.morph.d_1$Digi5)


d.m$FOREA[is.na(d.m$FOREA)] <- d.morph.d_1$FOREA[is.na(d.m$FOREA)]
d.m$Digi1[is.na(d.m$Digi1)] <- d.morph.d_1$Digi1[is.na(d.m$Digi1)]
d.m$Digi3[is.na(d.m$Digi3)] <- d.morph.d_1$Digi3[is.na(d.m$Digi3)]
d.m$Digi5[is.na(d.m$Digi5)] <- d.morph.d_1$Digi5[is.na(d.m$Digi5)]



#d.m
names(d.m)
head(subset(d.m, count.x>1))
