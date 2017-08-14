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

for(i in 1:3){
  print(i)
  print(names(eval(as.name(paste("d.d_",i, sep = "")))))
}

d.m <- within(d.m, {
  SPENU1 <- paste(Genus.x, Species.x)
  SPENU2 <- paste(Genus.y, Species.y)
  
  SPENU1[SPENU1 == "NA NA"] <- NA
  SPENU1 <- gsub(" NA$", "", SPENU1)
  SPENU1 <- gsub("\\s", "", SPENU1)
  
  SPENU2[SPENU2 == "NA NA"] <- NA
  SPENU2 <- gsub(" NA$", "", SPENU2)
  SPENU2 <- gsub("\\s", "", SPENU2)
  
  SPENU3 <- gsub("\\s", "", SPENU)
  })

head(d.m)
unique(d.m$SPENU1)
unique(d.m$SPENU2)
unique(d.m$SPENU3)

subset(d.m, is.na(SPENU1) & is.na(SPENU2) & is.na(SPENU3))

#überprüfe Artnamen, die möglicherweise falsch sind
d.nomatch <- subset(d.m, SPENU1 !=  SPENU2 | SPENU1 !=  SPENU3 | SPENU2 !=  SPENU3)
subset(d.nomatch[,c("SPENU1", "SPENU2", "SPENU3")], is.na(SPENU1))
subset(d.nomatch[,c("SPENU1", "SPENU2", "SPENU3")], is.na(SPENU2))
subset(d.nomatch[,c("SPENU1", "SPENU2", "SPENU3")], is.na(SPENU3))


#überprüfe Morphologiedaten, die nicht übereinstimmen
#ungleiche Werte in SammlungStutz.xlsx und IMPORT_Fledermaus_v1.xls
d.check1 <- subset(d.m, (mapply(grepl, as.character(FOREA), as.character(Col.Obj.Attribut.Remarks.x))==0 |
                           mapply(grepl, as.character(Digi1), as.character(Col.Obj.Attribut.Remarks.x))==0 |
                           mapply(grepl, as.character(Digi3), as.character(Col.Obj.Attribut.Remarks.x))==0 |
                           mapply(grepl, as.character(Digi5), as.character(Col.Obj.Attribut.Remarks.x))==0) &
                     !is.na(SPENU1) &
                     !is.na(Col.Obj.Attribut.Remarks.x))

#NAs in SammlungStutz.xlsx, aber nicht in IMPORT_Fledermaus_v1.xls
d.check1 <- rbind(d.check1,subset(d.m, is.na(FOREA) &
                                    is.na(Digi1) &
                                    is.na(Digi3) &
                                    is.na(Digi5) &
                                    !is.na(Col.Obj.Attribut.Remarks.x) &
                                    Col.Obj.Attribut.Remarks.x != "FOREA (Unterarmlänge), 0/Digi1 (finger bone 1), 0/ Digi1 (finger bone 3), 0/Digi1 (finger bone 5), 0"))


#ungleiche Werte in SammlungStutz.xlsx und IMPORT_Fledermaeuse2_v1.xls
d.check2 <- subset(d.m, (mapply(grepl, as.character(FOREA), as.character(Col.Obj.Attribut.Remarks.y))==0 |
                           mapply(grepl, as.character(Digi1), as.character(Col.Obj.Attribut.Remarks.y))==0 |
                           mapply(grepl, as.character(Digi3), as.character(Col.Obj.Attribut.Remarks.y))==0 |
                           mapply(grepl, as.character(Digi5), as.character(Col.Obj.Attribut.Remarks.y))==0) &
                     !is.na(SPENU2) &
                     !is.na(Col.Obj.Attribut.Remarks.y))

#NAs in SammlungStutz.xlsx, aber nicht in IMPORT_Fledermaeuse2_v1.xls
d.check2 <- rbind(d.check2, subset(d.m, is.na(FOREA) &
                                     is.na(Digi1) &
                                     is.na(Digi3) &
                                     is.na(Digi5) &
                                     !is.na(Col.Obj.Attribut.Remarks.y) &
                                     Col.Obj.Attribut.Remarks.y != "FOREA (Unterarmlänge), 0/Digi1 (finger bone 1), 0/ Digi1 (finger bone 3), 0/Digi1 (finger bone 5), 0"))


#Exportiere Resultate in Datenblätter
setStyleAction(wb_3, XLC$STYLE_ACTION.NONE)
createSheet(wb_3, "IMPORT_Fledermaus_v1.xls")
createSheet(wb_3, "IMPORT_Fledermaeuse2_v1.xls")
writeWorksheet(wb_3, d.check1, "IMPORT_Fledermaus_v1.xls")
writeWorksheet(wb_3, d.check2, "IMPORT_Fledermaeuse2_v1.xls")
saveWorkbook(wb_3)
