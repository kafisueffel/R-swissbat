#First Test
setwd("/Users/christoph_user/Dropbox/SwissBat/Z?rich/Swissbat/5_separate_Einzeldaten")
setwd("C:/Users/kkr/Dropbox_CT/Dropbox/SwissBat/Z?rich/Swissbat/5_separate_Einzeldaten")
require(XLConnect)
wb_1 <- loadWorkbook("IMPORT_Fledermaus_v1.xls")
d.d_1 <- readWorksheet(wb_1, 1)
head(d.d_1)

#Extrahiere Ortsinformationen
l.loc <- strsplit(d.d_1$Verbatim.Locality, ", ")
d.loc <- as.data.frame(do.call(rbind, l.loc))
names(d.loc) <- c("ortname", "adresszusatz", "Kanton", "Land")

#Extrahiere Morphologie
d.d_1$Col.Obj.Attribut.Remarks <- gsub("/ ", "/", d.d_1$Col.Obj.Attribut.Remarks)

l.morph <- strsplit(d.d_1$Col.Obj.Attribut.Remarks, "/")
d.morph <- as.data.frame(do.call(rbind, l.morph))
names(d.morph) <- c("Unterarm", "Daumen", "Strahl3", "Strahl5")
head(d.morph)



d.morph <- within(d.morph, {
  Unterarm <- gsub("FOREA \\(Unterarmlänge\\), ", "", Unterarm)
  Daumen <- gsub("Digi1 \\(finger bone 1\\), ", "", Daumen)
  Strahl3 <- gsub("Digi1 \\(finger bone 3\\), ", "", Strahl3)
  Strahl5 <- gsub("Digi1 \\(finger bone 5\\), ", "", Strahl5)
})


d.morph <- within(d.morph, {
  Unterarm <- as.numeric(Unterarm)
  Daumen <- as.numeric(Daumen)
  Strahl3 <- as.numeric(Strahl3)
  Strahl5 <- as.numeric(Strahl5)
})