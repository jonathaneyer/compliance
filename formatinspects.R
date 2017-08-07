excellist <- list.files(pattern = ".xlsx") 
excellist <- excellist[substr(excellist,1,1) != "~"]

for(j in excellist){
  d <- read_excel(j)
  assign(j,d)
}







derp <- function(dat){
  names(dat)[1] <- "header"
  
  tabnames <- dat$header[grepl("Table",dat$header)]
  
  proh_s <- which(grepl("Prohibited Activity Notices Issued by Establishment",dat$header))
  proh_e <- ifelse(which(grepl("Prohibited Activity Notices Issued by Establishment",tabnames))== length(tabnames),length(dat$header),which(grepl(substr(tabnames[(which(grepl("Prohibited Activity Notices Issued by Establishment",tabnames))+1)],1,10),dat$header)))
  prohibit <- dat[(proh_s+1):(proh_e-1),]
  prohibit <- Filter(function(x)!all(is.na(x)), prohibit)
  prohibit <- subset(prohibit, !(grepl("Establishment/Firm", prohibit$header)))
  prohibit <- subset(prohibit, is.na(as.numeric(prohibit$header)))
  
  largeadmin_s <- which(grepl("Administrative Actions: Large Establishments",dat$header))
  largeadmin_e <- ifelse(which(grepl("Administrative Actions: Large Establishments",tabnames))== length(tabnames),length(dat$header),which(grepl(substr(tabnames[(which(grepl("Administrative Actions: Large Establishments",tabnames))+1)],1,10),dat$header)))
  largeadmin <- dat[(largeadmin_s+1):(largeadmin_e-1),]
  names(largeadmin) <- paste(largeadmin[2,],largeadmin[3,], sep = "_")
  largeadmin <- largeadmin[,names(largeadmin) != "NA_NA"]
  names(largeadmin)[1] <- "header"
  largeadmin <- Filter(function(x)!all(is.na(x)), largeadmin)
  largeadmin <- subset(largeadmin, !(grepl("Establishment/", largeadmin$header)))
  largeadmin <- subset(largeadmin, !(grepl("Administrative Actions", largeadmin$header)))
  largeadmin <- subset(largeadmin, is.na(as.numeric(largeadmin$header)))
  largeadmin <- subset(largeadmin, !rowSums(largeadmin == "SSOP",na.rm = T))
  largeadmin <- subset(largeadmin, !rowSums(largeadmin == "ABBR SPS",na.rm = T))
  largeadmin <- subset(largeadmin, !rowSums(largeadmin == "INH",na.rm = T))
  largeadmin <- subset(largeadmin, !rowSums(largeadmin == "INT",na.rm = T))
  largeadmin <- subset(largeadmin, !rowSums(largeadmin == "LOI",na.rm = T))
  largeadmin <- subset(largeadmin, !rowSums(largeadmin == "LOW",na.rm = T))
  largeadmin <- subset(largeadmin, !grepl(pattern = "ABBR",x = largeadmin$header))
  
  smalladmin_s <- which(grepl("Administrative Actions: Small Establishments",dat$header))
  smalladmin_e <-ifelse(which(grepl("Administrative Actions: Small Establishments",tabnames))== length(tabnames),length(dat$header),which(grepl(substr(tabnames[(which(grepl("Administrative Actions: Small Establishments",tabnames))+1)],1,10),dat$header)))
  smalladmin <- dat[(smalladmin_s+1):(smalladmin_e-1),]
  names(smalladmin) <- paste(smalladmin[2,],smalladmin[3,], sep = "_")
  smalladmin <- smalladmin[,names(smalladmin) != "NA_NA"]
  names(smalladmin)[1] <- "header"
  smalladmin <- Filter(function(x)!all(is.na(x)), smalladmin)
  smalladmin <- subset(smalladmin, !(grepl("Establishment/", smalladmin$header)))
  smalladmin <- subset(smalladmin, !(grepl("Administrative Actions", smalladmin$header)))
  smalladmin <- subset(smalladmin, is.na(as.numeric(smalladmin$header)))
  smalladmin <- subset(smalladmin, !rowSums(smalladmin == "SSOP",na.rm = T))
  smalladmin <- subset(smalladmin, !rowSums(smalladmin == "SSOP",na.rm = T))
  smalladmin <- subset(smalladmin, !rowSums(smalladmin == "ABBR SPS",na.rm = T))
  smalladmin <- subset(smalladmin, !rowSums(smalladmin == "INH",na.rm = T))
  smalladmin <- subset(smalladmin, !rowSums(smalladmin == "INT",na.rm = T))
  smalladmin <- subset(smalladmin, !rowSums(smalladmin == "LOI",na.rm = T))
  smalladmin <- subset(smalladmin, !rowSums(smalladmin == "LOW",na.rm = T))
  
  vsmalladmin_s <- which(grepl("Administrative Actions: Very Small Establishments",dat$header))
  vsmalladmin_e <- ifelse(which(grepl("Administrative Actions: Very Small Establishments",tabnames))== length(tabnames),length(dat$header),which(grepl(substr(tabnames[(which(grepl("Administrative Actions: Very Small Establishments",tabnames))+1)],1,10),dat$header)))
  vsmalladmin <- dat[(vsmalladmin_s+1):(vsmalladmin_e-1),]
  names(vsmalladmin) <- paste(vsmalladmin[2,],vsmalladmin[3,], sep = "_")
  vsmalladmin <- vsmalladmin[,names(vsmalladmin) != "NA_NA"]
  names(vsmalladmin)[1] <- "header"
  vsmalladmin <- Filter(function(x)!all(is.na(x)), vsmalladmin)
  vsmalladmin <- subset(vsmalladmin, !(grepl("Establishment/", vsmalladmin$header)))
  vsmalladmin <- subset(vsmalladmin, !(grepl("Administrative Actions", vsmalladmin$header)))
  vsmalladmin <- subset(vsmalladmin, is.na(as.numeric(vsmalladmin$header)))
  vsmalladmin <- subset(vsmalladmin, !rowSums(vsmalladmin == "SSOP",na.rm = T))
  vsmalladmin <- subset(vsmalladmin, !rowSums(vsmalladmin == "SSOP",na.rm = T))
  vsmalladmin <- subset(vsmalladmin, !rowSums(vsmalladmin == "ABBR SPS",na.rm = T))
  vsmalladmin <- subset(vsmalladmin, !rowSums(vsmalladmin == "INH",na.rm = T))
  vsmalladmin <- subset(vsmalladmin, !rowSums(vsmalladmin == "INT",na.rm = T))
  vsmalladmin <- subset(vsmalladmin, !rowSums(vsmalladmin == "LOI",na.rm = T))
  vsmalladmin <- subset(vsmalladmin, !rowSums(vsmalladmin == "LOW",na.rm = T))
  
  warn_s <- which(grepl("Notices of Warnings Issued",dat$header))
  warn_e <-  ifelse(which(grepl("Notices of Warnings Issued",tabnames))== length(tabnames),length(dat$header),which(grepl(substr(tabnames[(which(grepl("Notices of Warnings Issued",tabnames))+1)],1,10),dat$header)))
  warnings <- dat[(warn_s+1):(warn_e-1),] 
  warnings <- subset(warnings, !(grepl("Firm/Establishment", warnings$header)))
  warnings <- subset(warnings, !(grepl("Inquiries", warnings$header)))
  warnings <- subset(warnings, !(grepl("INFORMATION", warnings$header)))
  warnings <- subset(warnings, is.na(as.numeric(warnings$header)))
  warnings <- Filter(function(x)!all(is.na(x)), warnings)
  
  return(list(prohibit, largeadmin, smalladmin,vsmalladmin, warnings))
  
}


derp15 <- function(dat){
  names(dat)[1] <- "header"
  
  tabnames <- dat$header[grepl("Table",dat$header)]
  
  proh_s <- which(grepl("Prohibited Activity Notices Issued by Establishment",dat$header))
  proh_e <- ifelse(which(grepl("Prohibited Activity Notices Issued by Establishment",tabnames))== length(tabnames),length(dat$header),which(grepl(substr(tabnames[(which(grepl("Prohibited Activity Notices Issued by Establishment",tabnames))+1)],1,10),dat$header)))
  prohibit <- dat[(proh_s+1):(proh_e-1),]
  prohibit <- Filter(function(x)!all(is.na(x)), prohibit)
  prohibit <- Filter(function(x)!all(is.na(x)), prohibit)
  prohibit <- subset(prohibit, !(grepl("Establishment/Firm", prohibit$header)))
  prohibit <- subset(prohibit, is.na(as.numeric(prohibit$header)))
  
  largeadmin_s <- which(grepl("Administrative Actions: Large Establishments",dat$header))
  largeadmin_e <- ifelse(which(grepl("Administrative Actions: Large Establishments",tabnames))== length(tabnames),length(dat$header),which(grepl(substr(tabnames[(which(grepl("Administrative Actions: Large Establishments",tabnames))+1)],1,10),dat$header)))
  largeadmin <- dat[(largeadmin_s+1):(largeadmin_e-1),]
  names(largeadmin) <- paste(largeadmin[2,],largeadmin[3,], sep = "_")
  largeadmin <- largeadmin[,names(largeadmin) != "NA_NA"]
  names(largeadmin)[1] <- "header"
  largeadmin <- Filter(function(x)!all(is.na(x)), largeadmin)
  largeadmin <- subset(largeadmin, !(grepl("Establishment/", largeadmin$header)))
  largeadmin <- subset(largeadmin, !(grepl("Administrative Actions", largeadmin$header)))
  largeadmin <- subset(largeadmin, is.na(as.numeric(largeadmin$header)))
  largeadmin <- subset(largeadmin, !rowSums(largeadmin == "SSOP",na.rm = T))
  largeadmin <- subset(largeadmin, !rowSums(largeadmin == "ABBR SPS",na.rm = T))
  largeadmin <- subset(largeadmin, !rowSums(largeadmin == "INH",na.rm = T))
  largeadmin <- subset(largeadmin, !rowSums(largeadmin == "INT",na.rm = T))
  largeadmin <- subset(largeadmin, !rowSums(largeadmin == "LOI",na.rm = T))
  largeadmin <- subset(largeadmin, !rowSums(largeadmin == "LOW",na.rm = T))
  
  smalladmin_s <- which(grepl("Administrative Actions: Small Establishments",dat$header))
  smalladmin_e <-ifelse(which(grepl("Administrative Actions: Small Establishments",tabnames))== length(tabnames),length(dat$header),which(grepl(substr(tabnames[(which(grepl("Administrative Actions: Small Establishments",tabnames))+1)],1,10),dat$header)))
  smalladmin <- dat[(smalladmin_s+1):(smalladmin_e-1),]
  names(smalladmin) <- paste(smalladmin[2,],smalladmin[3,], sep = "_")
  smalladmin <- smalladmin[,names(smalladmin) != "NA_NA"]
  names(smalladmin)[1] <- "header"
  smalladmin <- Filter(function(x)!all(is.na(x)), smalladmin)
  smalladmin <- subset(smalladmin, !(grepl("Establishment/", smalladmin$header)))
  smalladmin <- subset(smalladmin, !(grepl("Administrative Actions", smalladmin$header)))
  smalladmin <- subset(smalladmin, is.na(as.numeric(smalladmin$header)))
  smalladmin <- subset(smalladmin, !rowSums(smalladmin == "SSOP",na.rm = T))
  smalladmin <- subset(smalladmin, !rowSums(smalladmin == "SSOP",na.rm = T))
  smalladmin <- subset(smalladmin, !rowSums(smalladmin == "ABBR SPS",na.rm = T))
  smalladmin <- subset(smalladmin, !rowSums(smalladmin == "INH",na.rm = T))
  smalladmin <- subset(smalladmin, !rowSums(smalladmin == "INT",na.rm = T))
  smalladmin <- subset(smalladmin, !rowSums(smalladmin == "LOI",na.rm = T))
  smalladmin <- subset(smalladmin, !rowSums(smalladmin == "LOW",na.rm = T))
  
  vsmalladmin_s <- which(grepl("Administrative Actions: Very Small Establishments",dat$header))
  vsmalladmin_e <- ifelse(which(grepl("Administrative Actions: Very Small Establishments",tabnames))== length(tabnames),length(dat$header),which(grepl(substr(tabnames[(which(grepl("Administrative Actions: Very Small Establishments",tabnames))+1)],1,10),dat$header)))
  vsmalladmin <- dat[(vsmalladmin_s+1):(vsmalladmin_e-1),]
  names(vsmalladmin) <- paste(vsmalladmin[2,],vsmalladmin[3,], sep = "_")
  vsmalladmin <- vsmalladmin[,names(vsmalladmin) != "NA_NA"]
  names(vsmalladmin)[1] <- "header"
  vsmalladmin <- Filter(function(x)!all(is.na(x)), vsmalladmin)
  vsmalladmin <- subset(vsmalladmin, !(grepl("Establishment/", vsmalladmin$header)))
  vsmalladmin <- subset(vsmalladmin, !(grepl("Administrative Actions", vsmalladmin$header)))
  vsmalladmin <- subset(vsmalladmin, is.na(as.numeric(vsmalladmin$header)))
  vsmalladmin <- subset(vsmalladmin, !rowSums(vsmalladmin == "SSOP",na.rm = T))
  vsmalladmin <- subset(vsmalladmin, !rowSums(vsmalladmin == "SSOP",na.rm = T))
  vsmalladmin <- subset(vsmalladmin, !rowSums(vsmalladmin == "ABBR SPS",na.rm = T))
  vsmalladmin <- subset(vsmalladmin, !rowSums(vsmalladmin == "INH",na.rm = T))
  vsmalladmin <- subset(vsmalladmin, !rowSums(vsmalladmin == "INT",na.rm = T))
  vsmalladmin <- subset(vsmalladmin, !rowSums(vsmalladmin == "LOI",na.rm = T))
  vsmalladmin <- subset(vsmalladmin, !rowSums(vsmalladmin == "LOW",na.rm = T))
  
  warn_s <- which(grepl("Notices of Warning Issued by OIEA by Firm",dat$header))
  warn_e <-  ifelse(which(grepl("Notices of Warning Issued by OIEA by Firm",tabnames))== length(tabnames),length(dat$header),which(grepl(substr(tabnames[(which(grepl("Notices of Warning Issued by OIEA by Firm",tabnames))+1)],1,10),dat$header)))
  warnings <- dat[(warn_s+1):(warn_e-1),] 
  names(warnings) <- paste(warnings[2,],warnings[3,], sep = "_")
  warnings <- warnings[,names(warnings) != "NA_NA"]
  names(warnings)[1] <- "header"
  warnings <- subset(warnings, !(grepl("Firm/Establishment", largeadmin$header)))
  warnings <- subset(warnings, !(grepl("Inquiries", warnings$header)))
  warnings <- subset(warnings, !(grepl("INFORMATION", warnings$header)))
  warnings <- subset(warnings, is.na(as.numeric(largeadmin$header)))
  warnings <- subset(warnings, is.na(as.numeric(largeadmin$header)))
  warnings <- subset(warnings, is.na(as.numeric(largeadmin$header)))
  warnings <- Filter(function(x)!all(is.na(x)), warnings)
  
  
  return(list(prohibit, largeadmin, smalladmin,vsmalladmin, warnings))
  
}

fy11_q1_dat <- derp(dat = fy11_q1.xlsx)
fy11_q2_dat <- derp(dat = fy11_q2.xlsx)
fy11_q3_dat <- derp(dat = fy11_q3.xlsx)
fy11_q4_dat <- derp(dat = fy11_q4.xlsx)
fy12_q1_dat <- derp(dat = fy12_q1.xlsx)
fy12_q2_dat <- derp(dat = fy12_q2.xlsx)
fy12_q3_dat <- derp(dat = fy12_q3.xlsx)
fy12_q4_dat <- derp(dat = fy12_q4.xlsx)
fy13_q1_dat <- derp(dat = fy13_q1.xlsx)
fy13_q2_dat <- derp(dat = fy13_q2.xlsx)
fy13_q3_dat <- derp(dat = fy13_q3.xlsx)
fy13_q4_dat <- derp(dat = fy13_q4.xlsx)
fy14_q1_dat <- derp(dat = fy14_q1.xlsx)
fy14_q2_dat <- derp(dat = fy14_q2.xlsx)
fy14_q3_dat <- derp(dat = fy14_q3.xlsx)
fy14_q4_dat <- derp(dat = fy14_q4.xlsx)
fy15_q1_dat <- derp(dat = fy15_q1.xlsx)
fy15_q2_dat <- derp(dat = fy15_q2.xlsx)
fy15_q3_dat <- derp15(dat = fy15_q3.xlsx)
fy15_q4_dat <- derp15(dat = fy15_q4.xlsx)
fy16_q1_dat <- derp15(dat = fy16_q1.xlsx)
fy16_q2_dat <- derp15(dat = fy16_q2.xlsx)
fy16_q3_dat <- derp15(dat = fy16_q3.xlsx)
fy16_q4_dat <- derp15(dat = fy16_q4.xlsx)

library(XLConnect)

writeWorksheetToFile(data = fy11_q1_dat[1], file = paste("formatted/fy11_q1.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy11_q1_dat[2],  file = paste("formatted/fy11_q1.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy11_q1_dat[3], file  = paste("formatted/fy11_q1.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy11_q1_dat[4], file  = paste("formatted/fy11_q1.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy11_q1_dat[5], file  = paste("formatted/fy11_q1.xlsx", sep = ""),sheet = "warnings")  

writeWorksheetToFile(data = fy11_q2_dat[1], file = paste("formatted/fy11_q2.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy11_q2_dat[2],  file = paste("formatted/fy11_q2.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy11_q2_dat[3], file  = paste("formatted/fy11_q2.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy11_q2_dat[4], file  = paste("formatted/fy11_q2.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy11_q2_dat[5], file  = paste("formatted/fy11_q2.xlsx", sep = ""),sheet = "warnings")  

writeWorksheetToFile(data = fy11_q3_dat[1], file = paste("formatted/fy11_q3.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy11_q3_dat[2],  file = paste("formatted/fy11_q3.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy11_q3_dat[3], file  = paste("formatted/fy11_q3.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy11_q3_dat[4], file  = paste("formatted/fy11_q3.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy11_q3_dat[5], file  = paste("formatted/fy11_q3.xlsx", sep = ""),sheet = "warnings")  

writeWorksheetToFile(data = fy11_q4_dat[1], file = paste("formatted/fy11_q4.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy11_q4_dat[2],  file = paste("formatted/fy11_q4.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy11_q4_dat[3], file  = paste("formatted/fy11_q4.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy11_q4_dat[4], file  = paste("formatted/fy11_q4.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy11_q4_dat[5], file  = paste("formatted/fy11_q4.xlsx", sep = ""),sheet = "warnings")  

writeWorksheetToFile(data = fy12_q1_dat[1], file = paste("formatted/fy12_q1.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy12_q1_dat[2],  file = paste("formatted/fy12_q1.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy12_q1_dat[3], file  = paste("formatted/fy12_q1.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy12_q1_dat[4], file  = paste("formatted/fy12_q1.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy12_q1_dat[5], file  = paste("formatted/fy12_q1.xlsx", sep = ""),sheet = "warnings")  

writeWorksheetToFile(data = fy12_q2_dat[1], file = paste("formatted/fy12_q2.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy12_q2_dat[2],  file = paste("formatted/fy12_q2.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy12_q2_dat[3], file  = paste("formatted/fy12_q2.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy12_q2_dat[4], file  = paste("formatted/fy12_q2.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy12_q2_dat[5], file  = paste("formatted/fy12_q2.xlsx", sep = ""),sheet = "warnings")  

writeWorksheetToFile(data = fy12_q3_dat[1], file = paste("formatted/fy12_q3.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy12_q3_dat[2],  file = paste("formatted/fy12_q3.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy12_q3_dat[3], file  = paste("formatted/fy12_q3.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy12_q3_dat[4], file  = paste("formatted/fy12_q3.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy12_q3_dat[5], file  = paste("formatted/fy12_q3.xlsx", sep = ""),sheet = "warnings")  

writeWorksheetToFile(data = fy12_q4_dat[1], file = paste("formatted/fy12_q4.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy12_q4_dat[2],  file = paste("formatted/fy12_q4.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy12_q4_dat[3], file  = paste("formatted/fy12_q4.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy12_q4_dat[4], file  = paste("formatted/fy12_q4.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy12_q4_dat[5], file  = paste("formatted/fy12_q4.xlsx", sep = ""),sheet = "warnings")  

writeWorksheetToFile(data = fy13_q1_dat[1], file = paste("formatted/fy13_q1.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy13_q1_dat[2],  file = paste("formatted/fy13_q1.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy13_q1_dat[3], file  = paste("formatted/fy13_q1.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy13_q1_dat[4], file  = paste("formatted/fy13_q1.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy13_q1_dat[5], file  = paste("formatted/fy13_q1.xlsx", sep = ""),sheet = "warnings")  

writeWorksheetToFile(data = fy13_q2_dat[1], file = paste("formatted/fy13_q2.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy13_q2_dat[2],  file = paste("formatted/fy13_q2.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy13_q2_dat[3], file  = paste("formatted/fy13_q2.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy13_q2_dat[4], file  = paste("formatted/fy13_q2.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy13_q2_dat[5], file  = paste("formatted/fy13_q2.xlsx", sep = ""),sheet = "warnings")  

writeWorksheetToFile(data = fy13_q3_dat[1], file = paste("formatted/fy13_q3.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy13_q3_dat[2],  file = paste("formatted/fy13_q3.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy13_q3_dat[3], file  = paste("formatted/fy13_q3.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy13_q3_dat[4], file  = paste("formatted/fy13_q3.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy13_q3_dat[5], file  = paste("formatted/fy13_q3.xlsx", sep = ""),sheet = "warnings")  

writeWorksheetToFile(data = fy13_q4_dat[1], file = paste("formatted/fy13_q4.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy13_q4_dat[2],  file = paste("formatted/fy13_q4.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy13_q4_dat[3], file  = paste("formatted/fy13_q4.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy13_q4_dat[4], file  = paste("formatted/fy13_q4.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy13_q4_dat[5], file  = paste("formatted/fy13_q4.xlsx", sep = ""),sheet = "warnings")  

writeWorksheetToFile(data = fy14_q1_dat[1], file = paste("formatted/fy14_q1.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy14_q1_dat[2],  file = paste("formatted/fy14_q1.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy14_q1_dat[3], file  = paste("formatted/fy14_q1.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy14_q1_dat[4], file  = paste("formatted/fy14_q1.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy14_q1_dat[5], file  = paste("formatted/fy14_q1.xlsx", sep = ""),sheet = "warnings")  

writeWorksheetToFile(data = fy14_q2_dat[1], file = paste("formatted/fy14_q2.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy14_q2_dat[2],  file = paste("formatted/fy14_q2.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy14_q2_dat[3], file  = paste("formatted/fy14_q2.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy14_q2_dat[4], file  = paste("formatted/fy14_q2.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy14_q2_dat[5], file  = paste("formatted/fy14_q2.xlsx", sep = ""),sheet = "warnings")  

writeWorksheetToFile(data = fy14_q3_dat[1], file = paste("formatted/fy14_q3.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy14_q3_dat[2],  file = paste("formatted/fy14_q3.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy14_q3_dat[3], file  = paste("formatted/fy14_q3.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy14_q3_dat[4], file  = paste("formatted/fy14_q3.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy14_q3_dat[5], file  = paste("formatted/fy14_q3.xlsx", sep = ""),sheet = "warnings")  

writeWorksheetToFile(data = fy14_q4_dat[1], file = paste("formatted/fy14_q4.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy14_q4_dat[2],  file = paste("formatted/fy14_q4.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy14_q4_dat[3], file  = paste("formatted/fy14_q4.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy14_q4_dat[4], file  = paste("formatted/fy14_q4.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy14_q4_dat[5], file  = paste("formatted/fy14_q4.xlsx", sep = ""),sheet = "warnings")  

writeWorksheetToFile(data = fy15_q1_dat[1], file = paste("formatted/fy15_q1.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy15_q1_dat[2],  file = paste("formatted/fy15_q1.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy15_q1_dat[3], file  = paste("formatted/fy15_q1.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy15_q1_dat[4], file  = paste("formatted/fy15_q1.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy15_q1_dat[5], file  = paste("formatted/fy15_q1.xlsx", sep = ""),sheet = "warnings")  

writeWorksheetToFile(data = fy15_q2_dat[1], file = paste("formatted/fy15_q2.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy15_q2_dat[2],  file = paste("formatted/fy15_q2.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy15_q2_dat[3], file  = paste("formatted/fy15_q2.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy15_q2_dat[4], file  = paste("formatted/fy15_q2.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy15_q2_dat[5], file  = paste("formatted/fy15_q2.xlsx", sep = ""),sheet = "warnings")  

writeWorksheetToFile(data = fy15_q3_dat[1], file = paste("formatted/fy15_q3.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy15_q3_dat[2],  file = paste("formatted/fy15_q3.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy15_q3_dat[3], file  = paste("formatted/fy15_q3.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy15_q3_dat[4], file  = paste("formatted/fy15_q3.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy15_q3_dat[5], file  = paste("formatted/fy15_q3.xlsx", sep = ""),sheet = "warnings")  

writeWorksheetToFile(data = fy15_q4_dat[1], file = paste("formatted/fy15_q4.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy15_q4_dat[2],  file = paste("formatted/fy15_q4.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy15_q4_dat[3], file  = paste("formatted/fy15_q4.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy15_q4_dat[4], file  = paste("formatted/fy15_q4.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy15_q4_dat[5], file  = paste("formatted/fy15_q4.xlsx", sep = ""),sheet = "warnings")  

writeWorksheetToFile(data  = fy16_q1_dat[1], file = paste("formatted/fy16_q1.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy16_q1_dat[2],  file = paste("formatted/fy16_q1.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy16_q1_dat[3], file  = paste("formatted/fy16_q1.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy16_q1_dat[4], file  = paste("formatted/fy16_q1.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy16_q1_dat[5], file  = paste("formatted/fy16_q1.xlsx", sep = ""),sheet = "warnings")  

writeWorksheetToFile(data  = fy16_q2_dat[1], file = paste("formatted/fy16_q2.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy16_q2_dat[2],  file = paste("formatted/fy16_q2.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy16_q2_dat[3], file  = paste("formatted/fy16_q2.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy16_q2_dat[4], file  = paste("formatted/fy16_q2.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy16_q2_dat[5], file  = paste("formatted/fy16_q2.xlsx", sep = ""),sheet = "warnings")  

writeWorksheetToFile(data  = fy16_q3_dat[1], file = paste("formatted/fy16_q3.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy16_q3_dat[2],  file = paste("formatted/fy16_q3.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy16_q3_dat[3], file  = paste("formatted/fy16_q3.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy16_q3_dat[4], file  = paste("formatted/fy16_q3.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy16_q3_dat[5], file  = paste("formatted/fy16_q3.xlsx", sep = ""),sheet = "warnings")  

writeWorksheetToFile(data  = fy16_q4_dat[1], file = paste("formatted/fy16_q4.xlsx", sep =""),sheet = "prohibited")  
writeWorksheetToFile(data = fy16_q4_dat[2],  file = paste("formatted/fy16_q4.xlsx", sep = ""),sheet = "large")  
writeWorksheetToFile(data = fy16_q4_dat[3], file  = paste("formatted/fy16_q4.xlsx", sep = ""),sheet = "small")  
writeWorksheetToFile(data = fy16_q4_dat[4], file  = paste("formatted/fy16_q4.xlsx", sep = ""),sheet = "vsmall")
writeWorksheetToFile(data = fy16_q4_dat[5], file  = paste("formatted/fy16_q4.xlsx", sep = ""),sheet = "warnings")  
