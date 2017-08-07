library(readxl)
library(tidyverse)
library(stringr)

x <- "a1~!@#$%^&*(){}_+:\"<>?,./;'[]-="


organizations <- read_xls("C:/Users/Jonathan/Downloads/MPI_Directory_by_Establishment_Number.xls") %>%
  mutate(X__1 = str_replace_all(tolower(X__1), "[[:punct:]]", ""))



osha1 <- read.csv("C:/Users/Jonathan/Downloads/osha_inspection_20170706.csv/osha_inspection-1.csv", stringsAsFactors = F) %>%
  #filter(sic_code == 2011) %>%
  mutate(estab_name = str_replace_all(tolower(estab_name), "[[:punct:]]", ""))
osha2 <-read.csv("C:/Users/Jonathan/Downloads/osha_inspection_20170706.csv/osha_inspection-2.csv", stringsAsFactors = F) %>%
   #           filter(sic_code == 2011) %>%
  mutate(estab_name = str_replace_all(tolower(estab_name), "[[:punct:]]", ""))
osha3 <-read.csv("C:/Users/Jonathan/Downloads/osha_inspection_20170706.csv/osha_inspection-3.csv", stringsAsFactors = F) %>%
  #filter(sic_code == 2011)%>%
  mutate(estab_name = str_replace_all(tolower(estab_name), "[[:punct:]]", ""))
osha4 <-read.csv("C:/Users/Jonathan/Downloads/osha_inspection_20170706.csv/osha_inspection-4.csv", stringsAsFactors = F) %>%
  #filter(sic_code == 2011)%>%
  mutate(estab_name = str_replace_all(tolower(estab_name), "[[:punct:]]", "")) 
osha5 <-read.csv("C:/Users/Jonathan/Downloads/osha_inspection_20170706.csv/osha_inspection-5.csv", stringsAsFactors = F) %>%
  #filter(sic_code == 2011)%>%
  mutate(estab_name = str_replace_all(tolower(estab_name), "[[:punct:]]", ""))

z <- rbind(osha1, osha2, osha3, osha4, osha5)

distance.methods<-c('osa','lv','dl','hamming','lcs','qgram','cosine','jaccard','jw')
dist.methods<-list()
for(m in 1:length(distance.methods))
{
  dist.name.enh<-matrix(NA, ncol = length(organizations$X__1),nrow = length(z$estab_name))
  for(i in 1:length(organizations$X__1)) {
    for(j in 1:length(z$estab_name)) { 
      dist.name.enh[j,i]<-stringdist(tolower(organizations[i,]$X__1),tolower(z[j,]$estab_name),method = distance.methods[m])      
      #adist.enhance(source2.devices[i,]$name,source1.devices[j,]$name)
    }  
  }
  dist.methods[[distance.methods[m]]]<-dist.name.enh
}

match.s1.s2.enh<-NULL
for(m in 1:length(dist.methods))
{
  
  dist.matrix<-as.matrix(dist.methods[[distance.methods[m]]])
  min.name.enh<-apply(dist.matrix, 1, base::min)
  for(i in 1:nrow(dist.matrix))
  {
    s2.i<-match(min.name.enh[i],dist.matrix[i,])
    s1.i<-i
    match.s1.s2.enh<-rbind(data.frame(s2.i=s2.i,s1.i=s1.i,s2name=source2.devices[s2.i,]$name, s1name=source1.devices[s1.i,]$name, adist=min.name.enh[i],method=distance.methods[m]),match.s1.s2.enh)
  }
}