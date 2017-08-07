library(readr)
library(readxl)

excellist <- list.files(path = "usdainspects/formatted/" ,pattern = "*.xlsx") 
excellist <- excellist[substr(excellist,1,1) != "~"]

for(j in excellist){
  d <- read_excel(paste("usdainspects/formatted/",j, sep = ""),sheet = 1)  
  names(d) <- paste("V",seq(1,ncol(d)),sep = "")
  if(exists("prohibited")){
    prohibited <- rbind(prohibited, as.data.frame(d))
  }
  if(!exists("prohibited")){
    prohibited <- as.data.frame(d)
  }
  
  l <- read_excel(paste("usdainspects/formatted/",j, sep = ""),sheet = 2)  
  names(l) <- paste("V",seq(1,ncol(l)),sep = "")
  if(exists("large")){
    large <- rbind(large, as.data.frame(l))
  }
  if(!exists("large")){
    large <- as.data.frame(l)
  }
  
  s <- read_excel(paste("usdainspects/formatted/",j, sep = ""),sheet = 3)  
  names(s) <- paste("V",seq(1,ncol(s)),sep = "")
  if(exists("small")){
    small <- rbind(small, as.data.frame(s))
  }
  if(!exists("small")){
    small <- as.data.frame(s)
  }
  
  vs <- read_excel(paste("usdainspects/formatted/",j, sep = ""),sheet = 4)  
  names(vs) <- paste("V",seq(1,ncol(vs)),sep = "")
  if(exists("vsmall")){
    vsmall <- rbind(vsmall, as.data.frame(vs))
  }
  if(!exists("vsmall")){
    vsmall <- as.data.frame(vs)
  }
  
  
  w <- read_excel(paste("usdainspects/formatted/",j, sep = ""),sheet = 5)  
  names(w) <- paste("V",seq(1,ncol(w)),sep = "")
  if(exists("warn")){
    warn <- rbind(warn, as.data.frame(w))
  }
  if(!exists("warn")){
    warn <- as.data.frame(w)
  }
  
}

prohibited <- na.omit(prohibited)
warn <- na.omit(warn)

for(q in 1:nrow(large)){
  if(is.na(large$V1[q])){large$V1[q] <- large$V1[(q-1)]
  }
}

for(q in 1:nrow(small)){
  if(is.na(small$V1[q])){small$V1[q] <- small$V1[(q-1)]
  }
}

for(q in 1:nrow(vsmall)){
  if(is.na(vsmall$V1[q])){vsmall$V1[q] <- vsmall$V1[(q-1)]
  }
}

prohibited$date <- as.Date(as.numeric(prohibited$V3),origin = "1899-12-30")
warn$date <- as.Date(as.numeric(warn$V4),origin = "1899-12-30")
small$noiadate <- as.Date(as.numeric(small$V2),origin = "1899-12-30") 
small$deferdate <- as.Date(as.numeric(small$V3),origin = "1899-12-30") 
small$suspstartdate <- as.Date(as.numeric(small$V4),origin = "1899-12-30") 
small$suspenddate <- as.Date(as.numeric(small$V5),origin = "1899-12-30") 
vsmall$noiadate <- as.Date(as.numeric(vsmall$V2),origin = "1899-12-30") 
vsmall$deferdate <- as.Date(as.numeric(vsmall$V3),origin = "1899-12-30") 
vsmall$suspstartdate <- as.Date(as.numeric(vsmall$V4),origin = "1899-12-30") 
vsmall$suspenddate <- as.Date(as.numeric(vsmall$V5),origin = "1899-12-30") 
large$noiadate <- as.Date(as.numeric(large$V2),origin = "1899-12-30") 
large$deferdate <- as.Date(as.numeric(large$V3),origin = "1899-12-30") 
large$suspstartdate <- as.Date(as.numeric(large$V4),origin = "1899-12-30") 
large$suspenddate <- as.Date(as.numeric(large$V5),origin = "1899-12-30")

for(j in 1:nrow(small)){
dat <- strsplit(small$V1[j], split = "(?<=[a-zA-Z])\\s*(?=[0-9])", perl = TRUE)
small$name[j] <- unlist(dat)[1]
small$city[j] <- unlist(strsplit(unlist(strsplit(unlist(dat)[length(unlist(dat))],"\n"))[2],","))[1]
small$state[j] <- unlist(strsplit(unlist(strsplit(unlist(dat)[length(unlist(dat))],"\n"))[2],","))[2]
}

