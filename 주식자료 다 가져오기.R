library("readxl")   #install.packages("readxl")
library("qdap")   #install.packages("qdap")
library("qdap")   #install.packages("stringr")
library("stringr")
library("dplyr")
library("data.table")
library("zoo") #install.packages('zoo')
##########################################퀀트데이터1-재무자료

# Load a file
setwd('C:\\Users\\Administrator\\Desktop\\Dropbox\\주식 (석경하교수님)\\퀀트rawdata')
filename1<-"퀀트데이터1-재무자료.xlsx"

sheetnames<-excel_sheets(filename1) #시트이름 가져오기

data1<-list()

#데이터 정제 및 만들기 반복문
for(i in sheetnames){
  a<-read_excel(filename1,sheet = i)
  
  #Count rows and columns
  a_nrow<-nrow(a)
  a_ncol<-ncol(a)
  
  #Make a variable names
  timenames<-paste0(a[3,3:a_ncol] %>% str_sub(3,4),
                    mgsub(c("1Q","2Q","3Q","4Q"),c("03","06","09","12"),a[2,3:a_ncol]))
  
  id<-a[6:a_nrow,1]
  
  #Make values field
  values<-a[6:a_nrow,3:a_ncol]
  
  #Make a data frame
  data1[[i]]<-cbind(id,values)
  names(data1[[i]])<-c('id',timenames)
  data1[[i]] <- data1[[i]] %>% data.table()
}

##########################################퀀트데이터2-roe,roa,성장률

# Load a file
filename2<-"퀀트데이터2-roe,roa,성장률.xlsx"

sheetnames2<-excel_sheets(filename2) #시트이름 가져오기

data2<-list()

#데이터 정제 및 만들기 반복문
for(i in sheetnames2){
  a<-read_excel(filename2,sheet = i)
  
  #Count rows and columns
  a_nrow<-nrow(a)
  a_ncol<-ncol(a)
  
  #Make a variable names
  timenames<-paste0(a[3,3:a_ncol] %>% str_sub(3,4),
                    mgsub(c("1Q","2Q","3Q","4Q"),c("03","06","09","12"),a[2,3:a_ncol]))
  
  id<-a[6:a_nrow,1]
  
  #Make values field
  values<-a[6:a_nrow,3:a_ncol]
  
  #Make a data frame
  data2[[i]]<-cbind(id,values)
  names(data2[[i]])<-c('id',timenames)
  data2[[i]] <- data2[[i]] %>% data.table()
}

##########################################퀀트데이터3-4분기누적실적

# Load a file
filename3<-"퀀트데이터3-4분기누적실적.xlsx"

sheetnames3<-excel_sheets(filename3) #시트이름 가져오기

data3<-list()

#데이터 정제 및 만들기 반복문
for(i in sheetnames3){
  a<-read_excel(filename3,sheet = i)
  
  #Count rows and columns
  a_nrow<-nrow(a)
  a_ncol<-ncol(a)
  
  #Make a variable names
  timenames<-paste0(a[3,3:a_ncol] %>% str_sub(3,4),
                    mgsub(c("1Q","2Q","3Q","4Q"),c("03","06","09","12"),a[2,3:a_ncol]))
  
  id<-a[6:a_nrow,1]
  
  #Make values field
  values<-a[6:a_nrow,3:a_ncol]
  
  #Make a data frame
  data3[[i]]<-cbind(id,values)
  names(data3[[i]])<-c('id',timenames)
  data3[[i]] <- data3[[i]] %>% data.table()
}

##########################################퀀트데이터4-시총,수정주가

# Load a file
filename4<-"퀀트데이터4-시총,수정주가.xlsx"

sheetnames4<-excel_sheets(filename4) #시트이름 가져오기

data4<-list()

#데이터 정제 및 만들기 반복문
for(i in sheetnames4){
  a<-read_excel(filename4,sheet = i)
  
  #Count rows and columns
  a_nrow<-nrow(a)
  a_ncol<-ncol(a)
  
  #Make a variable names
  
  id<-a[8,2:a_ncol] %>% unlist()
  
  timenames<-a[14:a_nrow,1] %>% unlist() %>% as.integer()-365*70-19
  timenames_1<-as.Date(timenames) %>% as.character() %>% substr(3,4)
  timenames_2<-as.Date(timenames) %>% as.character() %>% substr(6,7)
  timenames<-paste0(timenames_1,timenames_2)
  
  #Make values field
  vf<-a[14:a_nrow,2:a_ncol] %>% t() %>% as.data.frame() %>%data.frame(row.names = NULL)
  
  #Make a data frame
  data4[[i]]<-cbind(id,vf)
  names(data4[[i]])<-c('id',timenames)
  data4[[i]] <- data4[[i]] %>% data.table()
}

View(data1[[1]]) #1~21
View(data2[[1]]) #1~9
View(data3[[1]]) #1~3
View(data4[[1]]) #1~2