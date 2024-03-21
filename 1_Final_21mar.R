#install.packages("readxl")
#install.packages("openxlsx")
library(readxl)
library(openxlsx)
getwd()

# jinu given maximo data
df1= read.xlsx("I:\\1_upwork\\004_work with Jinu_____data analysis\\demo_R_Copy\\mm.xlsx",
               startRow=5, colNames=T, detectDates=T, ) 
nrow(df1)+6 #1_4,5, second last 1
head(df1,1)
tail(df1[,1],1)

#df1 correction
df1[27] = as.Date( as.integer( unlist(df1[27])), origin = "1899-12-30") 
df1[34] [is.na(df1[34])] <-  df1[27] [is.na(df1[34])]                    

#df_final
df1_f= df1[1:(nrow(df1)-1), c(4,6:10,15:17,11,12,19:23,26,29:37)]    

df1_f[5]= as.numeric(unlist(df1_f[5]))
df1_f[10]= as.numeric(unlist(df1_f[10]))
df1_f[11]= as.numeric(unlist(df1_f[11]))

write.xlsx(df1_f,"I:\\1_upwork\\004_work with Jinu_____data analysis\\demo_R_Copy\\r_maximo_needed.xlsx",asTable = T,colNames=T)








