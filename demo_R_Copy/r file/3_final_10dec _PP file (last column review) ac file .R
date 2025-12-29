#install.packages("readxl")
#install.packages("openxlsx")
library(readxl)
library(openxlsx)
library(dplyr)
getwd()


##########################################################################################
#for wo

wo1= read.xlsx("I:\\1_upwork\\004_work with Jinu_____data analysis\\demo_R_Copy\\wo.xlsx",
               sheet=1,colNames=T, detectDates=F) 
wo2= read.xlsx("I:\\1_upwork\\004_work with Jinu_____data analysis\\demo_R_Copy\\wo.xlsx",
               sheet=2,colNames=T, detectDates=F) 

wo1=wo1[,c(1:4,20)]
#wo1=wo1[,c(18,19,9,10,22)]
head(wo1[5],1)
wo2=wo2[,c(1:4,21)] #15
head(wo2[5],1)
wo=rbind(wo1,wo2)

wo[1]=trimws(unlist(wo[1]))  
wo[2]=trimws(unlist(wo[2]))  
wo[3]=trimws(unlist(wo[3]))  
wo[4]=trimws(unlist(wo[4]))  
wo[5]=trimws(unlist(wo[5])) 

pr_ii<- paste(wo[[3]], wo[[4]],sep="_")
wo <- data.frame(pr_it= pr_ii,wo)
wo_1 <- wo[!duplicated(wo[1],fromLast= T), ] 

#########################################################################################
pp= read.xlsx("I:\\1_upwork\\004_work with Jinu_____data analysis\\demo_R_Copy\\pp.xlsx",
               sheet=3,colNames=T, detectDates=F) 

pp=pp[,c(7,8,12,16,17,21)]
#wo1=wo1[,c(18,19,9,10,22)]
head(pp[6],1)


pp[1]=trimws(unlist(pp[1]))  
pp[2]=trimws(unlist(pp[2]))  
pp[3]=trimws(unlist(pp[3]))  
pp[4]=trimws(unlist(pp[4]))  
pp[5]=trimws(unlist(pp[5])) 

pr_ii<- paste(pp[[1]], pp[[2]],sep="_")
pp1 <- data.frame(pr_it= pr_ii,pp)
pp_1 <- pp1[!duplicated(pp1[1],fromLast= T), ] 

#########################################################################################





# jinu given maximo data
df1= read.xlsx("I:\\1_upwork\\004_work with Jinu_____data analysis\\demo_R_Copy\\mm.xlsx",
               startRow=5, colNames=T, detectDates=T ) 

#1_5, second last 1
nrow(df1)+6                                                 
head(df1[1],1)
tail(df1[,1],1)

#df1 correction
df1[27] = as.Date( as.integer( unlist(df1[27])), origin = "1899-12-30") 
df1[34] [is.na(df1[34])] <-  df1[27] [is.na(df1[34])]                    

#df_final
df1_f= df1[1:(nrow(df1)-1), c(4,6:10,16:17,23,26,29,32:35)]   #29 ##############$$$$$$

df1_f[5]= as.numeric(unlist(df1_f[5]))


#write.xlsx(df1_f,"I:\\1_upwork\\004_work with Jinu_____data analysis\\
#demo_R_Copy\\2_half_ok_maximo_needed.xlsx",asTable = T,colNames=T)


#new sheet for output
com1_df = df1_f      
com1_df[1]=trimws(unlist(com1_df[1]))  
com1_df[2]=trimws(unlist(com1_df[2]))  
com1_df[3]=trimws(unlist(com1_df[3]))  
com1_df[6]=trimws(unlist(com1_df[6]))  
com1_df[5]=trimws(unlist(com1_df[5]))  
com1_df[9]=trimws(unlist(com1_df[9]))  
com1_df[10]=trimws(unlist(com1_df[10]))   


# Insert a new blank column named "BPA Category" at position 5
com1_df <- cbind(
  com1_df[, 1:4], 
  "BPA Category" = NA, 
  com1_df[, 5:ncol(com1_df)]
)


com1_df= com1_df[c(1,2,8,9,5,3,4,6,7,10:16)]

# Add new column to the first position
pr_3id<- paste(com1_df[[1]], com1_df[[2]],sep="_") 
pr_5id<- paste(com1_df[[1]], com1_df[[2]],com1_df[[6]],com1_df[[10]],com1_df[[11]],sep="_") 
#
com_3v_df <- data.frame(pr_it= pr_3id, com1_df) #with all inkformation
com_5v_df <- data.frame(pr_it= pr_5id, com1_df)  
#
com_3v_rm_df <- com_3v_df[!duplicated(com_3v_df[1],fromLast= T), ] 
com_5v_rm_df <- com_5v_df[!duplicated(com_5v_df[1],fromLast= T), ] 
#





############################################################################################
# for aj data
jn_gv_aj= read.xlsx("I:\\1_upwork\\004_work with Jinu_____data analysis\\demo_R_Copy\\Pending SC Action.xlsx",
                    sheet=3,colNames=T, detectDates=F) #########$
jn_gv_aj= jn_gv_aj[c(7:24)] #############################$

jn_gv_aj[3] = as.Date( as.integer( unlist(jn_gv_aj[3])), origin = "1899-12-30")
jn_gv_aj[12] = as.Date( as.integer( unlist(jn_gv_aj[12])), origin = "1899-12-30")
jn_gv_aj[15] = as.Date( as.integer( unlist(jn_gv_aj[15])), origin = "1899-12-30")



#join file 3
ju_aj1= jn_gv_aj[c(1,2,6)]
#############################$  ju_aj1[2][is.na(ju_aj1[2])] <- 1

id1<- paste(ju_aj1[[1]], ju_aj1[[2]],sep="_")
ju_3v_aj <- data.frame(ju_aj1,pr_it= id1)

aj123_r <- left_join(ju_3v_aj, com_3v_rm_df, by = "pr_it") 
aj123_r_f= aj123_r[c(1,2,7:20)]


#join file 5
ju_aj2= jn_gv_aj[c(1,2,6,10,11)]#
############$  ju_aj2[2][is.na(ju_aj2[2])] <- 1

id2<- paste(ju_aj2[[1]], ju_aj2[[2]],ju_aj2[[3]],ju_aj2[[4]],ju_aj2[[5]],sep="_") 
ju_5v_aj <- data.frame(ju_aj2,pr_it= id2)

aj12345_r <- left_join(ju_5v_aj, com_5v_rm_df, by = "pr_it")
aj12345_r_f= aj12345_r[c(1,2,9:22)]#11,12,>4,5


##################################################
aj12_rm<- left_join(ju_3v_aj, wo_1, by = "pr_it")
rm1=aj12_rm[9]

aj123_r_f=cbind(aj123_r_f,rm1)
aj12345_r_f=cbind(aj12345_r_f,rm1)

##################################################
pp12_rm<- left_join(ju_3v_aj, pp_1, by = "pr_it")
pm1=pp12_rm[10]

aj123_r_f=cbind(aj123_r_f,pm1)
aj12345_r_f=cbind(aj12345_r_f,pm1)


#################################################
#3 set for combine
jn_aj_gv_out = jn_gv_aj[c(3:18)]   #33##################3
r_aj_3v_out = aj123_r_f[3:18] #28
r_aj_5v_out = aj12345_r_f[3:18]  #28

jn_aj_gv_out[jn_aj_gv_out == ""] <- NA
r_aj_3v_out[r_aj_3v_out== ""] <- NA
r_aj_5v_out[r_aj_5v_out == ""] <- NA

# Merge the sheets, prioritizing non-blank or non-NA values
merged_data1 <- r_aj_5v_out
merged_data1[is.na(merged_data1)] <- r_aj_3v_out[is.na(merged_data1)]
merged_data1[is.na(merged_data1)] <- jn_aj_gv_out[is.na(merged_data1)]


merged_data1[9]= as.numeric(unlist(merged_data1[9]))
merged_data1[6]= as.numeric(unlist(merged_data1[6]))


#final AJ file
write.xlsx(merged_data1,
           "I:\\1_upwork\\004_work with Jinu_____data analysis\\demo_R_Copy\\3_NTP_Maximo.xlsx",
           asTable = T,colNames=T)

