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
df1_f= df1[1:(nrow(df1)-1), c(4,6:10,15:17,11,12,19:23,26,29:37)]    

df1_f[5]= as.numeric(unlist(df1_f[5]))
df1_f[10]= as.numeric(unlist(df1_f[10]))
df1_f[11]= as.numeric(unlist(df1_f[11]))

#write.xlsx(df1_f,"I:\\1_upwork\\004_work with Jinu_____data analysis\\
           #demo_R_Copy\\2_half_ok_maximo_needed.xlsx",asTable = T,colNames=T)


#new sheet for output
com1_df = df1_f      
com1_df[1]=trimws(unlist(com1_df[1]))  
com1_df[2]=trimws(unlist(com1_df[2]))  
com1_df[3]=trimws(unlist(com1_df[3]))  
com1_df[6]=trimws(unlist(com1_df[6]))  
com1_df[5]=trimws(unlist(com1_df[5]))  
com1_df[16]=trimws(unlist(com1_df[16]))  
com1_df[17]=trimws(unlist(com1_df[17]))   

differences3 <- com1_df!= df1_f
a=colSums(differences3==T, na.rm=T)
sum(a)
a
if (any(!is.na(differences3) & differences3)) {
  cat("Differences found at the following indices:\n")
  print(which(!is.na(differences3) & differences3, arr.ind = TRUE))
} else {
  cat("The datasets are identical.\n")
}





# Add new column to the first position
pr_3id<- paste(com1_df[[1]], com1_df[[2]],sep="_") 
pr_5id<- paste(com1_df[[1]], com1_df[[2]],com1_df[[3]],com1_df[[16]],com1_df[[17]],sep="_") 
#
com_3v_df <- data.frame(pr_it= pr_3id, com1_df) #with all inkformation
com_5v_df <- data.frame(pr_it= pr_5id, com1_df)  
#
com_3v_rm_df <- com_3v_df[!duplicated(com_3v_df[1],fromLast= T), ] 
com_5v_rm_df <- com_5v_df[!duplicated(com_5v_df[1],fromLast= T), ] 
#





############################################################################################
# for aj data
jn_gv_aj= read.xlsx("I:\\1_upwork\\004_work with Jinu_____data analysis\\demo_R_Copy\\Material Status Master Tracker.xlsx",
                    sheet=1,colNames=T, detectDates=F) 


jn_gv_aj[2] = as.Date( as.integer( unlist(jn_gv_aj[2])), origin = "1899-12-30")
jn_gv_aj[10] = as.Date( as.integer( unlist(jn_gv_aj[10])), origin = "1899-12-30")
jn_gv_aj[11] = as.Date( as.integer( unlist(jn_gv_aj[11])), origin = "1899-12-30")
jn_gv_aj[20] = as.Date( as.integer( unlist(jn_gv_aj[20])), origin = "1899-12-30")
jn_gv_aj[21] = as.Date( as.integer( unlist(jn_gv_aj[21])), origin = "1899-12-30")
jn_gv_aj[24] = as.Date( as.integer( unlist(jn_gv_aj[24])), origin = "1899-12-30")
jn_gv_aj[26] = as.Date( as.integer( unlist(jn_gv_aj[26])), origin = "1899-12-30")
jn_gv_aj[29] = as.Date( as.integer( unlist(jn_gv_aj[29])), origin = "1899-12-30")
jn_gv_aj[31] = as.Date( as.integer( unlist(jn_gv_aj[31])), origin = "1899-12-30")
#join file 3
ju_aj1= jn_gv_aj[c(9,13,7)]
ju_aj1[2][is.na(ju_aj1[2])] <- 1

id1<- paste(ju_aj1[[1]], ju_aj1[[2]],sep="_")
ju_3v_aj <- data.frame(ju_aj1,pr_it= id1)

aj123_r <- left_join(ju_3v_aj, com_3v_rm_df, by = "pr_it") 
aj123_r_f= aj123_r[c(3,1,2,4,11,12,13,6,10,9,14:30)]


#join file 5
ju_aj2= jn_gv_aj[c(9,13,7,22,23)]#
ju_aj2[2][is.na(ju_aj2[2])] <- 1

id2<- paste(ju_aj2[[1]], ju_aj2[[2]],ju_aj2[[3]],ju_aj2[[4]],ju_aj2[[5]],sep="_") 
ju_5v_aj <- data.frame(ju_aj2,pr_it= id2)

aj12345_r <- left_join(ju_5v_aj, com_5v_rm_df, by = "pr_it")
aj12345_r_f= aj12345_r[c(3,1,2,6,13,14,15,8,12,11,16:32)]#11,12,>4,5


##################################################
aj12_rm<- left_join(ju_3v_aj, wo_1, by = "pr_it")
rm1=aj12_rm[9]

aj123_r_f=cbind(aj123_r_f,rm1)
aj12345_r_f=cbind(aj12345_r_f,rm1)

#################################################
#3 set for combine
jn_aj_gv_out = jn_gv_aj[c(11:33)]   #33##################3
r_aj_3v_out = aj123_r_f[6:28] #28
r_aj_5v_out = aj12345_r_f[6:28]  #28

jn_aj_gv_out[jn_aj_gv_out == ""] <- NA
r_aj_3v_out[r_aj_3v_out== ""] <- NA
r_aj_5v_out[r_aj_5v_out == ""] <- NA

# Merge the sheets, prioritizing non-blank or non-NA values
merged_data1 <- r_aj_5v_out
merged_data1[is.na(merged_data1)] <- r_aj_3v_out[is.na(merged_data1)]
merged_data1[is.na(merged_data1)] <- jn_aj_gv_out[is.na(merged_data1)]

merged_data1[3]= as.numeric(unlist(merged_data1[3]))
merged_data1[5]= as.numeric(unlist(merged_data1[5]))
merged_data1[6]= as.numeric(unlist(merged_data1[6]))
merged_data1[7]= as.numeric(unlist(merged_data1[7]))

#final AJ file
write.xlsx(merged_data1,
    "I:\\1_upwork\\004_work with Jinu_____data analysis\\demo_R_Copy\\1_AJ_final_rm.xlsx",
    asTable = T,colNames=T)



##############################################################################################
# for bu data
jn_gv_bu= read.xlsx("I:/1_upwork/004_work with Jinu_____data analysis/demo_R_Copy/Material Status Master Tracker.xlsx",
                    sheet=2,colNames=T, detectDates=F)  




jn_gv_bu[2] = as.Date( as.integer( unlist(jn_gv_bu[2])), origin = "1899-12-30")
jn_gv_bu[10] = as.Date( as.integer( unlist(jn_gv_bu[10])), origin = "1899-12-30")
jn_gv_bu[11] = as.Date( as.integer( unlist(jn_gv_bu[11])), origin = "1899-12-30")
jn_gv_bu[20] = as.Date( as.integer( unlist(jn_gv_bu[20])), origin = "1899-12-30")
jn_gv_bu[21] = as.Date( as.integer( unlist(jn_gv_bu[21])), origin = "1899-12-30")
jn_gv_bu[24] = as.Date( as.integer( unlist(jn_gv_bu[24])), origin = "1899-12-30")
jn_gv_bu[26] = as.Date( as.integer( unlist(jn_gv_bu[26])), origin = "1899-12-30")
jn_gv_bu[29] = as.Date( as.integer( unlist(jn_gv_bu[29])), origin = "1899-12-30")
jn_gv_bu[31] = as.Date( as.integer( unlist(jn_gv_bu[31])), origin = "1899-12-30")
#join file 3
ju_bu1= jn_gv_bu[c(9,13,7)]
ju_bu1[2][is.na(ju_bu1[2])] <- 1

id2<- paste(ju_bu1[[1]], ju_bu1[[2]],sep="_") #
ju_3v_bu <- data.frame(ju_bu1,pr_it= id2)

bu123_r <- left_join(ju_3v_bu, com_3v_rm_df, by = "pr_it")
bu123_r_f= bu123_r[c(3,1,2,4,11,12,13,6,10,9,14:30)]

#join file 5
ju_bu2= jn_gv_bu[c(9,13,7,22,23)]
ju_bu2[2][is.na(ju_bu2[2])] <- 1

id22<- paste(ju_bu2[[1]], ju_bu2[[2]],ju_bu2[[3]],ju_bu2[[4]],ju_bu2[[5]],sep="_") 
ju_5v_bu <- data.frame(ju_bu2,pr_it= id22)

bu12345_r <- left_join(ju_5v_bu, com_5v_rm_df, by = "pr_it")
bu12345_r_f= bu12345_r[c(3,1,2,6,13,14,15,8,12,11,16:32)] #11,12,>4,5

######
bu12_rm<- left_join(ju_3v_bu, wo_1, by = "pr_it")
rm2=bu12_rm[9]

bu123_r_f=cbind(bu123_r_f,rm2)
bu12345_r_f=cbind(bu12345_r_f,rm2)

######


# combile 3 data set
jn_bu_gv_out = jn_gv_bu[c(11:33)]
r_bu_3v_out = bu123_r_f[6:28]
r_bu_5v_out = bu12345_r_f[6:28]


jn_bu_gv_out[jn_bu_gv_out == ""] <- NA
r_bu_3v_out[r_bu_3v_out== ""] <- NA
r_bu_5v_out[r_bu_5v_out == ""] <- NA

# Merge the sheets, prioritizing non-blank or non-NA values
merged_data11 <- r_bu_5v_out
merged_data11[is.na(merged_data11)] <- r_bu_3v_out[is.na(merged_data11)]
merged_data11[is.na(merged_data11)] <- jn_bu_gv_out[is.na(merged_data11)]

merged_data11[3]= as.numeric(unlist(merged_data11[3]))
merged_data11[5]= as.numeric(unlist(merged_data11[5]))
merged_data11[6]= as.numeric(unlist(merged_data11[6]))
merged_data11[7]= as.numeric(unlist(merged_data11[7]))


# final bu file
write.xlsx(merged_data11,"I:\\1_upwork\\004_work with Jinu_____data analysis\\demo_R_Copy\\1_BU_final_rm.xlsx",
           asTable = T,colNames=T)


########################################################################################
