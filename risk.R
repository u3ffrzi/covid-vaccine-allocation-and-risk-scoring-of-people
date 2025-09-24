library(readstata13)
library(tidyverse)
pop=read.dta13("/home/naser/Desktop/risk/pop_district_80-92_single_age.dta")
pop=pop[pop$year==1396,]

pop=aggregate(pred2pop~sex_name+district_code+age,data = pop,FUN = sum)
pop$pred2pop=round(pop$pred2pop,digits = 0)
pop2=pop %>% uncount(pred2pop)
write_csv(pop2,"/home/naser/Desktop/risk/frame.csv")
#####################################################################
library(readstata13)
library(openxlsx)
library(tidyverse)
library(readr)
library(haven)
prevalence=read_dta("/home/naser/Desktop/risk/prevalence.dta")
prevalence$sex_name[prevalence$sex_name=="Female"]="female"
prevalence$sex_name[prevalence$sex_name=="Male"]="male"
colnames(prevalence)=c("sex_name","age_cat","liver_disease_prevalence","kidney_disease_prevalence",
                       "cardiovascular_prevalence","diabetes_prevalence","idd_prevalence","malignancy_prevalence","province_code")
prevalence$age=NA
name=c(1,10,15,20,25,30,35,40,45,5,50,55,60,65,70,78,80,85,0)
age=data.frame(age_cat=unique(prevalence$age_cat),freq=c(4,rep(5,16),1,1),first=c(1,10,15,20,25,30,35,40,45,5,50,55,60,65,70,75,80,85,0))
PR=prevalence[1,]
PR=PR[-1,]
for (i in unique(prevalence$age_cat)){
  pr=prevalence[prevalence$age_cat==i,]
  fr=age$freq[age$age_cat==i]
  first=age$first[age$age_cat==i]
  j=0
  while (j<fr) {
    pr$age=first+j
    PR=bind_rows(PR,pr)
    j=j+1
  }
}
prevalence=PR[,-2]
########################################################################
population=read.dta13("/home/naser/Desktop/risk/pop_district_80-92_single_age.dta")
population=population[population$year==1396,]
population=aggregate(pred2pop~population$sex_name+population$age+district_code,data=population,FUN=sum)
colnames(population)=c("sex_name","age","district_code","pred2pop")
population$pred2pop=round(population$pred2pop,digits = 0)
pop=read_csv("/home/naser/Desktop/risk/frame.csv")
pop$code=1:nrow(pop)
pop$malignancy=0
pop$idd=0
pop$liver_disease=0
pop$kidney_disease =0
pop$diabetes=0
pop$cardiovascular=0
pop$death=NA
pop$province_code=substr(pop$district_code,1,2)
district="0503"
sex="male"
age="0"
Data=data.frame(pop[1,])
Data=Data[-1,]
population$age=as.numeric(population$age)
pop=left_join(pop,population,by=c("district_code","sex_name","age"))
pop=left_join(pop,prevalence,by=c("province_code","sex_name","age"))
pop=data.frame(pop)

# pop2=pop[ pop$age==80 | pop$age==81 | pop$age==82,]
# pr=prevalence[ prevalence$age==80 | prevalence$age==81 | prevalence$age==82,]

f=function(data){
  data$malignancy[sample(1:nrow(data),  round(unique(data$malignancy_prevalence)*unique(data$pred2pop),digits = 0))]=1
  data$idd[sample(1:nrow(data),  round(unique(data$idd_prevalence)*unique(data$pred2pop),digits = 0))]=1
  data$liver_disease[sample(1:nrow(data),  round(unique(data$liver_disease_prevalence)*unique(data$pred2pop),digits = 0))]=1
  data$kidney_disease[sample(1:nrow(data),  round(unique(data$kidney_disease_prevalence)*unique(data$pred2pop),digits = 0))]=1
  data$diabetes[sample(1:nrow(data),  round(unique(data$diabetes_prevalence)*unique(data$pred2pop),digits = 0))]=1
  data$cardiovascular[sample(1:nrow(data),  round(unique(data$cardiovascular_prevalence)*unique(data$pred2pop),digits = 0))]=1
  return(data)
}

simulated_data=pop %>% group_by(district_code,age,sex_name) %>%do(f(.data))


write_csv(simulated_data,"/home/naser/Desktop/risk/simulated_data.csv")

###########

simulated_data=read_csv("/home/naser/Desktop/risk/simulated_data.csv")

covid=read.xlsx("/home/naser/Desktop/risk/covid.xlsx")
# covid=covid[,-ncol(covid)]
# colnames(covid)=c(h[1],"district_code",h[3:11])
pop=simulated_data[,c(1,2,3,4,5,6,7,8,9,10,11)]

Data=bind_rows(pop,covid)
k=glm(death ~ sex_name + age + malignancy+idd+liver_disease+kidney_disease+diabetes+cardiovascular, data = Data, family = binomial(link = "logit"))
Data$predict=predict(k,Data,type = "response")

write_csv(Data,"/home/naser/Desktop/risk/predicted_data.csv")


#######
library(readstata13)
library(openxlsx)
library(tidyverse)
library(readr)
library(haven)

Data=read_csv("/home/naser/Desktop/risk/predicted_data.csv")

Data=Data[!is.na(Data$code),]
Data1=Data[Data$district_code=="0503" & Data$age==73 & Data$,]
k=381219.5049/81025924
# cut1=min(Data$predict)+k
# cut2=min(Data$predict)+2*k
# cut3=min(Data$predict)+3*k
# cut4=min(Data$predict)+4*k
cut1=5*k
cut2=20*k
cut3=40*k
cut4=60*k


Data$cat1=NA
Data$cat2=NA
Data$cat3=NA
Data$cat4=NA
Data$cat5=NA
Data$cat1[Data$predict<cut1]=1
Data$cat2[Data$predict<cut2 & Data$predict>=cut1]=1
Data$cat3[Data$predict<cut3 & Data$predict>=cut2]=1
Data$cat4[Data$predict<cut4 & Data$predict>=cut3]=1
Data$cat5[ Data$predict>=cut4]=1
# risk=aggregate(cbind(cat1,cat2,cat3,cat4,cat5)~district_code,data = Data,FUN =function(x) sum(x,na.rm = T))
d1=Data %>% group_by(district_code)%>%summarise(cat1=sum(cat1,na.rm = T),
                                                cat2=sum(cat2,na.rm = T),
                                                cat3=sum(cat3,na.rm = T),
                                                cat4=sum(cat4,na.rm = T),
                                                cat5=sum(cat5,na.rm = T),)

d3 = d1 %>% summarise(cat1s=sum(cat1), cat2s=sum(cat2),cat3s=sum(cat3),cat4s=sum(cat4),cat5s=sum(cat5))
        
write_csv(d1,"/home/naser/Desktop/risk/aggregated_data.csv")
