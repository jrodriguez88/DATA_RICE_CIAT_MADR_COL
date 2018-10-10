#############################



#### Package and INPUTS Requeriments

if(require(tidyverse)==FALSE){install.packages("tidyverse")}
if(require(openxlsx)==FALSE){install.packages("openxlsx")}


#### CORDINATES
#LOC_ID	LON	LAT	ALT
#MRCO	-75.8544	8.8109	15
#AIHU	-75.2406	3.2530	380
#SDTO	-74.9848	3.9136	415
#VVME	-73.4743	4.0294	315
#YOCS	-72.2983	5.3302	250

local <- "AIHU"
CROP_SYS <- "IRRIGATED" 
ESTAB <- "DIRECT-SEED"
LAT <- 3.253
LON <- -75.24
ALT <- 380

read_INPUT_data <- function(filename) {
    sheets <- readxl::excel_sheets(filename)
    x <-    lapply(sheets, function(X) readxl::read_excel(filename, sheet = X))
    names(x) <- sheets
    x
}
INPUT_data <- read_INPUT_data("AIPE ENSAYOS  MADR.xlsx")


extract_id <- function(tb){
    id <- mutate(tb,ID=paste0(substr(gsub('^(.{1}).', local, x = tb$siembra),1,4),
                              substr(tb$siembra, 4, 5),
                              str_sub(tb$siembra, 6,-10) ))
    return(id)
}  



### Info Trial (GENERALIDADES)

trial_info <- INPUT_data$`GENERALIDADES DEL ENSAYO`%>%
    na.omit()

### Cal Plant Density

#plant_density <- INPUT_data$`REP MUESTREOS`%>%
#    group_by(genot,siembra)%>%
#    na.omit()%>%
#    summarize_all(funs(mean))%>%
#    tbl_df()%>%
#    extract_id()%>%
#    mutate(CULTIVAR=genot,
#           id_s=substr(ID,5,nchar(ID)))%>%
#    select(ID, id_s, CULTIVAR, planxmetro)

plant_density <- INPUT_data$`REP MUESTREOS`%>%
    group_by(genot,siembra)%>%
    summarize_all(funs(mean, sd), na.rm=TRUE)%>%
    tbl_df()%>%
    extract_id()%>%
    mutate(CULTIVAR=genot,
           id_s=substr(ID,5,nchar(ID)))%>%
    select(ID, id_s, CULTIVAR, planxmetro_mean, planxmetro_sd)


#grep(pattern = substr(colnames(trial_info)[2:ncol(trial_info)],1, 6), x = plant_density$ID)

### Row spacing (m)
id_s <- str_sub(colnames(trial_info)[2:ncol(trial_info)], 1,-10)
dist_s <- str_sub(trial_info[5,2:ncol(trial_info)], 1,4)

row_spa <- cbind(id_s, dist_s)%>%
    tbl_df()%>%
    mutate(dist_s=as.numeric(dist_s))

plant_density <- dplyr::left_join(plant_density, row_spa, by="id_s")%>%
    mutate(NPLDS=round(planxmetro_mean*(1/dist_s)))

for (i in 1:nrow(plant_density)){
    
    if (plant_density$dist_s[i]==0.20){
        plant_density$NPLDS[i]= (plant_density$NPLDS[i])*2
    }
}

    



### Read Phenological Data

PHEN_obs <- INPUT_data$`REP FENOLOGIA` %>%
    select(grep("INDICADORES", names(INPUT_data$`REP FENOLOGIA`), value = F ):ncol(INPUT_data$`REP FENOLOGIA`))%>%
    tbl_df()%>%
    na.omit()%>%
    mutate(siembra=`PARA INDICADORES`,
           CULTIVAR=`Row Labels__1`,
           LOC_ID=local,
           PDAT=as.Date(`Average of FSIEM__1`, format("%m.%d.%Y")),
           EDAT=as.Date(`Average of FEMER__1`, format("%m.%d.%Y")),
           IDAT=as.Date(`Average of FIP__1`, format("%m.%d.%Y")),
           FDAT=as.Date(`Average of FLO50__1`, format("%m.%d.%Y")),
           MDAT=as.Date(`Average of FCOS__1`, format("%m.%d.%Y")))%>%
    extract_id()%>%
    dplyr::select(ID, LOC_ID, CULTIVAR, PDAT, EDAT, IDAT, FDAT, MDAT)



##### Read Fertilization Table
FERT_obs <- INPUT_data$FERTILIZACION %>%
            na.omit()

FERT_obs$Aplicacion[grep("abona", FERT_obs$Aplicacion)] <- "1_Preabonamiento"
FERT_obs$DDE[FERT_obs$DDE<1] <- 1

FERT_obs <- mutate(FERT_obs,
                        siembra= Siembra,
                        LOC_ID=local,
                        DDE = as.double(DDE),
                        FERT_No= as.numeric(as.character(substr(FERT_obs$Aplicacion,1,1))),
                        N= as.double(`N Kg/ha`),
                        P= as.double(`P kg/ha`),
                        K= as.double(FERT_obs$`K kg/ha`))%>%
            extract_id()%>%
            na.omit()%>%
            merge(plant_density[c(1,3)])%>%
            dplyr::select(ID, LOC_ID, CULTIVAR, FERT_No, DDE, N, P, K)


if( local== "AIHU"){
   ID_AIHU <- INPUT_data$`REP MUESTREOS`%>%
    group_by(genot,siembra, muestreo)%>%
    summarize_all(funs(max))%>%
    select(genot, siembra,muestreo,fmues)

INPUT_data$AF <- merge(ID_AIHU,INPUT_data$AF, by.x=c("genot", "muestreo", "fmues"), by.y = c("genot", "muestreo", "fmues"), all.x = T )%>%
                tbl_df()
INPUT_data$MSHV  <- merge(ID_AIHU,INPUT_data$MSHV  , by.x=c("genot", "muestreo", "fmues"), by.y = c("genot", "muestreo", "fmues"), all.x = T )%>%
    tbl_df()%>%
    mutate(siembra=siembra.x)
INPUT_data$MSTALL<- merge(ID_AIHU,INPUT_data$MSTALL, by.x=c("genot", "muestreo", "fmues"), by.y = c("genot", "muestreo", "fmues"), all.x = T )%>%
    tbl_df()%>%
    mutate(siembra=siembra.x)
INPUT_data$MSHM  <- merge(ID_AIHU,INPUT_data$MSHM  , by.x=c("genot", "muestreo", "fmues"), by.y = c("genot", "muestreo", "fmues"), all.x = T )%>%
    tbl_df()%>%
    mutate(siembra=siembra.x)
INPUT_data$MSPAN <- merge(ID_AIHU,INPUT_data$MSPAN , by.x=c("genot", "muestreo", "fmues"), by.y = c("genot", "muestreo", "fmues"), all.x = T )%>%
    tbl_df()%>%
    mutate(siembra=siembra.x)
INPUT_data$MSTOT <- merge(ID_AIHU,INPUT_data$MSTOT , by.x=c("genot", "muestreo", "fmues"), by.y = c("genot", "muestreo", "fmues"), all.x = T )%>%
    tbl_df()%>%
    mutate(siembra=siembra.x) 
    
}


### Extract PLANT_gro

AF <- INPUT_data$AF %>%
    na.omit()%>%
    tbl_df()%>%
    extract_id()%>%
    mutate(LAI_OBS=(`AFm2 (cm2/m2)`)/10000, 
           LAI_SD=(stderr_mean__1)/10000,
           CULTIVAR=genot,
           SAMPLING_DATE=as.Date(fmues, format("%m.%d.%Y")))%>%
    dplyr::select(ID, CULTIVAR, SAMPLING_DATE,LAI_OBS, LAI_SD)
 

HV <- INPUT_data$MSHV %>%
    na.omit()%>%
    tbl_df()%>%
    extract_id()%>%
    mutate(WLVG_OBS=(`MSHVm2 (g/m2)`)*10, 
           WLVG_SD=(stderr_mean__1)*10,
           CULTIVAR=genot,
           SAMPLING_DATE=as.Date(fmues, format("%m.%d.%Y")))%>%
    dplyr::select(ID, CULTIVAR, SAMPLING_DATE,WLVG_OBS, WLVG_SD)

ST <- INPUT_data$MSTALL %>%
    na.omit()%>%
    tbl_df()%>%
    extract_id()%>%
    mutate(WST_OBS=(`MSTALLm2 (g/m2)`)*10, 
           WST_SD=(stderr_mean__1)*10,
           CULTIVAR=genot,
           SAMPLING_DATE=as.Date(fmues, format("%m.%d.%Y")))%>%
    dplyr::select(ID, CULTIVAR, SAMPLING_DATE,WST_OBS, WST_SD)

HM <- INPUT_data$MSHM %>%
    na.omit()%>%
    tbl_df()%>%
    extract_id()%>%
    mutate(WLVD_OBS=(`MSHMm2 (g/m2)`)*10, 
           WLVD_SD=(stderr_mean__1)*10,
           CULTIVAR=genot,
           SAMPLING_DATE=as.Date(fmues, format("%m.%d.%Y")))%>%
    dplyr::select(ID, CULTIVAR, SAMPLING_DATE,WLVD_OBS, WLVD_SD)


MP <- INPUT_data$MSPAN %>%
    na.omit()%>%
    tbl_df()%>%
    extract_id()%>%
    mutate(WSO_OBS=(`MSPANm2 (g/m2)`)*10, 
           WSO_SD=(stderr_mean__1)*10,
           CULTIVAR=genot,
           SAMPLING_DATE=as.Date(fmues, format("%m.%d.%Y")))%>%
    dplyr::select(ID, CULTIVAR, SAMPLING_DATE,WSO_OBS, WSO_SD)

MT <- INPUT_data$MSTOT %>%
    na.omit()%>%
    tbl_df()%>%
    extract_id()%>%
    mutate(WAGT_OBS=(`MSTOTm2 (g/m2)`)*10, 
           WAGT_SD=(stderr_mean__1)*10,
           CULTIVAR=genot,
           SAMPLING_DATE=as.Date(fmues, format("%m.%d.%Y")))%>%
    dplyr::select(ID, CULTIVAR, SAMPLING_DATE,WAGT_OBS, WAGT_SD)

NH <- INPUT_data$NHV %>%
    na.omit()%>%
    tbl_df()%>%
    extract_id()%>%
    mutate(NLV_OBS=round(NHVm2), 
           NLV_SD=round(stderr_mean__1),
           CULTIVAR=genot,
           SAMPLING_DATE=as.Date(fmues, format("%m.%d.%Y")))%>%
    dplyr::select(ID, CULTIVAR, SAMPLING_DATE,NLV_OBS, NLV_SD)

NS <- INPUT_data$NTALL %>%
    na.omit()%>%
    tbl_df()%>%
    extract_id()%>%
    mutate(NST_OBS=round(NTALLm2), 
           NST_SD=round(stderr_mean__1),
           CULTIVAR=genot,
           SAMPLING_DATE=as.Date(fmues, format("%m.%d.%Y")))%>%
    dplyr::select(ID, CULTIVAR, SAMPLING_DATE,NST_OBS, NST_SD)

NP <- INPUT_data$NPAN %>%
    na.omit()%>%
    tbl_df()%>%
    extract_id()%>%
    mutate(NP_OBS=round(NPANm2), 
           NP_SD=round(stderr_mean__1),
           CULTIVAR=genot,
           SAMPLING_DATE=as.Date(fmues, format("%m.%d.%Y")))%>%
    dplyr::select(ID, CULTIVAR, SAMPLING_DATE,NP_OBS, NP_SD)

###### Read Weather data
WDATA <- INPUT_data$`QC CLIMA 2013 - 2016` %>% 
    tbl_df() %>%
    mutate(DATE = as.Date(Date, format("%m.%d.%Y")),
           DAY = day(Date),
           MONTH = month(Date),
           YEAR = year(Date),
           RAIN = Rain,
           SRAD = RF_CCM2_ESOL*0.041868,
           RHUM = RHUM_FUS) %>%
    dplyr::select(DATE, TMAX, TMIN, RAIN, SRAD, RHUM)
    
#library(ggplot2)
#ggplot(data=WDATA, aes(x=as.factor(month(DATE)), y=TMAX)) + geom_boxplot()

##### Read Soil Data

#Sol_dat <- 





### Read YIELD data

YIELD_raw <- INPUT_data$`REP RTO Y COMPONENTES` %>%
    tbl_df()%>%
    extract_id()%>%
    mutate(LOC_ID = local,
           CULTIVAR=genot,
           YIELD_LM= RTO_ML,
           YIELD_AREA= RTO_AREA,
           YIELD_CD= RTO_CD, 
           YIELD_POT1= NTXM2*GXPAN*(P1000G/1000)*10,
           YIELD_POT2= NPXM2*GLLXPAN*(P1000G/1000)*10) %>%
#           YIELD_AVG= mean(c(RTO_ML, RTO_AREA, RTO_CD), na.rm = T)) %>%
    dplyr::select(ID, LOC_ID, CULTIVAR, YIELD_LM, YIELD_AREA, YIELD_CD, YIELD_POT1, YIELD_POT2, IC, PFERT, P1000G, NPXM2, NTXM2, GLLXPAN, GXPAN)


YIELD_raw %>% select(1:6) %>% gather(key= "YIELD_M", value =  "YIELD_V", -c(ID, LOC_ID, CULTIVAR)) %>%
    mutate(label_id=substr(ID, 5, nchar(ID)))%>%
    ggplot(aes(label_id, YIELD_V, color=YIELD_M))+
    stat_summary(fun.data=mean_cl_boot, position = position_dodge(width=0.2))+
    facet_wrap(~CULTIVAR, scales = "free_x")+
    theme_bw()+ theme(axis.text.x = element_text(angle = 45))

#YIELD_raw %>% select(1:7) %>% gather(key= "YIELD_M", value =  "YIELD_V", -c(ID, LOC_ID, CULTIVAR)) %>%
#    mutate(label_id=substr(ID, 5, nchar(ID)))%>%
#    ggplot(aes(label_id, YIELD_V, fill=YIELD_M))+geom_boxplot()+facet_wrap(~CULTIVAR, scales = "free_x")+
#    theme_bw()+ theme(axis.text.x = element_text(angle = 45))


#yield_stat <- YIELD_raw %>% select(1:7) %>% 
#    gather(key= "YIELD_M", value =  "YIELD_V", -c(ID, LOC_ID, CULTIVAR)) %>% 
#    group_by(CULTIVAR, ID, YIELD_M) %>%
#    summarise_at("YIELD_V", .funs=c(y=mean, ymin=min, ymax=max), na.rm=T)
#    
#    summarise(y=smean.cl.boot(YIELD_V, conf.int=.95, B=1000, na.rm=TRUE, reps=FALSE)[1],
#           ymin=smean.cl.boot(YIELD_V, conf.int=.95, B=1000, na.rm=TRUE, reps=FALSE)[2],
#           ymax=smean.cl.boot(YIELD_V, conf.int=.95, B=1000, na.rm=TRUE, reps=FALSE)[3])

YIELD_obs <- INPUT_data$`RTO - COMP E INDICADORES` %>%
    tbl_df()%>%
    extract_id()%>%
    mutate(LOC_ID = local,
           CULTIVAR=genot,
           
           

#### Merge df to PLANT_gro

ORG_Num <- Reduce(function(x, y) {merge(x, y, all.x=TRUE)}, list(NH, NS, NP))%>%
    tbl_df()%>%
    mutate(LOC_ID=local)%>%
    dplyr::select(ID, LOC_ID, everything())
    
PLANT_gro <- Reduce(function(x, y) {merge(x, y, all.x=TRUE)}, list(HV, AF, HM, ST,  MP, MT, ORG_Num))%>%
        tbl_df()%>%
        mutate(LOC_ID=local)%>%
        dplyr::select(ID, LOC_ID, everything())
   
### Fill WSO=NA--->0 
    for (i in 1:nrow(PLANT_gro)) {
        
         if (PLANT_gro$WLVG_OBS[i]>0 && PLANT_gro$WST_OBS[i]>0 && is.na(PLANT_gro$WLVD_OBS[i])) {
            PLANT_gro$WLVD_OBS[i]=0
            PLANT_gro$WLVD_SD[i]=0
         } 
        
         if (PLANT_gro$WLVG_OBS[i]>0 && PLANT_gro$WST_OBS[i]>0 && is.na(PLANT_gro$WSO_OBS[i])){
        PLANT_gro$WSO_OBS[i]=0
        PLANT_gro$WSO_SD[i]=0
    }
        
  }   
 



## ORYZA BIOMASS must have integer sum(WLVG_OBS,WLVD_OBS,WST_OBS,WSO_OBS)==WAGT_OBS

PLANT_gro[5:ncol(PLANT_gro)] <- round(PLANT_gro[5:ncol(PLANT_gro)],1)   
for (i in 1: nrow(PLANT_gro)) {
    PLANT_gro$WAGT_OBS[i] =sum(PLANT_gro$WLVG_OBS[i],PLANT_gro$WLVD_OBS[i],PLANT_gro$WST_OBS[i],PLANT_gro$WSO_OBS[i])
}




### Create AGRO_man
AGRO_man <- merge(PHEN_obs, plant_density, all.x = T)%>%
    mutate(PROJECT= substr(ID,7,nchar(ID)),
           LAT=LAT,
           LONG=LON,
           ALT=ALT,
           CROP_SYS=CROP_SYS,
           ESTAB=ESTAB,
           TR_N=substr(ID,5,6)) %>%
    select(ID, LOC_ID, PROJECT, CULTIVAR, TR_N, LAT, LONG, ALT, PDAT, CROP_SYS, ESTAB, NPLDS)






### Create WB xlsx
if(devtools::find_rtools()) Sys.setenv(R_ZIPCMD= file.path(devtools:::get_rtools_path(),"zip"))

make_xls_by_cul <- function(){
    
AGRO_man <- split(AGRO_man, AGRO_man$CULTIVAR)
PLANT_gro <- split(PLANT_gro, PLANT_gro$CULTIVAR)
PHEN_obs <- split(PHEN_obs, PHEN_obs$CULTIVAR)
FERT_obs <- split(FERT_obs, FERT_obs$CULTIVAR) 
WTH_obs <- WDATA

for (i in 1:length(AGRO_man)) {
    
    list_of_datasets <- list("AGRO_man" = AGRO_man[[i]],
                             "PHEN_obs" = arrange(PHEN_obs[[i]], PHEN_obs[[i]]$ID), 
                             "FERT_obs" = FERT_obs[[i]], 
                             "PLANT_gro"= PLANT_gro[[i]],
                             "WTH_obs"  = WTH_obs)
    write.xlsx(list_of_datasets, file = paste0(local,"_",names(AGRO_man)[i],".xlsx"))
    
}

}
make_xls_by_cul()


