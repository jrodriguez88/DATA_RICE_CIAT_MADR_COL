#############################



#### Package and INPUTS Requeriments

if(require(dplyr)==FALSE){install.packages("dplyr")}
if(require(stringr)==FALSE){install.packages("stringr")}
if(require(openxlsx)==FALSE){install.packages("openxlsx")}

#### CORDINATES
#LOC_ID	LON	LAT	ALT
#MRCO	-75.8544	8.8109	15
#AIHU	-75.2406	3.2530	380
#SDTO	-74.9848	3.9136	415
#VVME	-73.4743	4.0294	315
#YOCS	-72.2983	5.3302	250



read_INPUT_data <- function(filename) {
    sheets <- readxl::excel_sheets(filename)
    x <-    lapply(sheets, function(X) readxl::read_excel(filename, sheet = X))
    names(x) <- sheets
    x
}
INPUT_data <- read_INPUT_data("AIPE ENSAYOS  MADR.xlsx")

local <- "AIHU"
CROP_SYS <- "IRRIGATED" 
ESTAB <- "DIRECT-SEED"
LAT <- 3.253
LON <- -75.24
ALT <- 380

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
    summarize_all(funs(mean), na.rm=TRUE)%>%
    tbl_df()%>%
    extract_id()%>%
    mutate(CULTIVAR=genot,
           id_s=substr(ID,5,nchar(ID)))%>%
    select(ID, id_s, CULTIVAR, planxmetro)


#grep(pattern = substr(colnames(trial_info)[2:ncol(trial_info)],1, 6), x = plant_density$ID)

### Row spacing (m)
id_s <- str_sub(colnames(trial_info)[2:ncol(trial_info)], 1,-10)
dist_s <- str_sub(trial_info[5,2:ncol(trial_info)], 1,4)

row_spa <- cbind(id_s, dist_s)%>%
    tbl_df()%>%
    mutate(dist_s=as.numeric(dist_s))

plant_density <- dplyr::left_join(plant_density, row_spa, by="id_s")%>%
    mutate(NPLDS=round(planxmetro*(1/dist_s)))

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


#### Merge df to PLANT_gro
    
PLANT_gro <- Reduce(function(x, y) {merge(x, y, all.x=TRUE)}, list(HV, AF, HM, ST,  MP, MT))%>%
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

for (i in 1:length(AGRO_man)) {
    
    list_of_datasets <- list("AGRO_man" = AGRO_man[[i]], "PHEN_obs" = arrange(PHEN_obs[[i]], PHEN_obs[[i]]$ID), "FERT_obs"= FERT_obs[[i]], "PLANT_gro" = PLANT_gro[[i]])
    write.xlsx(list_of_datasets, file = paste0(local,"_",names(AGRO_man)[i],".xlsx"))
    
}

}
make_xls_by_cul()


