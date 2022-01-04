
require(pacman)

p_load(tidyverse, openxlsx, data.table)
options(scipen=999)


datos <- fread("input/TABLA_VALORES_VIG_2020_V7.txt", header=T, sep="|")
correlativa <- fread("input/tabla_correlativa.csv", header=T)
datos1 <- datos[, c(2:16,25,26)]
datos2 <- datos1 %>% left_join(correlativa, by = c("EC_MO_ID" = "EC_MO_ID"))

reemplaza_coma_punto <- function(.x){str_replace_all(.x, ",", ".")}


datos2 %>%
  mutate_at(c("AREA_TERRENO1", "AREA_TERRENO2", "AREA_USO1", "AREA_USO2"),
            reemplaza_coma_punto)


datos2[, "AREA_TERRENO1"]<-str_replace_all(datos2$AREA_TERRENO1, ",", ".")
datos2$AREA_TERRENO2<-str_replace_all(datos2$AREA_TERRENO2, ",", ".")
datos2[, "AREA_USO1"]<-str_replace_all(datos2[, "AREA_USO1"], ",", ".")
datos2[, "AREA_USO2"]<-str_replace_all(datos2[, "AREA_USO2"], ",", ".")

attach(datos2)

condicion=NULL
YEAR = 2022
#datos2$FORMULA[k]==""
#datos1$AREA_TERRENO2[!is.na(datos1$AREA_TERRENO2)]


for(k in 1:nrow(datos2)){
  if(TABLA_VALOR[k] %in% c("T01-1", "T01-2", "T01-3", "T01-4", "T01-5", "T01-6", "T01-7", "T01-8")){
    if(is.na(AREA_TERRENO2[k]) | AREA_TERRENO2[k] == "" | AREA_TERRENO2[k] == "<NA>"){ # | AREA_TERRENO2[k] == "0" | AREA_TERRENO2[k]==0
      condicion[k]=paste0("IF MODELO_USO = '" , TABLA_VALOR[k], "' AND (",
                          PUNTAJE1[k], "<= PUNTAJE <=", PUNTAJE2[k], ") AND (", EDAD.PREDIO1[k], "<= EDAD <=", EDAD.PREDIO2[k], ") AND (",
                          AREA.USO1[k], "<= AREA_USO <=", AREA.USO2[k], ") AND (",ESTRATO1[k], "<= CODIGO_ESTRATO <=",ESTRATO2[k], 
                          ") THEN VALOR_UNITARIO_", YEAR, " = ", VAL_METRO_CUAD[k], ";")
    }else{
      condicion[k]=paste0("IF MODELO_USO = '" , TABLA_VALOR[k], "' AND (",
                          PUNTAJE1[k], "<= PUNTAJE <=", PUNTAJE2[k], ") AND (", EDAD.PREDIO1[k], "<= EDAD <=", EDAD.PREDIO2[k], ") AND (",
                          AREA.USO1[k], "<= AREA_USO <=", AREA.USO2[k], ") AND (",ESTRATO1[k], "<= CODIGO_ESTRATO <=",ESTRATO2[k], ") AND (", 
                          AREA_TERRENO1[k], "<= AREA_TERRENO <=", AREA_TERRENO2[k],
                          ") THEN VALOR_UNITARIO_", YEAR, " = ", VAL_METRO_CUAD[k], ";")
      
    }
    }else{
  if(is.na(AREA_TERRENO2[k]) | AREA_TERRENO2[k] == "" | AREA_TERRENO2[k] == "<NA>"){ # | AREA_TERRENO2[k] == "0" | AREA_TERRENO2[k]==0
  
    if( !(TABLA_VALOR[k] %in% c('T21','T22'))   ){
    condicion[k]=paste0("IF MODELO_USO = '" , TABLA_VALOR[k], "' AND (",TIPO_CARACT_RES1[k], "<=MARCA_MANZANA<=",TIPO_CARACT_RES2[k], ") AND (",
                     PUNTAJE1[k], "<= PUNTAJE <=", PUNTAJE2[k], ") AND (", EDAD.PREDIO1[k], "<= EDAD <=", EDAD.PREDIO2[k], ") AND (",
                     AREA.USO1[k], "<= AREA_USO <=", AREA.USO2[k], ") AND (",ESTRATO1[k], "<= CODIGO_ESTRATO <=",ESTRATO2[k], 
                     ") THEN VALOR_UNITARIO_", YEAR, " = ", VAL_METRO_CUAD[k], ";")
  }else{
    if( datos2$FORMULA[k] !=""     ){
      condicion[k]=paste0("IF MODELO_USO = '" , TABLA_VALOR[k], "' AND (",TIPO_CARACT_RES1[k], "<=MARCA_MANZANA<=",TIPO_CARACT_RES2[k], ") AND (",
                          PUNTAJE1[k], "<= PUNTAJE <=", PUNTAJE2[k], ") AND (", EDAD.PREDIO1[k], "<= EDAD <=", EDAD.PREDIO2[k], ") AND (",
                          AREA.USO1[k], "<=AREA_CONSTRUIDA<=", AREA.USO2[k], ") AND (",ESTRATO1[k], "<= CODIGO_ESTRATO <=",ESTRATO2[k], 
                          ") THEN PRODUCTO_", YEAR, " = ",FORMULA[k], ";")
    }else{
      condicion[k]=paste0("IF MODELO_USO = '" , TABLA_VALOR[k], "' AND (",TIPO_CARACT_RES1[k], "<=MARCA_MANZANA<=",TIPO_CARACT_RES2[k], ") AND (",
                          PUNTAJE1[k], "<= PUNTAJE <=", PUNTAJE2[k], ") AND (", EDAD.PREDIO1[k], "<= EDAD <=", EDAD.PREDIO2[k], ") AND (",
                          AREA.USO1[k], "<=AREA_CONSTRUIDA<=", AREA.USO2[k], ") AND (",ESTRATO1[k], "<= CODIGO_ESTRATO <=",ESTRATO2[k], 
                          ") THEN PRODUCTO_", YEAR, " = ", VAL_METRO_CUAD[k], ";")
    }
  }
  
    
    }else{
   
    if( !(TABLA_VALOR[k] %in% c('T21','T22'))   ){
      condicion[k]=paste0("IF MODELO_USO = '" , TABLA_VALOR[k], "' AND (",TIPO_CARACT_RES1[k], "<=MARCA_MANZANA<=",TIPO_CARACT_RES2[k], ") AND (",
                          PUNTAJE1[k], "<= PUNTAJE <=", PUNTAJE2[k], ") AND (", EDAD.PREDIO1[k], "<= EDAD <=", EDAD.PREDIO2[k], ") AND (",
                          AREA.USO1[k], "<= AREA_USO <=", AREA.USO2[k], ") AND (",ESTRATO1[k], "<= CODIGO_ESTRATO <=",ESTRATO2[k], ") AND (", 
                          AREA_TERRENO1[k], "<= AREA_TERRENO <=", AREA_TERRENO2[k],
                          ") THEN VALOR_UNITARIO_", YEAR, " = ", VAL_METRO_CUAD[k], ";")
    }else{
      if( datos2$FORMULA[k] !=""     ){
        condicion[k]=paste0("IF MODELO_USO = '" , TABLA_VALOR[k], "' AND (",TIPO_CARACT_RES1[k], "<=MARCA_MANZANA<=",TIPO_CARACT_RES2[k], ") AND (",
                            PUNTAJE1[k], "<= PUNTAJE <=", PUNTAJE2[k], ") AND (", EDAD.PREDIO1[k], "<= EDAD <=", EDAD.PREDIO2[k], ") AND (",
                            AREA.USO1[k], "<=AREA_CONSTRUIDA<=", AREA.USO2[k], ") AND (",ESTRATO1[k], "<= CODIGO_ESTRATO <=",ESTRATO2[k], ") AND (",
                            AREA_TERRENO1[k], "<= AREA_TERRENO <=", AREA_TERRENO2[k],
                            ") THEN PRODUCTO_", YEAR, " = ",FORMULA[k], ";")
      }else{
        condicion[k]=paste0("IF MODELO_USO = '" , TABLA_VALOR[k], "' AND (",TIPO_CARACT_RES1[k], "<=MARCA_MANZANA<=",TIPO_CARACT_RES2[k], ") AND (",
                            PUNTAJE1[k], "<= PUNTAJE <=", PUNTAJE2[k], ") AND (", EDAD.PREDIO1[k], "<= EDAD <=", EDAD.PREDIO2[k], ") AND (",
                            AREA.USO1[k], "<=AREA_CONSTRUIDA<=", AREA.USO2[k], ") AND (",ESTRATO1[k], "<= CODIGO_ESTRATO <=",ESTRATO2[k], ") AND (",
                            AREA_TERRENO1[k], "<= AREA_TERRENO <=", AREA_TERRENO2[k],
                            ") THEN PRODUCTO_", YEAR, " = ", VAL_METRO_CUAD[k], ";")
      }
    }
  }
  }
}



condicion_1<-chartr("[]", "()", condicion)
#condicion_1<-unique(condicion_1)


for(m in grep("T16",condicion_1)){
condicion_1[m]<-gsub("IF MODELO_USO = 'T16'",paste0("IF MODELO_USO = 'T16' AND CODIGO_USO IN ('",paste(paste0("0",CODIGO_USO1[m]:CODIGO_USO2[m]),collapse="','"), "')"),condicion_1[m])
}

condicion_1<-unique(condicion_1)

write.xlsx(condicion_1, "Condiciones_ANDREA2.xlsx",sheetName ="Condiciones" ,row.names=F,append=F)
#write.xlsx(condicion_uso_1, "Condiciones_ANDREA1.xlsx",sheetName ="Condiciones_Usos_multiples" ,row.names=F,append=T)

