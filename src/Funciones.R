require(pacman)

p_load(tidyverse, openxlsx, janitor, data.table, dtplyr)

convert_to_numeric <- function(.x){
  as.numeric(as.character(.x))
}



uso_row <- function(row){
  if(row$CODIGO_USO1 == row$CODIGO_USO2){
    row_1 <- row %>% 
      mutate(CODIGO_USO = CODIGO_USO1) %>% 
      dplyr::select(tabla, EC_MO_ID, CODIGO_USO,
                    EDAD_PREDIO1, PUNTAJE1, AREA_USO1,
                    AREA_USO2, VAL_METRO_CUAD, AREA_USO,
                    TIPO_CARACT_RES)
  }else{
    row_1 <- row %>% 
      pivot_longer(names_to = "USO_PREVIO", values_to = "CODIGO_USO", cols = c("CODIGO_USO1", "CODIGO_USO2")) %>% 
      as.data.frame() %>% 
      dplyr::select(tabla, EC_MO_ID, CODIGO_USO,
                    EDAD_PREDIO1, PUNTAJE1, AREA_USO1,
                    AREA_USO2, VAL_METRO_CUAD, AREA_USO,
                    TIPO_CARACT_RES)
  }
  return(row_1)
}
marca_manzana_row <- function(row){
  #row <- t2
  #row <- t(row) %>% as.data.frame()
  if(row$TIPO_CARACT_RES2 == 8 & row$TIPO_CARACT_RES1 %in% c(0, 1)){
    row_1 <- row  %>% mutate(TIPO_CARACT_RES = NA) %>% 
      dplyr::select(tabla, EC_MO_ID, CODIGO_USO1, CODIGO_USO2,
                    EDAD_PREDIO1, PUNTAJE1, AREA_USO1,
                    AREA_USO2, VAL_METRO_CUAD, AREA_USO,
                    TIPO_CARACT_RES)
  }else if(row$TIPO_CARACT_RES1 == row$TIPO_CARACT_RES2){
    row_1 <- row  %>% mutate(TIPO_CARACT_RES = TIPO_CARACT_RES1) %>% 
      dplyr::select(tabla, EC_MO_ID, CODIGO_USO1, CODIGO_USO2,
                    EDAD_PREDIO1, PUNTAJE1, AREA_USO1,
                    AREA_USO2, VAL_METRO_CUAD, AREA_USO,
                    TIPO_CARACT_RES)
  }else if((as.numeric(as.character(row$TIPO_CARACT_RES2)) == (as.numeric(as.character(row$TIPO_CARACT_RES1)) + 1))){
    row_1 <- row %>% 
      pivot_longer(names_to = "TIPO_CARACT_PREVIA", values_to = "TIPO_CARACT_RES", cols = c("TIPO_CARACT_RES1", "TIPO_CARACT_RES2")) %>% 
      as.data.frame() %>%
      dplyr::select(tabla, EC_MO_ID, CODIGO_USO1, CODIGO_USO2,
                    EDAD_PREDIO1, PUNTAJE1, AREA_USO1,
                    AREA_USO2, VAL_METRO_CUAD, AREA_USO,
                    TIPO_CARACT_RES)
  }else{
    TIPO_CARACT_RES <- as.numeric(as.character(row$TIPO_CARACT_RES1)):as.numeric(as.character(row$TIPO_CARACT_RES2))
    row_1 <- cbind(row, TIPO_CARACT_RES) %>%
      as.data.frame() %>%
      dplyr::select(tabla, EC_MO_ID, CODIGO_USO1, CODIGO_USO2,
                    EDAD_PREDIO1, PUNTAJE1, AREA_USO1,
                    AREA_USO2, VAL_METRO_CUAD, AREA_USO,
                    TIPO_CARACT_RES)
  }
  return(row_1)
}


make_txt_file <- function(excel_name, tabla_selected, AREA_USO1 = NA, AREA_USO2 = NA, tabla_correlativa){
  if(str_detect(tabla_selected, "^T01")){
    Tabla_valores <- read.xlsx(excel_name, tabla_selected)
    codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
    
    Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
    Tabla_valores[1, 1] <- "EDAD_PREDIO1"
    colnames(Tabla_valores) <- Tabla_valores[1, ]
    
    Tabla_valores <- Tabla_valores[-1, ]
    
    
    Tabla_final <- Tabla_valores %>% 
      mutate_if(is.character, convert_to_numeric) %>% 
      pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                    values_to = "VAL_METRO_CUAD") %>% 
      mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
             EDAD_PREDIO2 = EDAD_PREDIO1,
             PUNTAJE2 = PUNTAJE1,
             VIGENCIA = 2022,
             EC_MO_ID = codigo_produccion_sel$codigo_produccion,
             CODIGO_USO1	= "",
             CODIGO_USO2	= "",
             TIPO_CARACT_RES1 = 0,
             TIPO_CARACT_RES2 = 8,
             AREA_USO1 = "0",
             AREA_USO2 = "1000000",
             ESTRATO1 = 0,
             ESTRATO2 = 6,
             FORMULA = "",
             AREA_TERRENO1 = "",
             AREA_TERRENO2 = "", 
             OBS = "") %>% 
      dplyr::select(VIGENCIA,
                    EC_MO_ID,	
                    CODIGO_USO1,	
                    CODIGO_USO2,
                    TIPO_CARACT_RES1,	
                    TIPO_CARACT_RES2,
                    PUNTAJE1,	
                    PUNTAJE2,	
                    EDAD_PREDIO1,
                    EDAD_PREDIO2,
                    AREA_USO1,
                    AREA_USO2,
                    ESTRATO1,
                    ESTRATO2,
                    VAL_METRO_CUAD,
                    FORMULA,
                    AREA_TERRENO1,
                    AREA_TERRENO2, 
                    OBS)
    
  }else if(str_detect(tabla_selected, "^T02")){
    if(tabla_selected == "T02-1"){
      Tabla_valores <- read.xlsx(excel_name, tabla_selected)
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      
      Tabla_valores <- Tabla_valores[-1, ]
      
      
      Tabla_1 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "3",
               CODIGO_USO2	= "4",
               TIPO_CARACT_RES1 = 1,
               TIPO_CARACT_RES2 = 3,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
      Tabla_2 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "56",
               CODIGO_USO2	= "56",
               TIPO_CARACT_RES1 = 1,
               TIPO_CARACT_RES2 = 3,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_final <- Tabla_1 %>% rbind(Tabla_2) %>% as.data.frame()
      
    }else if(tabla_selected == "T02-2"){
      Tabla_valores <- read.xlsx(excel_name, tabla_selected)
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      
      Tabla_valores <- Tabla_valores[-1, ]
      
      
      Tabla_1 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "3",
               CODIGO_USO2	= "4",
               TIPO_CARACT_RES1 = 4,
               TIPO_CARACT_RES2 = 5,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
      Tabla_2 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "3",
               CODIGO_USO2	= "4",
               TIPO_CARACT_RES1 = 8,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
      Tabla_3 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "56",
               CODIGO_USO2	= "56",
               TIPO_CARACT_RES1 = 4,
               TIPO_CARACT_RES2 = 5,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_final <- Tabla_1 %>% rbind(Tabla_2) %>% 
        rbind(Tabla_3) %>% as.data.frame()
      
    }else if(tabla_selected == "T02-3"){
      Tabla_valores <- read.xlsx(excel_name, tabla_selected)
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      
      Tabla_valores <- Tabla_valores[-1, ]
      
      
      Tabla_1 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "3",
               CODIGO_USO2	= "4",
               TIPO_CARACT_RES1 = 6,
               TIPO_CARACT_RES2 = 6,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
      Tabla_2 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "56",
               CODIGO_USO2	= "56",
               TIPO_CARACT_RES1 = 6,
               TIPO_CARACT_RES2 = 6,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_final <- Tabla_1 %>% rbind(Tabla_2) %>% as.data.frame()
      
    }else if(tabla_selected == "T02-4"){
      Tabla_valores <- read.xlsx(excel_name, tabla_selected)
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      
      Tabla_valores <- Tabla_valores[-1, ]
      
      
      Tabla_1 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "3",
               CODIGO_USO2	= "4",
               TIPO_CARACT_RES1 = 7,
               TIPO_CARACT_RES2 = 7,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
      Tabla_2 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "56",
               CODIGO_USO2	= "56",
               TIPO_CARACT_RES1 = 7,
               TIPO_CARACT_RES2 = 7,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_final <- Tabla_1 %>% rbind(Tabla_2) %>% as.data.frame()
      
      
    }else{
      stop("Las tablas T02 solo son T02-1, T02-2, T02-3 y T02-4.")
    }
    
    
  }else if(str_detect(tabla_selected, "^T03")){
    
    if(AREA_USO1 == "0" & AREA_USO2 == "350"){
      
      Tabla_valores <- read.xlsx(excel_name, "0-350")
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      Tabla_valores <- Tabla_valores[-1, ]
      
      
      Tabla_1 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "10",
               CODIGO_USO2	= "11",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "0",
               AREA_USO2 = "350",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_2 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "25",
               CODIGO_USO2	= "25",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "0",
               AREA_USO2 = "350",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_3 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "80",
               CODIGO_USO2	= "80",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "0",
               AREA_USO2 = "350",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
    }else if(AREA_USO1 == "350,01" & AREA_USO2 == "1000"){
      
      Tabla_valores <- read.xlsx(excel_name, "350,01-1000")
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      Tabla_valores <- Tabla_valores[-1, ]
      
      Tabla_1 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "10",
               CODIGO_USO2	= "11",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "350,01",
               AREA_USO2 = "1000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_2 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "25",
               CODIGO_USO2	= "25",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "350,01",
               AREA_USO2 = "1000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_3 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "80",
               CODIGO_USO2	= "80",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "350,01",
               AREA_USO2 = "1000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
    }else if(AREA_USO1 == "1000,01" & AREA_USO2 == "6000"){
      
      Tabla_valores <- read.xlsx(excel_name, "1000,01-6000")
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      Tabla_valores <- Tabla_valores[-1, ]
      
      Tabla_1 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "10",
               CODIGO_USO2	= "11",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "1000,01",
               AREA_USO2 = "6000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_2 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "25",
               CODIGO_USO2	= "25",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "1000,01",
               AREA_USO2 = "6000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_3 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "80",
               CODIGO_USO2	= "80",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "1000,01",
               AREA_USO2 = "6000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
    }else if(AREA_USO1 == "6000,01" & AREA_USO2 == "1000000"){
      
      
      Tabla_valores <- read.xlsx(excel_name, "6000,01-1000000")
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      Tabla_valores <- Tabla_valores[-1, ]
      
      Tabla_1 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "10",
               CODIGO_USO2	= "11",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "6000,01",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_2 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "25",
               CODIGO_USO2	= "25",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "6000,01",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_3 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "80",
               CODIGO_USO2	= "80",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "6000,01",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
    }else{
      stop("La tabla T03 no considera ese rango de valores de áreas de uso.")
    }
    Tabla_final <- Tabla_1 %>% rbind(Tabla_2) %>% 
      rbind(Tabla_3) %>% as.data.frame()
    
  }else if(str_detect(tabla_selected, "^T04")){
    Tabla_valores <- read.xlsx(excel_name, "T04")
    codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
    
    Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
    Tabla_valores[1, 1] <- "EDAD_PREDIO1"
    colnames(Tabla_valores) <- Tabla_valores[1, ]
    
    Tabla_valores <- Tabla_valores[-1, ]
    
    
    Tabla_final <- Tabla_valores %>% 
      mutate_if(is.character, convert_to_numeric) %>% 
      pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                    values_to = "VAL_METRO_CUAD") %>% 
      mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
             EDAD_PREDIO2 = EDAD_PREDIO1,
             PUNTAJE2 = PUNTAJE1,
             VIGENCIA = 2022,
             EC_MO_ID = codigo_produccion_sel$codigo_produccion,
             CODIGO_USO1	= "33",
             CODIGO_USO2	= "33",
             TIPO_CARACT_RES1 = 0,
             TIPO_CARACT_RES2 = 8,
             AREA_USO1 = "0",
             AREA_USO2 = "1000000",
             ESTRATO1 = 0,
             ESTRATO2 = 6,
             FORMULA = "",
             AREA_TERRENO1 = "",
             AREA_TERRENO2 = "", 
             OBS = "") %>% 
      dplyr::select(VIGENCIA,
                    EC_MO_ID,	
                    CODIGO_USO1,	
                    CODIGO_USO2,
                    TIPO_CARACT_RES1,	
                    TIPO_CARACT_RES2,
                    PUNTAJE1,	
                    PUNTAJE2,	
                    EDAD_PREDIO1,
                    EDAD_PREDIO2,
                    AREA_USO1,
                    AREA_USO2,
                    ESTRATO1,
                    ESTRATO2,
                    VAL_METRO_CUAD,
                    FORMULA,
                    AREA_TERRENO1,
                    AREA_TERRENO2, 
                    OBS)
    
  }else if(str_detect(tabla_selected, "^T05")){
    if(tabla_selected == "T05-1"){
      Tabla_valores <- read.xlsx(excel_name, tabla_selected)
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      
      Tabla_valores <- Tabla_valores[-1, ]
      
      
      Tabla_1 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "15",
               CODIGO_USO2	= "15",
               TIPO_CARACT_RES1 = 1,
               TIPO_CARACT_RES2 = 3,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
      Tabla_2 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "20",
               CODIGO_USO2	= "20",
               TIPO_CARACT_RES1 = 1,
               TIPO_CARACT_RES2 = 3,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_3 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "58",
               CODIGO_USO2	= "58",
               TIPO_CARACT_RES1 = 1,
               TIPO_CARACT_RES2 = 3,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_final <- Tabla_1 %>% rbind(Tabla_2) %>%
        rbind(Tabla_3) %>% as.data.frame()
      
    }else if(tabla_selected == "T05-2"){
      Tabla_valores <- read.xlsx(excel_name, tabla_selected)
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      
      Tabla_valores <- Tabla_valores[-1, ]
      
      
      Tabla_1 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "15",
               CODIGO_USO2	= "15",
               TIPO_CARACT_RES1 = 4,
               TIPO_CARACT_RES2 = 5,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
      Tabla_2 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "20",
               CODIGO_USO2	= "20",
               TIPO_CARACT_RES1 = 4,
               TIPO_CARACT_RES2 = 5,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_3 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "58",
               CODIGO_USO2	= "58",
               TIPO_CARACT_RES1 = 4,
               TIPO_CARACT_RES2 = 5,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
      
      Tabla_4 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "15",
               CODIGO_USO2	= "15",
               TIPO_CARACT_RES1 = 8,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
      Tabla_5 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "20",
               CODIGO_USO2	= "20",
               TIPO_CARACT_RES1 = 8,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_6 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "58",
               CODIGO_USO2	= "58",
               TIPO_CARACT_RES1 = 8,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
      Tabla_final <- Tabla_1 %>% rbind(Tabla_2) %>%
        rbind(Tabla_3) %>% rbind(Tabla_4) %>%
        rbind(Tabla_5) %>%  rbind(Tabla_6) %>% as.data.frame()
      
    }else if(tabla_selected == "T05-3"){
      Tabla_valores <- read.xlsx(excel_name, tabla_selected)
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      
      Tabla_valores <- Tabla_valores[-1, ]
      
      
      Tabla_1 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "15",
               CODIGO_USO2	= "15",
               TIPO_CARACT_RES1 = 6,
               TIPO_CARACT_RES2 = 6,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
      Tabla_2 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "20",
               CODIGO_USO2	= "20",
               TIPO_CARACT_RES1 = 6,
               TIPO_CARACT_RES2 = 6,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_3 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "58",
               CODIGO_USO2	= "58",
               TIPO_CARACT_RES1 = 6,
               TIPO_CARACT_RES2 = 6,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_final <- Tabla_1 %>% rbind(Tabla_2) %>%
        rbind(Tabla_3) %>% as.data.frame()
      
      
    }else if(tabla_selected == "T05-4"){
      Tabla_valores <- read.xlsx(excel_name, tabla_selected)
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      
      Tabla_valores <- Tabla_valores[-1, ]
      
      
      Tabla_1 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "15",
               CODIGO_USO2	= "15",
               TIPO_CARACT_RES1 = 7,
               TIPO_CARACT_RES2 = 7,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
      Tabla_2 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "20",
               CODIGO_USO2	= "20",
               TIPO_CARACT_RES1 = 7,
               TIPO_CARACT_RES2 = 7,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_3 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "58",
               CODIGO_USO2	= "58",
               TIPO_CARACT_RES1 = 7,
               TIPO_CARACT_RES2 = 7,
               AREA_USO1 = "0",
               AREA_USO2 = "1000000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_final <- Tabla_1 %>% rbind(Tabla_2) %>%
        rbind(Tabla_3) %>% as.data.frame()
      
      
    }else{
      stop("Las tablas T05 solo son T05-1, T05-2, T05-3 y T05-4.")
    }
    
  }else if(str_detect(tabla_selected, "^T06")){
    
    if(AREA_USO1 == "350" & AREA_USO2 == "500"){
      
      Tabla_valores <- read.xlsx(excel_name, "350-500")
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      Tabla_valores <- Tabla_valores[-1, ]
      
      Tabla_final <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "12",
               CODIGO_USO2	= "12",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "350",
               AREA_USO2 = "500",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
    }else if(AREA_USO1 == "500,01" & AREA_USO2 == "750"){
      
      Tabla_valores <- read.xlsx(excel_name, "500,01-750")
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      Tabla_valores <- Tabla_valores[-1, ]
      
      Tabla_final <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "12",
               CODIGO_USO2	= "12",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "500,01",
               AREA_USO2 = "750",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
      
    }else if(AREA_USO1 == "750,01" & AREA_USO2 == "2000"){
      
      Tabla_valores <- read.xlsx(excel_name, "750,01-2000")
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      Tabla_valores <- Tabla_valores[-1, ]
      
      Tabla_final <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "12",
               CODIGO_USO2	= "12",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "750,01",
               AREA_USO2 = "2000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
    }else if(AREA_USO1 == "2000,01" & AREA_USO2 == "5000"){
      
      Tabla_valores <- read.xlsx(excel_name, "2000,01-5000")
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      Tabla_valores <- Tabla_valores[-1, ]
      
      
      Tabla_final <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "12",
               CODIGO_USO2	= "12",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "2000,01",
               AREA_USO2 = "5000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
    }else if(AREA_USO1 == "5000,01" & AREA_USO2 == "10000"){
      
      Tabla_valores <- read.xlsx(excel_name, "5000,01-10000")
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      Tabla_valores <- Tabla_valores[-1, ]
      
      
      Tabla_final <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "12",
               CODIGO_USO2	= "12",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "5000,01",
               AREA_USO2 = "10000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
    }else{
      stop("La tabla T03 no considera ese rango de valores de áreas de uso.")
    }
  }else if(str_detect(tabla_selected, "^T07")){
    if(AREA_USO1 == "0" & AREA_USO2 == "100"){
      
      Tabla_valores <- read.xlsx(excel_name, "0-100")
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      Tabla_valores <- Tabla_valores[-1, ]
      
      Tabla_final <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "22",
               CODIGO_USO2	= "22",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "0",
               AREA_USO2 = "100",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
    }else if(AREA_USO1 == "100,01" & AREA_USO2 == "350"){
      
      Tabla_valores <- read.xlsx(excel_name, "100,01-350")
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      Tabla_valores <- Tabla_valores[-1, ]
      
      
      Tabla_final <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "22",
               CODIGO_USO2	= "22",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "100,01",
               AREA_USO2 = "350",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
    }else if(AREA_USO1 == "350,01" & AREA_USO2 == "100000"){
      
      Tabla_valores <- read.xlsx(excel_name, "350,01-100000")
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      Tabla_valores <- Tabla_valores[-1, ]
      
      
      Tabla_final <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "22",
               CODIGO_USO2	= "22",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "350,01",
               AREA_USO2 = "100000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
    }else{
      stop("La tabla T07 no considera ese rango de valores de áreas de uso.")
    }
  }else if(str_detect(tabla_selected, "^T19")){
    if(AREA_USO1 == "0" & AREA_USO2 == "350"){
      Tabla_valores <- read.xlsx(excel_name, "0-350")
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      
      Tabla_valores <- Tabla_valores[-1, ]
      
      
      Tabla_1 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "8",
               CODIGO_USO2	= "8",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "0",
               AREA_USO2 = "350",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
      Tabla_2 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "14",
               CODIGO_USO2	= "14",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "0",
               AREA_USO2 = "350",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_3 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "23",
               CODIGO_USO2	= "23",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "0",
               AREA_USO2 = "350",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_final <- Tabla_1 %>% rbind(Tabla_2) %>%
        rbind(Tabla_3) %>% as.data.frame()
      
    }else if(AREA_USO1 == "350,01" & AREA_USO2 == "1000"){
      Tabla_valores <- read.xlsx(excel_name, "350,01-1000")
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      
      Tabla_valores <- Tabla_valores[-1, ]
      
      
      Tabla_1 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "8",
               CODIGO_USO2	= "8",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "350,01",
               AREA_USO2 = "1000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
      Tabla_2 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "14",
               CODIGO_USO2	= "14",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "350,01",
               AREA_USO2 = "1000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_3 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "23",
               CODIGO_USO2	= "23",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "350,01",
               AREA_USO2 = "1000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_final <- Tabla_1 %>% rbind(Tabla_2) %>%
        rbind(Tabla_3) %>% as.data.frame()
      
    }else if(AREA_USO1 == "1000,01" & AREA_USO2 == "6000"){
      Tabla_valores <- read.xlsx(excel_name, "1000,01-6000")
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      
      Tabla_valores <- Tabla_valores[-1, ]
      
      
      Tabla_1 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "8",
               CODIGO_USO2	= "8",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "1000,01",
               AREA_USO2 = "6000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
      Tabla_2 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "14",
               CODIGO_USO2	= "14",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "1000,01",
               AREA_USO2 = "6000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_3 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "23",
               CODIGO_USO2	= "23",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "1000,01",
               AREA_USO2 = "6000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_final <- Tabla_1 %>% rbind(Tabla_2) %>%
        rbind(Tabla_3) %>% as.data.frame()
      
    }else if(AREA_USO1 == "6000,01" & AREA_USO2 == "100000"){
      Tabla_valores <- read.xlsx(excel_name, "6000,01-100000")
      codigo_produccion_sel <- tabla_correlativa[tabla == tabla_selected, "codigo_produccion"]
      
      Tabla_valores[1, ] <- as.character(paste0("V", Tabla_valores[1, ]))
      Tabla_valores[1, 1] <- "EDAD_PREDIO1"
      colnames(Tabla_valores) <- Tabla_valores[1, ]
      
      Tabla_valores <- Tabla_valores[-1, ]
      
      
      Tabla_1 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "8",
               CODIGO_USO2	= "8",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "6000,01",
               AREA_USO2 = "100000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      
      Tabla_2 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "14",
               CODIGO_USO2	= "14",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "6000,01",
               AREA_USO2 = "100000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_3 <- Tabla_valores %>% 
        mutate_if(is.character, convert_to_numeric) %>% 
        pivot_longer( names_to = "PUNTAJE1", cols = 2:102,
                      values_to = "VAL_METRO_CUAD") %>% 
        mutate(PUNTAJE1 = as.numeric(as.character(str_replace_all(PUNTAJE1, "V", ""))),
               EDAD_PREDIO2 = EDAD_PREDIO1,
               PUNTAJE2 = PUNTAJE1,
               VIGENCIA = 2022,
               EC_MO_ID = codigo_produccion_sel$codigo_produccion,
               CODIGO_USO1	= "23",
               CODIGO_USO2	= "23",
               TIPO_CARACT_RES1 = 0,
               TIPO_CARACT_RES2 = 8,
               AREA_USO1 = "6000,01",
               AREA_USO2 = "100000",
               ESTRATO1 = 0,
               ESTRATO2 = 6,
               FORMULA = "",
               AREA_TERRENO1 = "",
               AREA_TERRENO2 = "", 
               OBS = "") %>% 
        dplyr::select(VIGENCIA,
                      EC_MO_ID,	
                      CODIGO_USO1,	
                      CODIGO_USO2,
                      TIPO_CARACT_RES1,	
                      TIPO_CARACT_RES2,
                      PUNTAJE1,	
                      PUNTAJE2,	
                      EDAD_PREDIO1,
                      EDAD_PREDIO2,
                      AREA_USO1,
                      AREA_USO2,
                      ESTRATO1,
                      ESTRATO2,
                      VAL_METRO_CUAD,
                      FORMULA,
                      AREA_TERRENO1,
                      AREA_TERRENO2, 
                      OBS)
      Tabla_final <- Tabla_1 %>% rbind(Tabla_2) %>%
        rbind(Tabla_3) %>% as.data.frame()
      
    }
    
  }else{
    stop("La tabla no es considerada.")
  }
  return(Tabla_final)
}




validate_values_tables <- function(df_table, tabla_correlativa){
  #df_table <- make_txt_file(excel_name = "input/T01.xlsx", tabla_selected = "T01-3", AREA_USO1 = NA, AREA_USO2 = NA)
  
  table <- tabla_correlativa[tabla_correlativa$codigo_produccion == unique(df_table$EC_MO_ID), 
                             "tabla"]$tabla
  if(str_detect(table, "^T01")){
    if(any(df_table$CODIGO_USO1 != "" | df_table$CODIGO_USO2 != "")){
      stop("Error en el código del uso.")
    }
    if(any(df_table$TIPO_CARACT_RES1 != 0 | df_table$TIPO_CARACT_RES2 != 8)){
      stop("Tipo de característica residencial erroneo.")
    }
    
    if((min(df_table$EDAD_PREDIO1) != 0 | max(df_table$EDAD_PREDIO1) != 100) |
       (min(df_table$EDAD_PREDIO2) != 0 | max(df_table$EDAD_PREDIO2) != 100) |
       (min(df_table$PUNTAJE1) != 0 | max(df_table$PUNTAJE1) != 100) |
       (min(df_table$PUNTAJE2) != 0 | max(df_table$PUNTAJE2) != 100)){
      stop("Edades o puntajes incorrectos.")
    }
    if(any(df_table$AREA_USO1 != "0" | df_table$AREA_USO2 != "1000000")){
      stop("Área uso incorrecta.")
    }
    
    if(any(df_table$ESTRATO1 != 0 | df_table$ESTRATO2 != 6)){
      stop("Estrato incorrecto.")
    }
    
    
  }else if(str_detect(table, "^T02")){
    if(!identical(unique(df_table$CODIGO_USO1), c("3", "56")) |
       !identical(unique(df_table$CODIGO_USO2), c("4", "56"))){
      stop("El código uso está erroneo.")
    }
    if((min(df_table$EDAD_PREDIO1) != 0 | max(df_table$EDAD_PREDIO1) != 100) |
       (min(df_table$EDAD_PREDIO2) != 0 | max(df_table$EDAD_PREDIO2) != 100) |
       (min(df_table$PUNTAJE1) != 0 | max(df_table$PUNTAJE1) != 100) |
       (min(df_table$PUNTAJE2) != 0 | max(df_table$PUNTAJE2) != 100)){
      stop("Edades o puntajes incorrectos.")
    }
    if(any(df_table$ESTRATO1 != 0 | df_table$ESTRATO2 != 6)){
      stop("Estrato incorrecto.")
    }
    
    if(any(df_table$AREA_USO1 != "0" | df_table$AREA_USO2 != "1000000")){
      stop("Área uso incorrecta.")
    }
    
    if(table == "T02-1"){
      
      if(any(df_table$TIPO_CARACT_RES1 != 1 | df_table$TIPO_CARACT_RES2 != 3)){
        stop("Tipo de característica residencial erroneo.")
      }
      
    }else if(table == "T02-2"){
      if(!identical(unique(df_table$TIPO_CARACT_RES1), c(4, 8)) |
         !identical(unique(df_table$TIPO_CARACT_RES2), c(5, 8))){
        stop("Tipo de característica residencial erroneo.")
      }
      
    }else if(table == "T02-3"){
      
      if(!identical(unique(df_table$TIPO_CARACT_RES1), c(6)) |
         !identical(unique(df_table$TIPO_CARACT_RES2), c(6))){
        stop("Tipo de característica residencial erroneo.")
      }
      
    }else if(table == "T02-4"){
      if(!identical(unique(df_table$TIPO_CARACT_RES1), c(7)) |
         !identical(unique(df_table$TIPO_CARACT_RES2), c(7))){
        stop("Tipo de característica residencial erroneo.")
      }
      
    }else{
      stop("La tabla no está asignada correctamente.")
    }
    
  }else if(str_detect(table, "^T03")){
    if(!identical(unique(df_table$TIPO_CARACT_RES1), c(0)) |
       !identical(unique(df_table$TIPO_CARACT_RES2), c(8))){
      stop("Tipo de característica residencial erroneo.")
    }
    if((min(df_table$EDAD_PREDIO1) != 0 | max(df_table$EDAD_PREDIO1) != 100) |
       (min(df_table$EDAD_PREDIO2) != 0 | max(df_table$EDAD_PREDIO2) != 100) |
       (min(df_table$PUNTAJE1) != 0 | max(df_table$PUNTAJE1) != 100) |
       (min(df_table$PUNTAJE2) != 0 | max(df_table$PUNTAJE2) != 100)){
      stop("Edades o puntajes incorrectos.")
    }
    
    if(any(df_table$ESTRATO1 != 0 | df_table$ESTRATO2 != 6)){
      stop("Estrato incorrecto.")
    }
    
    if(!identical(unique(df_table$CODIGO_USO1), c("10", "25", "80")) |
       !identical(unique(df_table$CODIGO_USO2), c("11", "25", "80"))){
      stop("El código uso está erroneo.")
    }  
    if(unique(df_table$AREA_USO1) == "0" & unique(df_table$AREA_USO2) == "350"){
      
    }else if(unique(df_table$AREA_USO1) == "350,01" & unique(df_table$AREA_USO2) == "1000"){
      
    }else if(unique(df_table$AREA_USO1) == "1000,01" & unique(df_table$AREA_USO2) == "6000"){
      
    }else if(unique(df_table$AREA_USO1)  == "6000,01" & unique(df_table$AREA_USO2) == "1000000"){
      
    }else{
      stop("Las áreas asignadas no son correctas.")
    }
    
  }else if(str_detect(table, "^T04")){
    if(!identical(unique(df_table$CODIGO_USO1), c("33")) |
       !identical(unique(df_table$CODIGO_USO2), c("33"))){
      stop("El código uso está erroneo.")
    }  
    
    if(!identical(unique(df_table$TIPO_CARACT_RES1), c(0)) |
       !identical(unique(df_table$TIPO_CARACT_RES2), c(8))){
      stop("Tipo de característica residencial erroneo.")
    }
    if((min(df_table$EDAD_PREDIO1) != 0 | max(df_table$EDAD_PREDIO1) != 100) |
       (min(df_table$EDAD_PREDIO2) != 0 | max(df_table$EDAD_PREDIO2) != 100) |
       (min(df_table$PUNTAJE1) != 0 | max(df_table$PUNTAJE1) != 100) |
       (min(df_table$PUNTAJE2) != 0 | max(df_table$PUNTAJE2) != 100)){
      stop("Edades o puntajes incorrectos.")
    }
    if(any(df_table$ESTRATO1 != 0 | df_table$ESTRATO2 != 6)){
      stop("Estrato incorrecto.")
    }
    
  }else if(str_detect(table, "^T05")){
    if(!identical(unique(df_table$CODIGO_USO1), c("15", "20", "58")) |
       !identical(unique(df_table$CODIGO_USO2), c("15", "20", "58"))){
      stop("El código uso está erroneo.")
    }  
    
    if((min(df_table$EDAD_PREDIO1) != 0 | max(df_table$EDAD_PREDIO1) != 100) |
       (min(df_table$EDAD_PREDIO2) != 0 | max(df_table$EDAD_PREDIO2) != 100) |
       (min(df_table$PUNTAJE1) != 0 | max(df_table$PUNTAJE1) != 100) |
       (min(df_table$PUNTAJE2) != 0 | max(df_table$PUNTAJE2) != 100)){
      stop("Edades o puntajes incorrectos.")
    }
    if(any(df_table$ESTRATO1 != 0 | df_table$ESTRATO2 != 6)){
      stop("Estrato incorrecto.")
    }
    if(table == "T05-1"){
      if(!identical(unique(df_table$TIPO_CARACT_RES1), c(1)) |
         !identical(unique(df_table$TIPO_CARACT_RES2), c(3))){
        stop("Tipo de característica residencial erroneo.")
      }
      
    }else if(table == "T05-2"){
      if(!identical(unique(df_table$TIPO_CARACT_RES1), c(4, 8)) |
         !identical(unique(df_table$TIPO_CARACT_RES2), c(5, 8))){
        stop("Tipo de característica residencial erroneo.")
      }
    }else if(table == "T05-3"){
      if(!identical(unique(df_table$TIPO_CARACT_RES1), c(6)) |
         !identical(unique(df_table$TIPO_CARACT_RES2), c(6))){
        stop("Tipo de característica residencial erroneo.")
      }
    }else if(table == "T05-4"){
      if(!identical(unique(df_table$TIPO_CARACT_RES1), c(7)) |
         !identical(unique(df_table$TIPO_CARACT_RES2), c(7))){
        stop("Tipo de característica residencial erroneo.")
      }
      
    }else{
      stop("La tabla no está asignada correctamente.")
    }
  }else if(str_detect(table, "^T06")){
    if(!identical(unique(df_table$CODIGO_USO1), c("12")) |
       !identical(unique(df_table$CODIGO_USO2), c("12"))){
      stop("El código uso está erroneo.")
    }  
    
    if((min(df_table$EDAD_PREDIO1) != 0 | max(df_table$EDAD_PREDIO1) != 100) |
       (min(df_table$EDAD_PREDIO2) != 0 | max(df_table$EDAD_PREDIO2) != 100) |
       (min(df_table$PUNTAJE1) != 0 | max(df_table$PUNTAJE1) != 100) |
       (min(df_table$PUNTAJE2) != 0 | max(df_table$PUNTAJE2) != 100)){
      stop("Edades o puntajes incorrectos.")
    }
    if(any(df_table$ESTRATO1 != 0 | df_table$ESTRATO2 != 6)){
      stop("Estrato incorrecto.")
    }
    if(!identical(unique(df_table$TIPO_CARACT_RES1), c(0)) |
       !identical(unique(df_table$TIPO_CARACT_RES2), c(8))){
      stop("Tipo de característica residencial erroneo.")
    }
    
    if(unique(df_table$AREA_USO1) == "350" & unique(df_table$AREA_USO2) == "500"){
    }else if(unique(df_table$AREA_USO1) == "500,01" & unique(df_table$AREA_USO2) == "750"){
      
    }else if(unique(df_table$AREA_USO1) == "750,01" & unique(df_table$AREA_USO2) == "2000"){
      
    }else if(unique(df_table$AREA_USO1) == "2000,01" & unique(df_table$AREA_USO2) == "5000"){
      
    }else if(unique(df_table$AREA_USO1)  == "5000,01" & unique(df_table$AREA_USO2) == "10000"){
      
    }else{
      stop("Las áreas asignadas no son correctas.")
    }
    
  }else if(str_detect(table, "^T19")){
    if(!identical(unique(df_table$CODIGO_USO1), c("8", "14", "23")) |
       !identical(unique(df_table$CODIGO_USO2), c("8", "14", "23"))){
      stop("El código uso está erroneo.")
    }  
    if(!identical(unique(df_table$TIPO_CARACT_RES1), c(0)) |
       !identical(unique(df_table$TIPO_CARACT_RES2), c(8))){
      stop("Tipo de característica residencial erroneo.")
    }
    if((min(df_table$EDAD_PREDIO1) != 0 | max(df_table$EDAD_PREDIO1) != 100) |
       (min(df_table$EDAD_PREDIO2) != 0 | max(df_table$EDAD_PREDIO2) != 100) |
       (min(df_table$PUNTAJE1) != 0 | max(df_table$PUNTAJE1) != 100) |
       (min(df_table$PUNTAJE2) != 0 | max(df_table$PUNTAJE2) != 100)){
      stop("Edades o puntajes incorrectos.")
    }
    if(any(df_table$ESTRATO1 != 0 | df_table$ESTRATO2 != 6)){
      stop("Estrato incorrecto.")
    }
    if(unique(df_table$AREA_USO1) == "0" & unique(df_table$AREA_USO2) == "350"){
      
    }else if(unique(df_table$AREA_USO1) == "350,01" & unique(df_table$AREA_USO2) == "1000"){
      
    }else if(unique(df_table$AREA_USO1) == "1000,01" & unique(df_table$AREA_USO2) == "6000"){
      
    }else if(unique(df_table$AREA_USO1)  == "6000,01" & unique(df_table$AREA_USO2) == "100000"){
      
    }else{
      stop("Las áreas asignadas no son correctas.")
    }
    
  }else{
    stop("La tabla no está considerada dentro del grupo 101x101.")
  }
  
}
