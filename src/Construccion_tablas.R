rm(list = ls())
require(pacman)

p_load(tidyverse, openxlsx, janitor, data.table, dtplyr)

source("src/Funciones.R")

options(scipen = 999)
tabla_correlativa <- read.xlsx("input/CARGUE TABLAS VIG 2022.xlsx", "TABLAS NPH") %>% 
  clean_names() %>% filter(codigo_produccion != 2022) %>% as.data.table()
#tabla_correlativa <- tabla_correlativa %>% lazy_dt() %>% 
  

excel_name <- "input/T01-Residencial todos los rangos Vigencia 2022_ajustado8_entrega 21092021.xlsx"
nombre_export <- "_REPORTE_SIE_VR_UNITARIO_TABLAS_NPH"
nombre_export_est <- "_ARCHIVO_CARGUE_VR_UNIT_ESTADISTICA"

# make_txt_file(excel_name = "input/T01.xlsx", tabla_selected = "T01-1", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# make_txt_file(excel_name = "input/T01.xlsx", tabla_selected = "T01-2", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# make_txt_file(excel_name = "input/T01.xlsx", tabla_selected = "T01-3", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# make_txt_file(excel_name = "input/T01.xlsx", tabla_selected = "T01-4", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# make_txt_file(excel_name = "input/T01.xlsx", tabla_selected = "T01-5", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# make_txt_file(excel_name = "input/T01.xlsx", tabla_selected = "T01-6", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# make_txt_file(excel_name = "input/T01.xlsx", tabla_selected = "T01-7", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# make_txt_file(excel_name = "input/T01.xlsx", tabla_selected = "T01-8", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# 
# 
# make_txt_file(excel_name = "input/T02.xlsx", tabla_selected = "T02-1", tabla_correlativa = tabla_correlativa)  %>% validate_values_tables(., tabla_correlativa)
# make_txt_file(excel_name = "input/T02.xlsx", tabla_selected = "T02-2", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# make_txt_file(excel_name = "input/T02.xlsx", tabla_selected = "T02-3", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# make_txt_file(excel_name = "input/T02.xlsx", tabla_selected = "T02-4", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# 
# make_txt_file(excel_name = "input/T04.xlsx", tabla_selected = "T04", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# 
# make_txt_file(excel_name = "input/T05.xlsx", tabla_selected = "T05-1", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# make_txt_file(excel_name = "input/T05.xlsx", tabla_selected = "T05-2", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# make_txt_file(excel_name = "input/T05.xlsx", tabla_selected = "T05-3", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# make_txt_file(excel_name = "input/T05.xlsx", tabla_selected = "T05-4", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# 
# 
# make_txt_file(excel_name = "input/T03.xlsx", tabla_selected = "T03", AREA_USO1 = "0", AREA_USO2 = "350", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# make_txt_file(excel_name = "input/T03.xlsx", tabla_selected = "T03", AREA_USO1 = "350,01", AREA_USO2 = "1000", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# make_txt_file(excel_name = "input/T03.xlsx", tabla_selected = "T03", AREA_USO1 = "1000,01", AREA_USO2 = "6000", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# make_txt_file(excel_name = "input/T03.xlsx", tabla_selected = "T03", AREA_USO1 = "6000,01", AREA_USO2 = "1000000", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# 
# make_txt_file(excel_name = "input/T06.xlsx", tabla_selected = "T06", AREA_USO1 = "350", AREA_USO2 = "500", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# make_txt_file(excel_name = "input/T06.xlsx", tabla_selected = "T06", AREA_USO1 = "500,01", AREA_USO2 = "750", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# make_txt_file(excel_name = "input/T06.xlsx", tabla_selected = "T06", AREA_USO1 = "750,01", AREA_USO2 = "2000", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# make_txt_file(excel_name = "input/T06.xlsx", tabla_selected = "T06", AREA_USO1 = "2000,01", AREA_USO2 = "5000", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# #make_txt_file(excel_name = "input/T06.xlsx", tabla_selected = "T06", AREA_USO1 = "5000,01", AREA_USO2 = "10000")
# 
# 
# make_txt_file(excel_name = "input/T19.xlsx", tabla_selected = "T19", AREA_USO1 = "0", AREA_USO2 = "350", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# make_txt_file(excel_name = "input/T19.xlsx", tabla_selected = "T19", AREA_USO1 = "350,01", AREA_USO2 = "1000", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)
# make_txt_file(excel_name = "input/T19.xlsx", tabla_selected = "T19", AREA_USO1 = "1000,01", AREA_USO2 = "6000", tabla_correlativa = tabla_correlativa) %>% validate_values_tables(., tabla_correlativa)


# df_completo <- pmap_dfr(list(c(rep("input/T01.xlsx", 8), rep("input/T02.xlsx", 4), "input/T04.xlsx", rep("input/T05.xlsx", 4), rep("input/T03.xlsx", 4), rep("input/T06.xlsx", 4), rep("input/T19.xlsx", 3), rep("input/T07.xlsx", 3)),
#               c(paste0("T01-", 1:8), paste0("T02-", 1:4), "T04", paste0("T05-", 1:4), rep("T03", 4), rep("T06", 4), rep("T19", 3), rep("T07", 3)),
#               c(rep(NA, 17), , , "0", "350,01", "1000,01", "0", "100,01", "350,01"),
#               c(rep(NA, 17), , , "350", "1000", "6000", "100", "350", "100000")),
#               ~make_txt_file(excel_name = ..1, tabla_selected = ..2, AREA_USO1 = ..3, AREA_USO2 = ..4, tabla_correlativa = tabla_correlativa))

df_completo_t01 <- pmap_dfr(list(rep("input/T01.xlsx", 8),
                              paste0("T01-", 1:8),
                              rep(NA, 8),
                              rep(NA, 8)),
                         ~make_txt_file(excel_name = ..1, tabla_selected = ..2, AREA_USO1 = ..3, AREA_USO2 = ..4, tabla_correlativa = tabla_correlativa))

df_completo_t02 <- pmap_dfr(list(rep("input/T02.xlsx", 4),
                                 paste0("T02-", 1:4),
                                 rep(NA, 4),
                                 rep(NA, 4)),
                            ~make_txt_file(excel_name = ..1, tabla_selected = ..2, AREA_USO1 = ..3, AREA_USO2 = ..4, tabla_correlativa = tabla_correlativa))

df_completo_t03 <- pmap_dfr(list(rep("input/T03.xlsx", 4),
                                 rep("T03", 4),
                                 c("0", "350,01", "1000,01", "6000,01"),
                                 c("350", "1000", "6000", "1000000")),
                            ~make_txt_file(excel_name = ..1, tabla_selected = ..2, AREA_USO1 = ..3, AREA_USO2 = ..4, tabla_correlativa = tabla_correlativa))

df_completo_t04 <- pmap_dfr(list(rep("input/T04.xlsx", 1),
                                 paste0("T04"),
                                 rep(NA, 1),
                                 rep(NA, 1)),
                            ~make_txt_file(excel_name = ..1, tabla_selected = ..2, AREA_USO1 = ..3, AREA_USO2 = ..4, tabla_correlativa = tabla_correlativa))

df_completo_t05 <- pmap_dfr(list(rep("input/T05.xlsx", 4),
                                 paste0("T05-", 1:4),
                                 rep(NA, 4),
                                 rep(NA, 4)),
                            ~make_txt_file(excel_name = ..1, tabla_selected = ..2, AREA_USO1 = ..3, AREA_USO2 = ..4, tabla_correlativa = tabla_correlativa))

df_completo_t06 <- pmap_dfr(list(rep("input/T06.xlsx", 5),
                                 rep("T06", 5),
                                 c("350", "500,01", "750,01", "2000,01", "5000,01"),
                                 c("500", "750", "2000", "5000", "10000")),
                            ~make_txt_file(excel_name = ..1, tabla_selected = ..2, AREA_USO1 = ..3, AREA_USO2 = ..4, tabla_correlativa = tabla_correlativa))

df_completo_t07 <- pmap_dfr(list(rep("input/T07.xlsx", 3),
                                 rep("T07", 3),
                                 c("0", "100,01", "350,01"),
                                 c("100", "350", "100000")),
                            ~make_txt_file(excel_name = ..1, tabla_selected = ..2, AREA_USO1 = ..3, AREA_USO2 = ..4, tabla_correlativa = tabla_correlativa))

df_completo_t19 <- pmap_dfr(list(rep("input/T19.xlsx", 4),
                                 rep("T19", 4),
                                 c("0", "350,01", "1000,01", "6000,01"),
                                 c("350", "1000", "6000", "100000")),
                            ~make_txt_file(excel_name = ..1, tabla_selected = ..2, AREA_USO1 = ..3, AREA_USO2 = ..4, tabla_correlativa = tabla_correlativa))


df_completo <- rbind(df_completo_t01,
df_completo_t02,
df_completo_t03,
df_completo_t04,
df_completo_t05,
df_completo_t06,
df_completo_t07,
df_completo_t19) %>% as.data.frame

df_completo <- df_completo %>% mutate(AREA_TERRENO1 = case_when(EC_MO_ID == 389 ~ "49",
                                                  TRUE ~ AREA_TERRENO1),
                       AREA_TERRENO2 = case_when(EC_MO_ID == 389 ~ "100000000",
                                                  TRUE ~ AREA_TERRENO2))

df_t01_2_terreno <- df_completo %>% 
  filter(EC_MO_ID == 390) %>% 
  mutate(EC_MO_ID = 389, AREA_TERRENO1 = "0,01",AREA_TERRENO2 = "48,99")

df_completo_1 <- rbind(df_completo, df_t01_2_terreno) %>% as.data.frame()

# df_separados <- pmap(list(c(rep("input/T01.xlsx", 8), rep("input/T02.xlsx", 4), "input/T04.xlsx", rep("input/T05.xlsx", 4), rep("input/T03.xlsx", 4), rep("input/T06.xlsx", 4), rep("input/T19.xlsx", 3)),
#                              c(paste0("T01-", 1:8), paste0("T02-", 1:4), "T04", paste0("T05-", 1:4), rep("T03", 4), rep("T06", 4), rep("T19", 3)),
#                              c(rep(NA, 17), "0", "350,01", "1000,01", "6000,01", "350", "500,01", "750,01", "2000,01", "0", "350,01", "1000,01"),
#                              c(rep(NA, 17), "350", "1000", "6000", "1000000", "500", "750", "2000", "5000", "350", "1000", "6000")),
#                         ~make_txt_file(excel_name = ..1, tabla_selected = ..2, AREA_USO1 = ..3, AREA_USO2 = ..4, tabla_correlativa = tabla_correlativa))


df_join <- df_completo %>% 
  mutate(TIPO_CARACT = case_when(TIPO_CARACT_RES1 != 0 | TIPO_CARACT_RES2 != 8 ~ paste0(TIPO_CARACT_RES1, "-",TIPO_CARACT_RES2),
                                 TRUE ~ "0"),
         AREA_USO = case_when(AREA_USO1 != 0 | AREA_USO2 != 1000000 ~ paste0(AREA_USO1, "-", AREA_USO2),
                              TRUE ~ "0")) %>% 
  dplyr::select(c("EC_MO_ID",
                  "CODIGO_USO1",
                  "CODIGO_USO2",
                  "TIPO_CARACT_RES1",
                  "TIPO_CARACT_RES2",
                  "EDAD_PREDIO1",
                  "PUNTAJE1",
                  "AREA_USO1",
                  "AREA_USO2",
                  "VAL_METRO_CUAD",
                  "TIPO_CARACT",
                  "AREA_USO")) %>% 
  left_join(tabla_correlativa[, c("codigo_produccion", "tabla")], by = c("EC_MO_ID" = "codigo_produccion"))

# 
# wb <- createWorkbook("Camilo Avellaneda")
# addWorksheet(wb, "Consolidado")
# writeData(wb, "Consolidado", df_completo_1)
# saveWorkbook(wb, paste0("output/", str_replace_all(Sys.Date(), c("-" = "", "2021" = "21")),
#                         nombre_export, ".xlsx"), overwrite = T)


df_to_sas_T05 <- df_join %>%
  filter(str_detect(tabla, "^T05")) %>% 
  split(., seq(nrow(.))) %>% 
  map_dfr(~marca_manzana_row(.x)) %>% 
  split(., seq(nrow(.))) %>% 
  map_dfr(~uso_row(.x))
  

df_to_sas_T06 <- df_join %>%
  filter(str_detect(tabla, "^T06")) %>% 
  split(., seq(nrow(.))) %>% 
  map_dfr(~marca_manzana_row(.x)) %>% 
  split(., seq(nrow(.))) %>% 
  map_dfr(~uso_row(.x))

save(df_to_sas_T05, df_to_sas_T06, file = "output/Base_sas_T05_T06.RData")

df_join %>% filter(tabla == "T02-2") %>% 
  group_by(TIPO_CARACT_RES2) %>% count
# uso_row(r[1,])
# r <- df_join %>% 
#   filter(tabla == "T02-2") %>% slice(1:10)



df_test <- df_join %>% 
  filter(tabla == "T02-2" & CODIGO_USO2 == 4 & EDAD_PREDIO1 == 0 & PUNTAJE1 == 0 & TIPO_CARACT_RES1 == 8) %>% 
  split(., seq(nrow(.))) %>% 
  map_dfr(~marca_manzana_row(.x))







df_to_sas <- df_to_sas %>% as.data.table()
df_to_sas[, .N, by = .(tabla, TIPO_CARACT_RES, CODIGO_USO)]
df_to_sas %>% setnames("tabla", "TABLA")
df_to_sas[, CODIGO_USO := str_pad(CODIGO_USO, width = 3, side = "left", pad = "0")]


wb <- createWorkbook("Camilo Avellaneda")
addWorksheet(wb, "Consolidado")
writeData(wb, "Consolidado", df_completo)
saveWorkbook(wb, paste0("output/", str_replace_all(Sys.Date(), c("-" = "", "2021" = "21")),
                        nombre_export, ".xlsx"), overwrite = T)

wb1 <- createWorkbook("Camilo Avellaneda")
addWorksheet(wb1, "Consolidado_est")
writeData(wb1, "Consolidado_est", consolidado[, c("EC_MO_ID", "EDAD_PREDIO1", "PUNTAJE1", "VAL_METRO_CUAD")])
saveWorkbook(wb1, paste0("output/", str_replace_all(Sys.Date(), c("-" = "", "2021" = "21")),
                        nombre_export_est, ".xlsx"), overwrite = T)



date <- str_replace_all(Sys.Date(), c("-" = "", "2021" = "21"))
write.csv(df_to_sas, paste0("output/", date, "_RESULTADO_BASE_VALORES_NPH_TABLAS_SAS.csv"), na = "", row.names = FALSE)
        
        





