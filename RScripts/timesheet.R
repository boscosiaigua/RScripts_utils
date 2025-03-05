# Emplenar timesheet a partir de les hores del fitxador #
# 
# Autora: Alba Marquez Torres
# Data: 04-03-2025
# 
# Aquest script passa les hores del fitxador al timesheet. Tenir en compte
# - Calcula el numero d'hores total per dia en format decimal. Per exemple, si es fan 4hores 30minuts, el resultat serà 4.5 
# - Calcula nomes les hores treballades presencials o teletreball
# - Les hores que NO calcula (formacio, camp, etc.) es crea un "avis" on les hores estan com a 999
# - A part d'actualitzar el timesheet, crea un arxiu anomenat "revisar_hores.xlsx" per comprovar els avisos
# - Calcula tant jornades intensives com parcials
# - Calcula les hores totals diaries al timesheet, en el cas de tenir mes d'un projecte, s'hauran de distribuir com correspo

#### 0. Carregar llibreries i arxius, establir directori de treball #### -------

library(readxl)
library(openxlsx)
library(dplyr)
library(stringr)
library(tidyr)

setwd("C:/Alba/z_altres/CTFC_documents/timesheet")

dir.create(file.path("output"), showWarnings = FALSE)

# hores del fitxador en excel
fitxador <- read_excel("Movimientos.xls", range = "B6:F372")

#Timesheet a actualitzar
timesheet <- loadWorkbook("model_timesheet_2025_75 HORES.xlsx")

#### 1. Modificar arxiu fitxador #### ------------------------------------------

#modifico l'arxiu original de fitxar amb les columnes que minteressa
#comproba a la consola que queden les columnes "Data" i "Marcatge"
(fitx<-fitxador[-c(2:4)]) 

fitx <- fitx %>%
  mutate(Data = as.Date(Data, format = "%d/%m/%Y")) %>%
  filter(!is.na(Data))

#### 2. Calcular diferencia hores a decimal #### -------------------------------

#funcio per calcular la diferencia d'hores
diff_hours <- function(marcatge) {
  if (is.na(marcatge) || marcatge == "") return(NA_real_)
  
  registros <- str_extract_all(marcatge, "(?:\\.?\\d*\\s*[EeSs])?\\s*\\d{2}:\\d{2}")[[1]]
  
  if (length(registros) < 2) return("999")
  
  etiquetas <- str_extract(registros, "^\\.?\\d*\\s*[EeSs]")
  horas <- str_extract(registros, "\\d{2}:\\d{2}")
  
  if (any(grepl("^\\.5[1-9]|\\.[6-9][0-9]?$", etiquetas))) return("999")
  
  if (any(is.na(horas))) return("999")
  
  horas <- as.POSIXct(paste0("2000-01-01 ", horas), format = "%Y-%m-%d %H:%M")
  
  orden <- order(horas)
  etiquetas <- etiquetas[orden]
  horas <- horas[orden]
  
  total_horas <- 0
  i <- 1
  while (i < length(horas)) {
    if (i + 1 <= length(horas)) {
      entrada <- horas[i]
      salida <- horas[i + 1]
      
      diferencia <- as.numeric(difftime(salida, entrada, units = "hours"))
      
      if (diferencia > 0) {
        total_horas <- total_horas + diferencia
      }
    }
    i <- i + 2  
  }
  
  return(round(total_horas, 2))
}

# diferencia de hores del fitxador en format "hora decimal" a la columna "Horas"
(fitx_dec_real <- fitx %>%
  mutate(Horas_real = sapply(Marcatges, diff_hours, simplify = TRUE)) %>%
  mutate(Horas_real = ifelse(is.na(Horas_real), "0", Horas_real)))

write.xlsx(fitx_dec_real, "output/hores_reals.xlsx")

(fitx_dec <- fitx_dec_real %>%
  mutate(Horas_real = as.numeric(Horas_real)) %>%  
  mutate(Horas = ifelse(Horas_real > 10 & Horas_real < 999, 10, Horas_real)) %>%
  select(-Horas_real))

#convertir data a mes/dia
fitx_dec_mes_dia <- fitx_dec %>%
  mutate(Data = as.Date(Data, format = "%d/%m/%Y"),
         Mes = format(Data, "%m"),
         Dia = format(Data, "%d"),
         Horas = ifelse(is.na(Horas), 0, Horas)) %>%
  select(-Data)  

#### 3a. Arxiu comprobar hores ftixador #### -----------------------------------

#pasar de format curt a format llarg per revisar les hores 
#que NO son presencials o teletreball (estan com a 999)
(fitx_long_check <- fitx_dec_mes_dia %>%
  mutate(Horas = as.character(Horas)) %>%  
  pivot_longer(cols = c(Marcatges, Horas), names_to = "Tipo", values_to = "Valor") %>%
  pivot_wider(names_from = Dia, values_from = Valor))

write.xlsx(fitx_long_check, "output/revisar_hores.xlsx")

#### 3b. Actualitzar timesheet #### --------------------------------------------
#pasar de format curt a format llarg
(fitx_long_final <- fitx_dec_mes_dia %>%
    select(-Marcatges) %>%
    mutate(Horas = as.numeric(Horas)) %>%
    pivot_longer(cols = c(Horas), names_to = "Tipo", values_to = "Valor") %>%
    pivot_wider(names_from = Dia, values_from = Valor))

#Convertir mesos com a numeros a nom
mesos <- c(
  "01" = "January", "02" = "February", "03" = "March", "04" = "April",
  "05" = "May", "06" = "June", "07" = "July", "08" = "August",
  "09" = "September", "10" = "October", "11" = "November", "12" = "December"
)

# Filtrar les filas de "Horas"
horas_data <- fitx_long_final %>% filter(Tipo == "Horas")

# Columnes de desti (B a AF → 2 a 32)
start_col <- 2  
end_col <- 32
start_Row <- 26 ## posa-ho a la fila 26, on estan les hores totals

# Passar les hores a la fulla d'excel corresponent segons el mes
for (i in 1:nrow(horas_data)) {
  mes_num <- horas_data$Mes[i] 
  if (mes_num %in% names(mesos)) {
    sheet_name <- mesos[[mes_num]] 
    horas_valores <- as.numeric(horas_data[i, -c(1,2)])
    writeData(timesheet, sheet_name, t(horas_valores), startCol = start_col, startRow = start_Row, colNames = FALSE)
  }
}

# Guardar arxiu
saveWorkbook(timesheet, "output/timesheet_actualizado.xlsx", overwrite = TRUE)
