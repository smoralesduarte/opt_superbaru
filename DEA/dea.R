# Load Excel Files
library(readxl)
library(dplyr)
library(stringr)
library(ggplot2)
library(ggthemes)
library(dplyr)
library(readxl)
library(plyr)
library(deaR)
library(ggridges)
library(lubridate)
library(tidyr)
library(hms)

#-----------------------------------------------------------------------------
info_colaboradores_path <-
  "Data de Colaboradores/10112023_Colab Información de Colaboradores.xlsx"


# 4th row is the header
sucursal <- readxl::excel_sheets(info_colaboradores_path)
tibble <- lapply(
  sucursal,
  function(x) {
    readxl::read_excel(info_colaboradores_path, sheet = x, skip = 3) %>%
      rename_all(str_to_lower)
  }
)

# Get length of each sheet and convert to vector
num_empleados <- sapply(tibble, nrow)

sucursal[[1]] <- "INTERAMERICANO"
df <- data.frame(sucursal, num_empleados)

# count the number of occupations which begin with CA
df$cajeros <- sapply(
  tibble,
  function(x) sum(grepl("^CA", x$ocupación))
)

# substract cajeros from num_empleados
df$num_empleados <- df$num_empleados - df$cajeros

info_sucursales_path <- "Data de Sucursales/00_Información de Sucursales.xlsx"

# remove NA rows
# select only ubicación and area columns
info_sucursales <- readxl::read_excel(info_sucursales_path, sheet = 2) %>%
  rename_all(str_to_lower) %>%
  filter(!is.na(ubicación)) %>%
  dplyr::rename(area = `área (m2)`) %>%
  select(ubicación, area)

# round area to integer
info_sucursales$area <- round(info_sucursales$area)

# remove * characters from ubicación
info_sucursales$ubicación <- gsub("\\*", "", info_sucursales$ubicación)

# join info_sucursales and df
df <- df %>%
  left_join(info_sucursales, by = c("sucursal" = "ubicación"))

info_ventas_path_1 <-
  "Data de Ventas/10112023_Data_DepartamentosVentaDiarias_12Meses_01.xlsx"


info_ventas_path_2 <-
  "Data de Ventas/10112023_Data_DepartamentosVentaDiarias_12Meses_02.xlsx"

# Read and join
info_ventas <- rbind(
  readxl::read_excel(info_ventas_path_1),
  readxl::read_excel(info_ventas_path_2)
) %>%
  rename_all(str_to_lower) %>%
  dplyr::rename(ventaneta = `$ventaneta`)

# sum ventaneta for each sucursal
info_ventas <- info_ventas %>%
  group_by(sucursal) %>%
  summarise(ventaneta = sum(ventaneta))

# join info_ventas and df
# df <- df %>%
#   left_join(info_ventas, by = c("sucursal" = "sucursal"))

info_transacciones_path <-
  "Data de Ventas/10112023_Data_TransaccionesVentaDiarias_12Meses.xlsx"

# Read and join
info_transacciones <- read_excel(info_transacciones_path) %>%
  rename_all(str_to_lower) %>%
  dplyr::rename(cant_transacciones = `#cant de transac`) %>%
  dplyr::rename(ventaneta = `$ventaneta`) %>%
  select(cant_transacciones, ventaneta, sucursal)

# sum ventaneta and cant_transacciones for each sucursal
info_transacciones <- info_transacciones %>%
  group_by(sucursal) %>%
  summarise_all(sum)

# join info_transacciones and df
df <- df %>%
  left_join(info_transacciones, by = c("sucursal" = "sucursal"))

df_original <- df
# make every column except sucursal double
df[, -1] <- sapply(df[, -1], as.double)

# store the means of each column
df_means <- colMeans(df[, -1])

# divide each column by its mean
df[, -1] <- as.data.frame(sweep(as.matrix(df[, -1]), 2, df_means, "/"))


superbaru_io <- make_deadata(
  df,
  ni = 3, no = 2, dmus = 1,
  inputs = c(2, 3, 4), outputs = 5:6
)

result_superbaru <- model_basic(
  superbaru_io,
  orientation = "io", rts = "crs", dmu_eval = 1:8, dmu_ref = 1:8
)

efficiencies <- efficiencies(result_superbaru)

target_input <- targets(result_superbaru)$target_input
target_output <- targets(result_superbaru)$target_output

target_input[, ] <- sweep(as.matrix(target_input[, ]), 2, df_means[1:3], "*")
target_output[, ] <- sweep(as.matrix(target_output[, ]), 2, df_means[4:5], "*")

target <- cbind(target_input, target_output)

# plot inefficient dmus
# plot(result_superbaru)



#-----------------------------------------------------------------------------
# intento
# dea con todas las profesiones que nos mandaron

info_colaboradores_path <-
  "Data de Colaboradores/EMPLEADOS POR SUCURSAL OCTUBRE 2023 v2.xlsx"
# 4th row is the header
sucursal <- readxl::excel_sheets(info_colaboradores_path)
tibble <- lapply(
  sucursal,
  function(x) {
    readxl::read_excel(info_colaboradores_path, sheet = x, skip = 3) %>%
      rename_all(str_to_lower)
  }
)

# Get length of each sheet and convert to vector
num_empleados <- sapply(tibble, nrow)

sucursal[[1]] <- "INTERAMERICANO"
df_2 <- data.frame(sucursal, num_empleados)


# borro rows adicionales que están apareciendo
df_2 <- df_2[-c(2, 10, 11), ]
# reindex
rownames(df_2) <- 1:nrow(df_2)
# le quito al tible unas hojas adicionales que están apareciendo
tibble <- tibble[-c(2, 10, 11)]



# count the number of occupations which begin with aseador
df_2$aseadores <- sapply(
  tibble,
  function(x) sum(grepl("ASEADOR", x$"ocupación agrupada"))
)
# count the number of occupations which begin with AYUDANTE DE BODEGA
df_2$"ayudante de bodega" <- sapply(
  tibble,
  function(x) sum(grepl("AYUDANTE DE BODEGA", x$"ocupación agrupada"))
)
# count the number of occupations which begin with  CAJERO/RA
df_2$"cajero" <- sapply(
  tibble,
  function(x) sum(grepl("CAJERO/RA", x$"ocupación agrupada"))
)
# count the number of occupations which begin with  DEPENDIENTE DE PANADERÍA/CAFETERÍA
df_2$"dep panadería/cafetería" <- sapply(
  tibble,
  function(x) sum(grepl("DEPENDIENTE DE PANADERÍA/CAFETERÍA", x$"ocupación agrupada"))
)
# count the number of occupations which begin with  DEPENDIENTE DE CARNE/DELI
df_2$"dep carne/deli" <- sapply(
  tibble,
  function(x) sum(grepl("DEPENDIENTE DE CARNE/DELI", x$"ocupación agrupada"))
)
# count the number of occupations which begin with   GERENCIA DE SUCURSAL
df_2$"gerencia" <- sapply(
  tibble,
  function(x) sum(grepl("GERENCIA DE SUCURSAL", x$"ocupación agrupada"))
)
# count the number of occupations which begin with   GONDOLEROS
df_2$"gondolero" <- sapply(
  tibble,
  function(x) sum(grepl("GONDOLEROS", x$"ocupación agrupada"))
)
# count the number of occupations which begin with   OFICIAL DE PROCESOS
df_2$"oficial procesos" <- sapply(
  tibble,
  function(x) sum(grepl("OFICIAL DE PROCESOS", x$"ocupación agrupada"))
)
# count the number of occupations which begin with  PAQUETERA
df_2$"paquetera" <- sapply(
  tibble,
  function(x) sum(grepl("PAQUETERA", x$"ocupación agrupada"))
)
# count the number of occupations which begin with RECIBIDOR/DESPACHADOR
df_2$"recibidor" <- sapply(
  tibble,
  function(x) sum(grepl("RECIBIDOR/DESPACHADOR", x$"ocupación agrupada"))
)
# count the number of occupations which begin with    SEGURIDAD Y VIGILANCIA
df_2$"seguridad" <- sapply(
  tibble,
  function(x) sum(grepl("SEGURIDAD Y VIGILANCIA", x$"ocupación agrupada"))
)
# count the number of occupations which begin with     SUPERVISOR DE BODEGA
df_2$"supervisor bodega" <- sapply(
  tibble,
  function(x) sum(grepl("SUPERVISOR DE BODEGA", x$"ocupación agrupada"))
)
# count the number of occupations which begin with SUPERVISOR DE SUCURSAL
df_2$"supervisor de sucursal" <- sapply(
  tibble,
  function(x) sum(grepl("SUPERVISOR DE SUCURSAL", x$"ocupación agrupada"))
)
# count the number of occupations which begin with     SUPERVISOR DE GONDOLAS
df_2$"supervisor de gondolas" <- sapply(
  tibble,
  function(x) sum(grepl("SUPERVISOR DE GONDOLAS", x$"ocupación agrupada"))
)
# count the number of occupations which begin with  DEPENDIENTE DE FRUTAS/VEGETALES
df_2$"dep frutas y verduras" <- sapply(
  tibble,
  function(x) sum(grepl("DEPENDIENTE DE FRUTAS/VEGETALES", x$"ocupación agrupada"))
)

#--------------------------------------------------------
# erase columns that contain a 0
# Identify columns containing 0
cols_with_zero_df2 <- sapply(df_2, function(col) any(col == 0))

# Remove columns with 0
df_2 <- df_2[, !cols_with_zero_df2]


# join info_sucursales and df_2
df_2 <- df_2 %>%
  left_join(info_sucursales, by = c("sucursal" = "ubicación"))
# le pongo la información de ivu
df_2$area[4] <- 382
# join info_transacciones and df
df_2 <- df_2 %>%
  left_join(info_transacciones, by = c("sucursal" = "sucursal"))
# le pongo la información de ivu
df_2$"cant_transacciones"[4] <- 332078
df_2$"ventaneta"[4] <- 3271645

# make every column except sucursal double
df_2[, -1] <- sapply(df_2[, -1], as.double)
# store the means of each column
df_2_means <- colMeans(df_2[, -1])

# divide each column by its mean
for (i in 2:ncol(df_2)) {
  df_2[, i] <- (df_2[, i] / mean(df_2[, i]))
}

# modelo dea
superbaru_io_2 <- make_deadata(
  df_2,
  ni = 11, no = 2,
  inputs = 2:12, outputs = 13:14
)

result_superbaru_2 <- model_basic(
  superbaru_io_2,
  orientation = "io", rts = "crs", dmu_eval = 1:8, dmu_ref = 1:8
)

efficiencies_2 <- efficiencies(result_superbaru_2)

target_input_2 <- targets(result_superbaru_2)$target_input
target_output_2 <- targets(result_superbaru_2)$target_output


target_2 <- cbind(target_input_2, target_output_2)
# me toca multiplicar por las medias de las columnas guardadas es df_2_means
for (i in 1:13) {
  target_2[, i] <- target_2[, i] * df_2_means[i]
}


## -----------------------------------------------------------------------
## -----------------------------------------------------------------------

# outputs
# ventas
info_transacciones_path <-
  "Data de Ventas/10112023_Data_TransaccionesVentaDiarias_12Meses.xlsx"

# Read and join
info_transacciones <- read_excel(info_transacciones_path) %>%
  rename_all(str_to_lower) %>%
  dplyr::rename(cant_transacciones = `#cant de transac`) %>%
  dplyr::rename(ventaneta = `$ventaneta`)

# convert date to a date class
info_transacciones$fecha <- as.Date(info_transacciones$fecha)
# Combine 'fecha' and 'hra_inic_transac' into a single datetime column
info_transacciones <- info_transacciones %>%
  mutate(DateTime = as.POSIXct(paste(fecha, hra_inic_transac), format = "%Y-%m-%d %H"))


# borro columnas innecesarias (hora, idcaja, cant_trasacciones, fecha)
info_transacciones <- info_transacciones %>%
  select(-hra_inic_transac, -idcaja, -cant_transacciones, -fecha)


# Create a new column based on the day of the month for the three categories:
# quincena, fin_mes, otro
info_transacciones <- info_transacciones %>%
  mutate(categoria_dia = case_when(
    day(DateTime) %in% 15:17 ~ "quincena",
    day(DateTime) %in% 28:31 ~ "fin_mes",
    TRUE ~ "otro"
  ))

# Extract the day of the week and create a new column for 4-hour periods
# from min(hour(info_transacciones$DateTime))= 1am to max(hour(info_transacciones$DateTime)) =  23pm
# create a column called periodo
info_transacciones <- info_transacciones %>%
  mutate(
    dia_semana = weekdays(DateTime),
    periodo = cut(hour(DateTime), breaks = seq(6, 22, by = 4), labels = FALSE)
  )

# Summarize total sales, difhours annd number of days in filters in each s,p, d, q
# pregunta: en num_dias estoy contando las apariciones de cada grupo de d, s, q, p
result_transacciones <- info_transacciones %>%
  group_by(sucursal, categoria_dia, dia_semana, periodo) %>%
  dplyr::summarize(ventas = sum(ventaneta), dif_hora = max(hour(DateTime)) - min(hour(DateTime)), num_dias = length(unique(DateTime)))




# añadir columna de m2
# Merge the tibbles based on the 'store' column
result_transacciones <- left_join(result_transacciones, info_sucursales, by = c("sucursal" = "ubicación"))

# Add a new column that is our final output
# result_transacciones <- result_transacciones %>%
#   mutate(output_final = ventas / ((num_dias) * (dif_hora) * (area)))
result_transacciones <- result_transacciones %>%
  mutate(output_final = ventas / (num_dias))

#CAMBIÉ: QUITO CATEGORÍA DE FIN_MES Y QINCENA
# Filter rows where 'categoria_dia' is equal to 'otro'
result_transacciones <- result_transacciones %>%
  filter(categoria_dia == "otro")


###### -------------------------------------------------------
# info_horarios_doleguita_path <-
#   "horarios empleados/HORARIO DE DOLEGUITA.xlsx"
# info_horarios_boquete_path <-
#   "horarios empleados/HORARIO BOQUETE.xlsx"
# info_horarios_inter_path <-
#   "horarios empleados/HORARIO INTERAMERICANO.XLSX"
# info_horarios_ivu_path <-
#   "horarios empleados/HORARIO IVU DOS PINOS.xlsx"
# info_horarios_mall_path <-
#   "horarios empleados/HORARIO MALL.xlsx"
# info_horarios_riviera_path <-
#   "horarios empleados/HORARIO RIVIERA.xlsx"
# info_horarios_volcan_path <-
#   "horarios empleados/HORARIO VOLCAN.xlsx"
# info_horarios_sancristobal_path <-
#   "horarios empleados/HORARIO SAN CRISTOBAL.xlsx"

# # Read doleguita, from 3rd row. Remove rows that contain any NA in any column
# info_horarios_doleguita <- read_excel(info_horarios_doleguita_path, skip = 7) %>%
#   rename_all(str_to_lower)

# # rename entrada...2 to e1, salida...3 to s1, entrada...4 to e2, salida...5 to s2
# names(info_horarios_doleguita) <- c("nombre", "e1", "s1", "e2", "s2", "ocupacion", "sucursal")
# info_horarios_doleguita <- info_horarios_doleguita %>%
#   dplyr::filter(!is.na(e2))


# info_horarios_boquete <- read_excel(info_horarios_boquete_path, skip = 6) %>%
#   select(-1) %>%
#   rename_all(str_to_lower)

# names(info_horarios_boquete) <- c("nombre", "e1", "s1", "e2", "s2", "ocupacion", "sucursal")
# info_horarios_boquete <- info_horarios_boquete %>%
#   dplyr::filter(!is.na(e2))


# info_horarios_inter <- read_excel(info_horarios_inter_path, skip=3) %>%
#   select(-1) %>%
#   rename_all(str_to_lower)

# names(info_horarios_inter) <- c("nombre", "e1", "s1", "e2", "s2", "ocupacion", "sucursal", "none")
# info_horarios_inter <- info_horarios_inter %>%
#   dplyr::filter(!is.na(e2)) %>%
#   select(-none)

# info_horarios_ivu <- read_excel(info_horarios_ivu_path, skip=3) %>%
#   select(-1) %>%
#   rename_all(str_to_lower)

# names(info_horarios_ivu) <- c("nombre", "e1", "s1", "e2", "s2", "ocupacion", "sucursal")
# info_horarios_ivu <- info_horarios_ivu %>%
#   dplyr::filter(!is.na(e2))


# info_horarios_mall <- read_excel(info_horarios_mall_path, skip=3) %>%
#   rename_all(str_to_lower) %>%
#   select(-firma) %>%
#   select(-1)

# names(info_horarios_mall) <- c("nombre", "e1", "s1", "e2", "s2", "ocupacion", "sucursal")
# info_horarios_mall <- info_horarios_mall %>%
#   dplyr::filter(!is.na(e2))

# info_horarios_riviera <- read_excel(info_horarios_riviera_path, skip=3) %>%
#   rename_all(str_to_lower) %>%
#   select(-1)

# names(info_horarios_riviera) <- c("nombre", "e1", "s1", "e2", "s2", "ocupacion", "sucursal")
# info_horarios_riviera <- info_horarios_riviera %>%
#   dplyr::filter(!is.na(e2))

# info_horarios_volcan <- read_excel(info_horarios_volcan_path, skip=6) %>%
#   rename_all(str_to_lower) %>%
#   select(-1)

# names(info_horarios_volcan) <- c("nombre", "e1", "s1", "e2", "s2", "ocupacion", "sucursal")
# info_horarios_volcan <- info_horarios_volcan %>%
#   dplyr::filter(!is.na(e2))


# info_horarios_sancristobal <- read_excel(info_horarios_sancristobal_path, skip=2) %>%
#   rename_all(str_to_lower)

# names(info_horarios_sancristobal) <- c("nombre", "e1", "s1", "e2", "s2", "ocupacion", "sucursal")
# info_horarios_sancristobal <- info_horarios_sancristobal %>%
#   dplyr::filter(!is.na(e2))


# # Merge the tibbles
# tibble_horarios <- rbind(
#   info_horarios_doleguita,
#   info_horarios_boquete,
#   info_horarios_inter,
#   info_horarios_ivu,
#   info_horarios_mall,
#   info_horarios_riviera,
#   info_horarios_volcan,
#   info_horarios_sancristobal
# )

####------------------------
# Run from here

excel_file_horarios <- "horarios empleados/Data de Empleados 22nov2023 v5.xlsx"
sheet_names_horarios <- excel_sheets(excel_file_horarios)

# Read all sheets into a list
all_data_horarios <- lapply(sheet_names_horarios, function(sheet) {
  read_excel(excel_file_horarios, sheet = sheet)
})

horarios_final <- rbind(all_data_horarios[[1]], all_data_horarios[[2]],
 all_data_horarios[[3]], all_data_horarios[[4]], 
 all_data_horarios[[5]], all_data_horarios[[6]], 
 all_data_horarios[[7]], all_data_horarios[[8]])

#dejo sólo las filas que vacaciones sea no
horarios_final <- horarios_final %>%
  filter(vacaciones == "NO")

#dejo sólo las profesiones que nos interesa
horarios_final <- horarios_final %>%
  filter(ocupacion %in% c("ASEADOR", "CAJERO/RA", "GONDOLEROS", "DEPENDIENTE DE CARNE/DELI", "DEPENDIENTE DE FRUTAS/VEGETALES" ))


#borro la columna vacaciones
horarios_final <- horarios_final %>%
  select(-vacaciones)

tibble_horarios <- horarios_final

# Convert 'your_column' to hour format
tibble_horarios <- tibble_horarios %>%
  mutate(e1 = as_hms(e1))
tibble_horarios <- tibble_horarios %>%
  mutate(e2 = as_hms(e2))
tibble_horarios <- tibble_horarios %>%
  mutate(s1 = as_hms(s1))
tibble_horarios <- tibble_horarios %>%
  mutate(s2 = as_hms(s2))



#####-------------------------------------
# Turn every date in the tibble into the hour as a fraction
tibble_horarios <- tibble_horarios %>%
  mutate(
    e1 = hour(e1) + minute(e1) / 60,
    s1 = hour(s1) + minute(s1) / 60,
    e2 = hour(e2) + minute(e2) / 60,
    s2 = hour(s2) + minute(s2) / 60
    )

# e1 is the first entry time, e2 is the second entry time, s1 is the first exit time, s2 is the second exit time
# There are 4 periods, 6-10, 10-14, 14-18, 18-22
p1  <- c(6, 10)
p2 <- c(10, 14)
p3 <- c(14, 18)
p4 <- c(18, 22)

# Create a function that calculates the hours worked in a period
horas_trabajadas <- function(e1, s1, period) {
  mincol <- function(x, y) {
    ifelse(x < y, x, y)
  }
  
  maxcol <- function(x, y) {
    ifelse(x > y, x, y)
  }

  ifelse(
    (mincol(s1, period[2]) - maxcol(e1, period[1])) > 0,
    mincol(s1, period[2]) - maxcol(e1, period[1]),
    0
  )

}


# Make 4 extra columns, counting the hours worked in each period per person
tibble_horarios <- tibble_horarios %>%
  mutate(
    horas_p1 = horas_trabajadas(e1, s1, p1) +
      horas_trabajadas(e2, s2, p1),
    horas_p2 = horas_trabajadas(e1, s1, p2) +
      horas_trabajadas(e2, s2, p2),
    horas_p3 = horas_trabajadas(e1, s1, p3) +
      horas_trabajadas(e2, s2, p3),
    horas_p4 = horas_trabajadas(e1, s1, p4) +
      horas_trabajadas(e2, s2, p4)
  )

# group by ocupacion and sucursal, and summarize the total hours worked in each period
tibble_horarios <- tibble_horarios %>%
  group_by(ocupacion, sucursal) %>%
  dplyr::summarize(
    horas_p1 = sum(horas_p1),
    horas_p2 = sum(horas_p2),
    horas_p3 = sum(horas_p3),
    horas_p4 = sum(horas_p4)
  )

# make a new tibble, the columns are the occupations and the rows are the periods and sucursal
tibble_horarios <- tibble_horarios %>%
  pivot_longer(
    cols = c(horas_p1, horas_p2, horas_p3, horas_p4),
    names_to = "periodo",
    values_to = "horas"
  ) %>%
  pivot_wider(
    names_from = "ocupacion",
    values_from = "horas"
  )

# remove `` from the column names
# names(tibble_horarios) <- gsub("'", "", names(tibble_horarios))
#error
result_transacciones <- result_transacciones %>%
  dplyr::filter(!is.na(periodo))

# rename periodo entries to an integer (horas_p1 to 1, horas_p2 to 2, etc)
tibble_horarios <- tibble_horarios %>%
  mutate(periodo = as.integer(gsub("horas_p", "", periodo)))

tibble_horarios <- tibble_horarios %>%
  mutate_if(is.character, str_trim)

################################
#cambios para sacar modelo después de hacer los cambios
tibble_optimizada <- tibble_horarios
#IVU
#2 cajero menos periodo 2 y periodo 3
tibble_optimizada <- tibble_optimizada %>%
  mutate(`CAJERO/RA` = ifelse(sucursal == "IVU DOS PINOS" & (periodo == 2 | periodo == 3),
  `CAJERO/RA` - 2 * 4,`CAJERO/RA`))
#1 carnes menos periodo 2
tibble_optimizada <- tibble_optimizada %>%
  mutate(`DEPENDIENTE DE CARNE/DELI` = ifelse(sucursal == "IVU DOS PINOS" & (periodo == 2 ),
  `DEPENDIENTE DE CARNE/DELI` - 4,`DEPENDIENTE DE CARNE/DELI`))
#2carnes periodo 3
tibble_optimizada <- tibble_optimizada %>%
  mutate(`DEPENDIENTE DE CARNE/DELI` = ifelse(sucursal == "IVU DOS PINOS" & (periodo == 3 ),
  `DEPENDIENTE DE CARNE/DELI` - 2*4,`DEPENDIENTE DE CARNE/DELI`))
#1 frutas p2
tibble_optimizada <- tibble_optimizada %>%
  mutate( `DEPENDIENTE DE FRUTAS/VEGETALES` = ifelse(sucursal == "IVU DOS PINOS" & (periodo == 2 ),
   `DEPENDIENTE DE FRUTAS/VEGETALES` - 4, `DEPENDIENTE DE FRUTAS/VEGETALES`))
#2 gondoleros p2
tibble_optimizada <- tibble_optimizada %>%
  mutate( GONDOLEROS = ifelse(sucursal == "IVU DOS PINOS" & (periodo == 2 ),
   GONDOLEROS - 2*4, GONDOLEROS))
#1 gondolero p3
tibble_optimizada <- tibble_optimizada %>%
  mutate( GONDOLEROS = ifelse(sucursal == "IVU DOS PINOS" & (periodo == 3 ),
   GONDOLEROS - 4, GONDOLEROS))

#RIVIERA
#2 cajeros p2 y p3
tibble_optimizada <- tibble_optimizada %>%
  mutate(`CAJERO/RA` = ifelse(sucursal == "RIVIERA" & (periodo == 2 | periodo == 3),
  `CAJERO/RA` - 2 * 4,`CAJERO/RA`))
#1 cajero p1
tibble_optimizada <- tibble_optimizada %>%
  mutate(`CAJERO/RA` = ifelse(sucursal == "RIVIERA" & (periodo == 1),
  `CAJERO/RA` - 4,`CAJERO/RA`))
#1 carnes p2 y p1
tibble_optimizada <- tibble_optimizada %>%
  mutate(`DEPENDIENTE DE CARNE/DELI` = ifelse(sucursal == "RIVIERA" & (periodo == 2 | periodo == 1 ),
  `DEPENDIENTE DE CARNE/DELI` - 4,`DEPENDIENTE DE CARNE/DELI`))
#1 frutas p2 y p3
tibble_optimizada <- tibble_optimizada %>%
  mutate( `DEPENDIENTE DE FRUTAS/VEGETALES` = ifelse(sucursal == "RIVIERA" & (periodo == 2| periodo == 3 ),
   `DEPENDIENTE DE FRUTAS/VEGETALES` - 4, `DEPENDIENTE DE FRUTAS/VEGETALES`))
# 1 gondolero p1 y p2
tibble_optimizada <- tibble_optimizada %>%
  mutate( GONDOLEROS = ifelse(sucursal == "RIVIERA" & (periodo == 1 | periodo == 2 ),
   GONDOLEROS - 4, GONDOLEROS))

#MALL 
#4 cajeros p2
tibble_optimizada <- tibble_optimizada %>%
  mutate(`CAJERO/RA` = ifelse(sucursal == "MALL" & (periodo == 2 ),
  `CAJERO/RA` - 4 * 4,`CAJERO/RA`))
#3 cajeros p3
tibble_optimizada <- tibble_optimizada %>%
  mutate(`CAJERO/RA` = ifelse(sucursal == "MALL" & (periodo == 3),
  `CAJERO/RA` - 3 * 4,`CAJERO/RA`))
#1 carnes p2 y p3
tibble_optimizada <- tibble_optimizada %>%
  mutate(`DEPENDIENTE DE CARNE/DELI` = ifelse(sucursal == "MALL" & (periodo == 2 | periodo == 3 ),
  `DEPENDIENTE DE CARNE/DELI` - 4,`DEPENDIENTE DE CARNE/DELI`))
#1 gondolero p3  y p2
tibble_optimizada <- tibble_optimizada %>%
  mutate( GONDOLEROS = ifelse(sucursal == "MALL" & (periodo == 3 | periodo == 2 ),
   GONDOLEROS - 4, GONDOLEROS))

#SAN CRISTÓBAL
# 2 cajeros p3
tibble_optimizada <- tibble_optimizada %>%
  mutate(`CAJERO/RA` = ifelse(sucursal == "SAN CRISTOBAL" & (periodo == 3),
  `CAJERO/RA` - 2 * 4,`CAJERO/RA`))
# 1 carnes p3
tibble_optimizada <- tibble_optimizada %>%
  mutate(`DEPENDIENTE DE CARNE/DELI` = ifelse(sucursal == "SAN CRISTOBAL" & (periodo == 3 ),
  `DEPENDIENTE DE CARNE/DELI` - 4,`DEPENDIENTE DE CARNE/DELI`))
#1 gondolero p3
tibble_optimizada <- tibble_optimizada %>%
  mutate( GONDOLEROS = ifelse(sucursal == "SAN CRISTOBAL" & (periodo == 3 ),
   GONDOLEROS - 4, GONDOLEROS))

#correr la siguiente línea si se quieren los resultados de la optimización
tibble_horarios <- tibble_optimizada

result_transacciones <- result_transacciones %>%
  left_join(tibble_horarios,
    by = c("sucursal" = "sucursal", "periodo" = "periodo")
  )





name <- apply(result_transacciones[1:4], 1, paste, collapse = " ")

final_dea <- tibble(name,
  result_transacciones[, 10:ncol(result_transacciones)],
  output_final = result_transacciones$output_final
)

# remove every row that contains a 0 entry
final_dea_original_zero <- cbind(result_transacciones[, 1:4], final_dea[, -1])
result_transacciones <- result_transacciones[apply(final_dea[, -1], 1, function(x) !any(x == 0)), ]
final_dea <- final_dea[apply(final_dea[, -1], 1, function(x) !any(x == 0)), ]
final_dea_original <- final_dea

# store the means of each column
final_dea_means <- colMeans(final_dea[, -1])

# divide each column by its mean
final_dea[, -1] <- as.data.frame(sweep(as.matrix(final_dea[, -1]), 2, final_dea_means, "/"))


final_dea_io <- make_deadata(
  final_dea,
  ni = 5, no = 1, dmus = 1,
  inputs = 2:6, outputs = 7
)

result_final_dea <- model_basic(
  final_dea_io,
  orientation = "io", rts = "crs"
  # , dmu_eval = 1:8, dmu_ref = 1:8
)

efficiencies <- efficiencies(result_final_dea)

target_input <- targets(result_final_dea)$target_input
target_output <- targets(result_final_dea)$target_output

target_input[, ] <- sweep(as.matrix(target_input[, ]), 2, final_dea_means[1:5], "*")
target_output[, ] <- sweep(as.matrix(target_output[, ]), 2, final_dea_means[6], "*")

target <- cbind(target_input, target_output)

# plot inefficient dmus
plot(result_final_dea)

# select the rows of the original and the target for which the inefficiencies are less than 0.6
final_dea_inefficient <- final_dea_original[efficiencies < 0.5, ]
target_inefficient <- target[efficiencies < 0.5, ]

# add efficiency column to final_dea_original and to target and first 4 columns of result_transacciones

final_dea_original <- cbind(result_transacciones[, 1:4], final_dea_original[, -1], efficiencies)
target <- cbind(result_transacciones[, 1:4], target, efficiencies)

# Export to excel with pretty formatting
library(openxlsx)
wb <- createWorkbook()
addWorksheet(wb, "Original")
writeData(wb, "Original", final_dea_original)
addWorksheet(wb, "Target")
writeData(wb, "Target", target)
addWorksheet(wb, "Original including 0")
writeData(wb, "Original including 0", final_dea_original_zero)
saveWorkbook(wb, "DEA.xlsx", overwrite = TRUE)

##-------------------------------------------------
#saber qué tan ocupadas están las cajas
# outputs
# ventas
info_transacciones_path <-
  "Data de Ventas/10112023_Data_TransaccionesVentaDiarias_12Meses.xlsx"

# Read and join
info_transacciones <- read_excel(info_transacciones_path) %>%
  rename_all(str_to_lower) %>%
  dplyr::rename(cant_transacciones = `#cant de transac`) %>%
  dplyr::rename(ventaneta = `$ventaneta`)

# convert date to a date class
info_transacciones$fecha <- as.Date(info_transacciones$fecha)
# Combine 'fecha' and 'hra_inic_transac' into a single datetime column
info_transacciones <- info_transacciones %>%
  mutate(DateTime = as.POSIXct(paste(fecha, hra_inic_transac), format = "%Y-%m-%d %H"))


# borro columnas innecesarias (hora, idcaja, cant_trasacciones, fecha)
info_transacciones <- info_transacciones %>%
  select(-hra_inic_transac, -cant_transacciones, -fecha)


# Create a new column based on the day of the month for the three categories:
# quincena, fin_mes, otro
info_transacciones <- info_transacciones %>%
  mutate(categoria_dia = case_when(
    day(DateTime) %in% 15:17 ~ "quincena",
    day(DateTime) %in% 28:31 ~ "fin_mes",
    TRUE ~ "otro"
  ))

# Extract the day of the week and create a new column for 4-hour periods
# from min(hour(info_transacciones$DateTime))= 1am to max(hour(info_transacciones$DateTime)) =  23pm
# create a column called periodo
info_transacciones <- info_transacciones %>%
  mutate(
    dia_semana = weekdays(DateTime),
    periodo = cut(hour(DateTime), breaks = seq(6, 22, by = 4), labels = FALSE)
  )

# Summarize total sales, difhours annd number of days in filters in each s,p, d, q
# pregunta: en num_dias estoy contando las apariciones de cada grupo de d, s, q, p
result_transacciones <- info_transacciones %>%
  group_by(sucursal, categoria_dia, dia_semana, periodo, idcaja) %>%
  dplyr::summarize(ventas_prom = mean(ventaneta))


# añadir columna de m2
# Merge the tibbles based on the 'store' column
result_transacciones <- left_join(result_transacciones, info_sucursales, by = c("sucursal" = "ubicación"))

# Add a new column that is our final output
# result_transacciones <- result_transacciones %>%
#   mutate(output_final = ventas / ((num_dias) * (dif_hora) * (area)))
result_transacciones <- result_transacciones %>%
   group_by(sucursal, categoria_dia, dia_semana, periodo) %>%
  dplyr::summarize(ventas_sucursal = sum(ventas_prom)/n_distinct(idcaja))

# Filter rows where 'categoria_dia' is equal to 'otro'
result_transacciones <- result_transacciones %>%
  filter(categoria_dia == "otro")

# Specify the file path for your Excel file
excel_file_path <- "ventas.xlsx"

# Write the tibble to an Excel file
write.xlsx(result_transacciones, excel_file_path)


######
##########lo mismo pero con #transacciones
# Read and join
info_transacciones <- read_excel(info_transacciones_path) %>%
  rename_all(str_to_lower) %>%
  dplyr::rename(cant_transacciones = `#cant de transac`) %>%
  dplyr::rename(ventaneta = `$ventaneta`)

# convert date to a date class
info_transacciones$fecha <- as.Date(info_transacciones$fecha)
# Combine 'fecha' and 'hra_inic_transac' into a single datetime column
info_transacciones <- info_transacciones %>%
  mutate(DateTime = as.POSIXct(paste(fecha, hra_inic_transac), format = "%Y-%m-%d %H"))


# borro columnas innecesarias (hora, idcaja, cant_trasacciones, fecha)
info_transacciones <- info_transacciones %>%
  select(-hra_inic_transac, -ventaneta, -fecha)


# Create a new column based on the day of the month for the three categories:
# quincena, fin_mes, otro
info_transacciones <- info_transacciones %>%
  mutate(categoria_dia = case_when(
    day(DateTime) %in% 15:17 ~ "quincena",
    day(DateTime) %in% 28:31 ~ "fin_mes",
    TRUE ~ "otro"
  ))

# Extract the day of the week and create a new column for 4-hour periods
# from min(hour(info_transacciones$DateTime))= 1am to max(hour(info_transacciones$DateTime)) =  23pm
# create a column called periodo
info_transacciones <- info_transacciones %>%
  mutate(
    dia_semana = weekdays(DateTime),
    periodo = cut(hour(DateTime), breaks = seq(6, 22, by = 4), labels = FALSE)
  )

# Summarize total sales, difhours annd number of days in filters in each s,p, d, q
# pregunta: en num_dias estoy contando las apariciones de cada grupo de d, s, q, p
result_transacciones <- info_transacciones %>%
  group_by(sucursal, categoria_dia, dia_semana, periodo, idcaja) %>%
  dplyr::summarize(transacciones_prom = mean(cant_transacciones))


# añadir columna de m2
# Merge the tibbles based on the 'store' column
result_transacciones <- left_join(result_transacciones, info_sucursales, by = c("sucursal" = "ubicación"))

# Add a new column that is our final output
# result_transacciones <- result_transacciones %>%
#   mutate(output_final = ventas / ((num_dias) * (dif_hora) * (area)))
result_transacciones <- result_transacciones %>%
   group_by(sucursal, categoria_dia, dia_semana, periodo) %>%
  dplyr::summarize(transacciones_sucursal = sum(transacciones_prom)/n_distinct(idcaja))

# Filter rows where 'categoria_dia' is equal to 'otro'
result_transacciones <- result_transacciones %>%
  filter(categoria_dia == "otro")

# Specify the file path for your Excel file
excel_file_path <- "transacciones.xlsx"

# Write the tibble to an Excel file
write.xlsx(result_transacciones, excel_file_path)