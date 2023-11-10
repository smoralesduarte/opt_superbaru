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
  df, ni = 3, no = 2, dmus = 1,
  inputs = c(2, 3, 4), outputs = 5:6
)

result_superbaru <- model_basic(
  superbaru_io, orientation = "io", rts = "crs", dmu_eval = 1:8, dmu_ref = 1:8
)

efficiencies <- efficiencies(result_superbaru)

target_input <- targets(result_superbaru)$target_input
target_output <- targets(result_superbaru)$target_output

target_input[, ] <- sweep(as.matrix(target_input[, ]), 2, df_means[1:3], "*")
target_output[, ] <- sweep(as.matrix(target_output[, ]), 2, df_means[4:5], "*")

target <- cbind(target_input, target_output)

# plot inefficient dmus
plot(efficiencies)

# plot a constant line
abline(h = 1, col = "red")

plot(rnorm(100))



#-----------------------------------------------------------------------------
#intento
#dea con todas las profesiones que nos mandaron

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


#borro rows adicionales que están apareciendo
df_2 <- df_2[-c(2, 10, 11), ]
#reindex
rownames(df_2) <- 1:nrow(df_2)
#le quito al tible unas hojas adicionales que están apareciendo
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
#erase columns that contain a 0 
# Identify columns containing 0
cols_with_zero_df2 <- sapply(df_2, function(col) any(col == 0))

# Remove columns with 0
df_2 <- df_2[, !cols_with_zero_df2]


# join info_sucursales and df_2
df_2 <- df_2 %>%
  left_join(info_sucursales, by = c("sucursal" = "ubicación"))
#le pongo la información de ivu
df_2$area[4] = 382
# join info_transacciones and df
df_2 <- df_2 %>%
  left_join(info_transacciones, by = c("sucursal" = "sucursal"))
#le pongo la información de ivu  
df_2$"cant_transacciones"[4]= 332078
df_2$"ventaneta"[4]= 3271645

# make every column except sucursal double
df_2[, -1] <- sapply(df_2[, -1], as.double)
# store the means of each column
df_2_means <- colMeans(df_2[, -1])

# divide each column by its mean
for(i in 2:ncol(df_2)){
 df_2[, i] = (df_2[, i] / mean(df_2[, i]))
}

#modelo dea
superbaru_io_2 <- make_deadata(
  df_2, ni = 11, no = 2,
  inputs = 2:12, outputs = 13:14
)

result_superbaru_2 <- model_basic(
  superbaru_io_2, orientation = "io", rts = "crs", dmu_eval = 1:8, dmu_ref = 1:8
)

efficiencies_2 <- efficiencies(result_superbaru_2)

target_input_2 <- targets(result_superbaru_2)$target_input
target_output_2 <- targets(result_superbaru_2)$target_output


target_2 <- cbind(target_input_2, target_output_2)
#me toca multiplicar por las medias de las columnas guardadas es df_2_means
for(i in 1:13){
  target_2[, i] = target_2[ , i] * df_2_means[i]
}


##-----------------------------------------------------------------------
##-----------------------------------------------------------------------

#outputs
#ventas
info_transacciones_path <-
  "Data de Ventas/10112023_Data_TransaccionesVentaDiarias_12Meses.xlsx"

# Read and join
info_transacciones <- read_excel(info_transacciones_path) %>%
  rename_all(str_to_lower) %>%
  dplyr::rename(cant_transacciones = `#cant de transac`) %>%
  dplyr::rename(ventaneta = `$ventaneta`)

#convert date to a date class
info_transacciones$fecha <- as.Date(info_transacciones$fecha)
# Combine 'fecha' and 'hra_inic_transac' into a single datetime column
info_transacciones <- info_transacciones %>%
  mutate(DateTime = as.POSIXct(paste(fecha, hra_inic_transac), format = "%Y-%m-%d %H"))


#borro columnas innecesarias (hora, idcaja, cant_trasacciones, fecha)
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
#from min(hour(info_transacciones$DateTime))= 1am to max(hour(info_transacciones$DateTime)) =  23pm
#create a column called periodo
info_transacciones <- info_transacciones %>%
  mutate(dia_semana = weekdays(DateTime),
         periodo = cut(hour(DateTime), breaks = seq(0, 24, by = 4), labels = FALSE))

# Summarize total sales, difhours annd number of days in filters in each s,p, d, q
#pregunta: en num_dias estoy contando las apariciones de cada grupo de d, s, q, p
result_transacciones <- info_transacciones %>%
  group_by(sucursal, categoria_dia, dia_semana, periodo) %>%
  dplyr::summarize(ventas = sum(ventaneta), dif_hora = max(hour(DateTime)) - min(hour(DateTime)), num_dias = n())

#añadir columna de m2
# Merge the tibbles based on the 'store' column
result_transacciones <- left_join(result_transacciones, info_sucursales, by = c("sucursal"= "ubicación"))

# Add a new column that is our final output
result_transacciones <- result_transacciones %>%
  mutate(output_final = ventas/((num_dias)*(dif_hora)*(area)))