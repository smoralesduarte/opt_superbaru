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
