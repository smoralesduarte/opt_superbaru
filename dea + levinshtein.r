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

#info colaboradores
info_colaboradores_inter <- read_excel(info_colaboradores_path, sheet = 1)

#borro las primeras filas que no tienen info
info_colaboradores_inter <- info_colaboradores_inter[-c(1,2,3),]
ocupaciones_inter <- info_colaboradores_inter[[6]]
ocupaciones_inter <- data.frame(Column_Name = ocupaciones_inter)
#borro "ocupacion"

#borro repetidas
ocupaciones_inter <- ocupaciones_inter %>% distinct
ocupaciones_inter <- ocupaciones_inter[-1,]



library(stringdist)
library(tidyverse)

# Calculate pairwise Levenshtein distances
distance_matrix_inter <- stringdistmatrix(ocupaciones_inter, ocupaciones_inter, method = "lv")


#find closest element
closest_elements_inter <- sapply(1:length(ocupaciones_inter), function(i) {
  min_distance_inter <- min(distance_matrix_inter[i, -i])
  closest_indices_inter <- which(distance_matrix_inter[i, -i] == min_distance_inter)
  closest_strings_inter <- ocupaciones_inter[-i][closest_indices_inter]
  return(closest_strings_inter)
})

#display results
for (i in 1:length(ocupaciones_inter)) {
  cat(paste("Closest elements to", ocupaciones_inter[i], "are:", closest_elements_inter[[i]], "\n"))
}




#ivu
#info colaboradores
info_colaboradores_ivu <- read_excel(info_colaboradores_path, sheet = 2)

#borro las primeras filas que no tienen info
info_colaboradores_ivu <- info_colaboradores_ivu[-c(1,2,3),]
ocupaciones_ivu <- info_colaboradores_ivu[[6]]
ocupaciones_ivu <- data.frame(Column_Name = ocupaciones_ivu)
#borro "ocupacion"

#borro repetidas
ocupaciones_ivu <- ocupaciones_ivu %>% distinct
ocupaciones_ivu <- ocupaciones_ivu[-1,]


# Calculate pairwise Levenshtein distances
distance_matrix_ivu <- stringdistmatrix(ocupaciones_ivu, ocupaciones_ivu, method = "lv")


#find closest element
closest_elements_ivu <- sapply(1:length(ocupaciones_ivu), function(i) {
  min_distance_ivu <- min(distance_matrix_ivu[i, -i])
  closest_indices_ivu <- which(distance_matrix_ivu[i, -i] == min_distance_ivu)
  closest_strings_ivu <- ocupaciones_ivu[-i][closest_indices_ivu]
  return(closest_strings_ivu)
})

#display results
for (i in 1:length(ocupaciones_ivu)) {
  cat(paste("Closest elements to", ocupaciones_ivu[i], "are:", closest_elements_ivu[[i]], "\n"))
}


#san cristobal
#info colaboradores
info_colaboradores_san <- read_excel(info_colaboradores_path, sheet = 3)

#borro las primeras filas que no tienen info
info_colaboradores_san <- info_colaboradores_san[-c(1,2,3),]
ocupaciones_san <- info_colaboradores_san[[6]]
ocupaciones_san <- data.frame(Column_Name = ocupaciones_san)
#borro "ocupacion"

#borro repetidas
ocupaciones_san <- ocupaciones_san %>% distinct
ocupaciones_san <- ocupaciones_san[-1,]


# Calculate pairwise Levenshtein distances
distance_matrix_san <- stringdistmatrix(ocupaciones_san, ocupaciones_san, method = "lv")


#find closest element
closest_elements_san <- sapply(1:length(ocupaciones_san), function(i) {
  min_distance_san <- min(distance_matrix_san[i, -i])
  closest_indices_san <- which(distance_matrix_san[i, -i] == min_distance_san)
  closest_strings_san <- ocupaciones_san[-i][closest_indices_san]
  return(closest_strings_san)
})

#display results
for (i in 1:length(ocupaciones_san)) {
  cat(paste("Closest elements to", ocupaciones_san[i], "are:", closest_elements_san[[i]], "\n"))
}



#RIVIERA
#info colaboradores
info_colaboradores_riv <- read_excel(info_colaboradores_path, sheet = 4)

#borro las primeras filas que no tienen info
info_colaboradores_riv <- info_colaboradores_riv[-c(1,2,3),]
ocupaciones_riv <- info_colaboradores_riv[[6]]
#hay un empleado con ocupación vacía: era gondolero pero se corrió la bd
ocupaciones_riv[25] = "GONDOLERO"
ocupaciones_riv <- data.frame(Column_Name = ocupaciones_riv)
#borro "ocupacion"

#borro repetidas
ocupaciones_riv <- ocupaciones_riv %>% distinct
ocupaciones_riv <- ocupaciones_riv[-1,]


# Calculate pairwise Levenshtein distances
distance_matrix_riv <- stringdistmatrix(ocupaciones_riv, ocupaciones_riv, method = "lv")


#find closest element
closest_elements_riv <- sapply(1:length(ocupaciones_riv), function(i) {
  min_distance_riv <- min(distance_matrix_riv[i, -i])
  closest_indices_riv <- which(distance_matrix_riv[i, -i] == min_distance_riv)
  closest_strings_riv <- ocupaciones_riv[-i][closest_indices_riv]
  return(closest_strings_riv)
})

#display results
for (i in 1:length(ocupaciones_riv)) {
  cat(paste("Closest elements to", ocupaciones_riv[i], "are:", closest_elements_riv[[i]], "\n"))
}



#dol
#info colaboradores
info_colaboradores_dol <- read_excel(info_colaboradores_path, sheet = 5)

#borro las primeras filas que no tienen info
info_colaboradores_dol <- info_colaboradores_dol[-c(1,2,3),]
ocupaciones_dol <- info_colaboradores_dol[[6]]
ocupaciones_dol <- data.frame(Column_Name = ocupaciones_dol)
#borro "ocupacion"

#borro repetidas
ocupaciones_dol <- ocupaciones_dol %>% distinct
ocupaciones_dol <- ocupaciones_dol[-1,]


# Calculate pairwise Levenshtein distances
distance_matrix_dol <- stringdistmatrix(ocupaciones_dol, ocupaciones_dol, method = "lv")


#find closest element
closest_elements_dol <- sapply(1:length(ocupaciones_dol), function(i) {
  min_distance_dol <- min(distance_matrix_dol[i, -i])
  closest_indices_dol <- which(distance_matrix_dol[i, -i] == min_distance_dol)
  closest_strings_dol <- ocupaciones_dol[-i][closest_indices_dol]
  return(closest_strings_dol)
})

#display results
for (i in 1:length(ocupaciones_dol)) {
  cat(paste("Closest elements to", ocupaciones_dol[i], "are:", closest_elements_dol[[i]], "\n"))
}


#mall
#info colaboradores
info_colaboradores_mall <- read_excel(info_colaboradores_path, sheet = 6)

#borro las primeras filas que no tienen info
info_colaboradores_mall <- info_colaboradores_mall[-c(1,2,3),]
ocupaciones_mall <- info_colaboradores_mall[[6]]
ocupaciones_mall <- data.frame(Column_Name = ocupaciones_mall)
#borro "ocupacion"

#borro repetidas
ocupaciones_mall <- ocupaciones_mall %>% distinct
ocupaciones_mall <- ocupaciones_mall[-1,]


# Calculate pairwise Levenshtein distances
distance_matrix_mall <- stringdistmatrix(ocupaciones_mall, ocupaciones_mall, method = "lv")


#find closest element
closest_elements_mall <- sapply(1:length(ocupaciones_mall), function(i) {
  min_distance_mall <- min(distance_matrix_mall[i, -i])
  closest_indices_mall <- which(distance_matrix_mall[i, -i] == min_distance_mall)
  closest_strings_mall <- ocupaciones_mall[-i][closest_indices_mall]
  return(closest_strings_mall)
})

#display results
for (i in 1:length(ocupaciones_mall)) {
  cat(paste("Closest elements to", ocupaciones_mall[i], "are:", closest_elements_mall[[i]], "\n"))
}


#boquete
#info colaboradores
info_colaboradores_boq <- read_excel(info_colaboradores_path, sheet = 7)

#borro las primeras filas que no tienen info
info_colaboradores_boq <- info_colaboradores_boq[-c(1,2,3),]
ocupaciones_boq <- info_colaboradores_boq[[6]]
ocupaciones_boq <- data.frame(Column_Name = ocupaciones_boq)
#borro "ocupacion"

#borro repetidas
ocupaciones_boq <- ocupaciones_boq %>% distinct
ocupaciones_boq <- ocupaciones_boq[-1,]


# Calculate pairwise Levenshtein distances
distance_matrix_boq <- stringdistmatrix(ocupaciones_boq, ocupaciones_boq, method = "lv")


#find closest element
closest_elements_boq <- sapply(1:length(ocupaciones_boq), function(i) {
  min_distance_boq <- min(distance_matrix_boq[i, -i])
  closest_indices_boq <- which(distance_matrix_boq[i, -i] == min_distance_boq)
  closest_strings_boq <- ocupaciones_boq[-i][closest_indices_boq]
  return(closest_strings_boq)
})

#display results
for (i in 1:length(ocupaciones_boq)) {
  cat(paste("Closest elements to", ocupaciones_boq[i], "are:", closest_elements_boq[[i]], "\n"))
}



#VOLCÁN
#info colaboradores
info_colaboradores_vol <- read_excel(info_colaboradores_path, sheet = 8)

#borro las primeras filas que no tienen info
info_colaboradores_vol <- info_colaboradores_vol[-c(1,2,3),]
ocupaciones_vol <- info_colaboradores_vol[[6]]
ocupaciones_vol <- data.frame(Column_Name = ocupaciones_vol)
#borro "ocupacion"

#borro repetidas
ocupaciones_vol <- ocupaciones_vol %>% distinct
ocupaciones_vol <- ocupaciones_vol[-1,]


# Calculate pairwise Levenshtein distances
distance_matrix_vol <- stringdistmatrix(ocupaciones_vol, ocupaciones_vol, method = "lv")


#find closest element
closest_elements_vol <- sapply(1:length(ocupaciones_vol), function(i) {
  min_distance_vol <- min(distance_matrix_vol[i, -i])
  closest_indices_vol <- which(distance_matrix_vol[i, -i] == min_distance_vol)
  closest_strings_vol <- ocupaciones_vol[-i][closest_indices_vol]
  return(closest_strings_vol)
})

#display results
for (i in 1:length(ocupaciones_vol)) {
  cat(paste("Closest elements to", ocupaciones_vol[i], "are:", closest_elements_vol[[i]], "\n"))
}
