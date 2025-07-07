############# - Para o Caged: garantir que o arquivo CSV já foi tratado pelo VBA ##############
# Ajuste SAZONAL



library(zoo)
library(writexl)
library(seasonal)
library(x13binary)
library(dplyr)

# Definindo o workspace e puxando o arquivo csv
choose.files()
diretorio <- choose.dir()
setwd(diretorio)
getwd()

Cheio_sujo <- read.csv("2025_Cheio.csv", sep = ";", dec = ",")
View(Cheio_sujo)

# Print the column names to verify
print(names(Cheio_sujo))

# Print the first element of Data to debug
print(Cheio_sujo$Data[1])

# Caso haja colunas "NA", elas serão removidas
Cheio <- Cheio_sujo[, colSums(is.na(Cheio_sujo))
< nrow(Cheio_sujo)]

colnames(Cheio) <- c(
  "Data", "Admissoes", "Desligamentos", "Saldo",
  "Agro", "Industria", "Extrativa",
  "Transformacao", "Energia",
  "Utilities", "Construcao",
  "Comercio", "Servicos", "Transporte", "Alojamento",
  "InfoEOutros", "Comunicacao",
  "Financeiro", "Imobiliario", "ProfCientificas",
  "Admnistrativas",
  "AdmPublicaEOutros", "AdmPublica", "Educacao",
  "Saude", "Domesticos",
  "OutrosServicos", "Recreacao", "OutrasAtividades",
  "OrgInternacionais", "NaoIdentificados"
)

View(Cheio)
print(names(Cheio))

# Print do primeiro elemento da coluna Data para debug
print(Cheio$Data[1])

# O formato de data está como dd/mm/yyyy. O R precisa trabalhar com YYYY-MM-DD. Ajustando:
Cheio$Data <- as.Date(Cheio$Data, format = "%d-%m-%Y")
print(Cheio$Data[1])

# Criando um objeto time series (ts) para cada coluna
columns <- names(Cheio)[!names(Cheio) %in% c("Data")]
print(columns)
str(Cheio)

# Definindo como número todas as colunas, exceto Data
Cheio[columns] <- lapply(Cheio[columns], as.numeric)

# Primeira analise: Grandes atividades
Grandes_Atividades <- Cheio[, c(
  "Agro", "Industria", "Construcao",
  "Comercio", "Servicos", "NaoIdentificados"
)]

View(Grandes_Atividades)
str(Grandes_Atividades)

str(Grandes_Atividades$Data)

# Criando excel com os dados brutos
Grandes_Atividades_Bruto <- data.frame(
  Data = Cheio$Data,
  lapply(Grandes_Atividades, as.numeric) # Convert each time series to numeric
)
write_xlsx(Grandes_Atividades_Bruto, "202505_GrandesAtividades_Brutos.xlsx")

# Montand ts para Grandes Atividades
Grandes_Atividades_ts <- lapply(names(Grandes_Atividades), function(col) {
  ts(Grandes_Atividades[[col]],
    start = c(
      as.numeric(substr(Cheio$Data[1], 1, 4)),
      as.numeric(substr(Cheio$Data[1], 6, 7))
    ),
    frequency = 12
  )
})

names(Grandes_Atividades_ts) <- names(Grandes_Atividades)
# Atribui dinamicamente cada time series a uma variavel
for (col in names(Grandes_Atividades_ts)) {
  assign(col, Grandes_Atividades[[col]])
}

print(names(Grandes_Atividades_ts))
str(Grandes_Atividades_ts)

# Verify the class of the time series objects
lapply(Grandes_Atividades_ts, class)

par(mfrow = c(6, 1))
for (col in names(Grandes_Atividades)) {
  plot(Grandes_Atividades_ts[[col]], main = col, xlab = "Time", ylab = col, yaxt = "n")
  axis(2, at = pretty(Grandes_Atividades_ts[[col]]), labels = format(pretty(Grandes_Atividades_ts[[col]]), scientific = FALSE))
}

ls()

# Apply the SEASONAL x13-arima-seats with custom parameters
Grandes_Atividades_ts_ajuste <- lapply(Grandes_Atividades_ts, function(ts_obj) {
  seas(ts_obj,
    transform.function = "none",
    x11 = list(
      mode = "add",
      sigmalim = c(0.01, 1.0), # Less restrictive, less smoothing
      seasonalma = "s3x1" # Shorter moving average, less smoothing
    )
  )
})

# Dynamically assign each adjusted series to a variable with "_ajustado" suffix
for (name in names(Grandes_Atividades)) {
  assign(paste0(name, "_ajustado"), Grandes_Atividades_ts[[name]])
}

# Verify the adjusted variables are created
ls(pattern = "_ajustado")

# Print the class of each adjusted series
lapply(Grandes_Atividades_ts_ajuste, class)

# Extracting the final adjusted series
Grandes_Atividades_ts_ajuste_final <- lapply(Grandes_Atividades_ts_ajuste, final)

# Dynamically assign each final adjusted series to a variable with "_SA" suffix
for (name in names(Grandes_Atividades_ts_ajuste_final)) {
  assign(paste0(name, "_SA"), Grandes_Atividades_ts_ajuste_final[[name]])
}

# Creating plots for each adjusted variable
Nomes_ajustados <- ls(pattern = "_sa") # Get all adjusted variable names

# Set up the layout to display multiple plots in one image
par(mfrow = c(1, 1))

for (name in Nomes_ajustados) {
  plot(
    get(name),
    main = paste("Ajuste Sazonal de", name), # Dynamically set the title
    xlab = "Mês",
    ylab = name,
    yaxt = "n"
  )
  axis(2, at = pretty(get(name)), labels = format(pretty(get(name)), scientific = FALSE))
}

lapply(Grandes_Atividades_ts_ajuste_final, class)
# Exporting the adjusted data to Excel

# Extract dates from one of the time series (assuming all have the same time index)
dates <- format(as.yearmon(time(Grandes_Atividades_ts_ajuste_final[[1]])), "%Y-%m")

# Create a data frame with all final adjusted series
Grandes_Atividades_SA <- data.frame(
  Data = dates,
  lapply(Grandes_Atividades_ts_ajuste_final, as.numeric) # Convert each time series to numeric
)

Grandes_Atividades_SA$Total <- rowSums(Grandes_Atividades_SA[, !(names(Grandes_Atividades_SA) %in% "Data")])

View(Grandes_Atividades_SA)

# Save the data frame to an Excel file
write_xlsx(Grandes_Atividades_SA, "202505_GrandesAtividades_Ajustada.xlsx")





# Segunda analise: Sub atividades

print(names(Cheio))

Sub_Atividades <- Cheio[, c(
  "Agro", "Extrativa", "Transformacao",
  "Energia", "Utilities", "Construcao",
  "Comercio", "Transporte", "Alojamento",
  "Comunicacao", "Financeiro", "Imobiliario",
  "ProfCientificas", "Admnistrativas", "AdmPublica",
  "Educacao", "Saude", "Domesticos",
  "Recreacao", "OutrasAtividades",
  "OrgInternacionais", "NaoIdentificados"
)]
View(Sub_Atividades)
str(Sub_Atividades)

# Criando excel com os dados brutos
Sub_Atividades_Bruto <- data.frame(
  Data = Cheio$Data,
  lapply(Sub_Atividades, as.numeric) # Convert each time series to numeric
)
write_xlsx(Sub_Atividades_Bruto, "202505_SubAtividades_Brutos.xlsx")

# Montand ts para Grandes Atividades
Sub_Atividades_ts <- lapply(names(Sub_Atividades), function(col) {
  ts(Sub_Atividades[[col]],
    start = c(
      as.numeric(substr(Cheio$Data[1], 1, 4)),
      as.numeric(substr(Cheio$Data[1], 6, 7))
    ),
    frequency = 12
  )
})

names(Sub_Atividades_ts) <- names(Sub_Atividades)
# Atribui dinamicamente cada time series a uma variavel
for (col in names(Sub_Atividades_ts)) {
  assign(col, Sub_Atividades[[col]])
}

print(names(Sub_Atividades_ts))
str(Sub_Atividades_ts)

# Verify the class of the time series objects
lapply(Sub_Atividades_ts, class)

par(mfrow = c(6, 1))
for (col in names(Sub_Atividades)) {
  plot(Sub_Atividades_ts[[col]], main = col, xlab = "Time", ylab = col, yaxt = "n")
  axis(2,
    at = pretty(Sub_Atividades_ts[[col]]),
    labels = format(pretty(Sub_Atividades_ts[[col]]), scientific = FALSE)
  )
}

ls()

# Apply the SEASONAL x13-arima-seats with custom parameters
Sub_Atividades_ts_ajuste <- lapply(Sub_Atividades_ts, function(ts_obj) {
  seas(ts_obj,
    transform.function = "none",
    x11 = list(
      mode = "add",
      sigmalim = c(0.01, 1.0), # Less restrictive, less smoothing
      seasonalma = "s3x1" # Shorter moving average, less smoothing
    )
  )
})

# Dynamically assign each adjusted series to a variable with "_ajustado" suffix
for (name in names(Sub_Atividades)) {
  assign(paste0(name, "_ajustado"), Sub_Atividades_ts[[name]])
}

# Verify the adjusted variables are created
ls(pattern = "_ajustado")

# Print the class of each adjusted series
lapply(Sub_Atividades_ts_ajuste, class)

# Extracting the final adjusted series
Sub_Atividades_ts_ajuste_final <- lapply(Sub_Atividades_ts_ajuste, final)

# Dynamically assign each final adjusted series to a variable with "_SA" suffix
for (name in names(Sub_Atividades_ts_ajuste_final)) {
  assign(paste0(name, "_SA"), Sub_Atividades_ts_ajuste_final[[name]])
}

# Creating plots for each adjusted variable
Nomes_ajustados <- ls(pattern = "_sa") # Get all adjusted variable names

# Set up the layout to display multiple plots in one image
par(mfrow = c(1, 1))

for (name in Nomes_ajustados) {
  plot(
    get(name),
    main = paste("Ajuste Sazonal de", name), # Dynamically set the title
    xlab = "Mês",
    ylab = name,
    yaxt = "n"
  )
  axis(2, at = pretty(get(name)), labels = format(pretty(get(name)), scientific = FALSE))
}

lapply(Sub_Atividades_ts_ajuste_final, class)
# Exporting the adjusted data to Excel

# Extract dates from one of the time series (assuming all have the same time index)
dates <- format(as.yearmon(time(Sub_Atividades_ts_ajuste_final[[1]])), "%Y-%m")

# Create a data frame with all final adjusted series
Sub_Atividades_SA <- data.frame(
  Data = dates,
  lapply(Sub_Atividades_ts_ajuste_final, as.numeric) # Convert each time series to numeric
)

Sub_Atividades_SA$Total <- rowSums(Sub_Atividades_SA[, !(names(Sub_Atividades_SA) %in% "Data")])

View(Sub_Atividades_SA)

# Save the data frame to an Excel file
write_xlsx(Sub_Atividades_SA, "202505_SubAtividades_Ajustada.xlsx")
