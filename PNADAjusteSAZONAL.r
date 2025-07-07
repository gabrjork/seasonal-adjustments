# Ajuste SAZONAL
library(sidrar)
library(zoo)
library(writexl)
library(seasonal)
library(x13binary)
library(dplyr)
library(tidyr)



choose.files() # workaround para forçar a função choose.dir() a funcionar, pois pode bugar
diretorio <- choose.dir()
setwd(diretorio)
getwd()

# Função para processar dados da API Sidra e transformar em formato largo
processar_dados_sidra <- function(tabela_codigo, variavel_codigo, categoria_mapping = NULL) {
  # Obtendo dados da API Sidra
  dados_suja <- get_sidra(tabela_codigo, variable = variavel_codigo, period = "all", classific = "all")

  print(paste("Processando tabela", tabela_codigo, "com variável", variavel_codigo))
  print("Estrutura original dos dados da API:")
  print(head(dados_suja))
  print("Nomes das colunas originais:")
  print(names(dados_suja))

  # Padronizando nomes das colunas
  colnames(dados_suja) <- make.names(colnames(dados_suja))

  # Convertendo trimestre móvel para data
  # Extraindo o último mês do trimestre móvel
  converter_trimestre_para_data <- function(trimestre_texto) {
    # Trimestre vem como "jan-fev-mar 2012"
    # Extraímos o ano
    ano <- as.numeric(substr(trimestre_texto, nchar(trimestre_texto) - 3, nchar(trimestre_texto)))

    # Extraímos os meses e pegamos o último
    meses_texto <- substr(trimestre_texto, 1, nchar(trimestre_texto) - 5)
    meses_split <- strsplit(meses_texto, "-")[[1]]
    ultimo_mes <- meses_split[length(meses_split)]

    # Mapeamento de meses
    meses_map <- c(
      "jan" = 1, "fev" = 2, "mar" = 3, "abr" = 4, "mai" = 5, "jun" = 6,
      "jul" = 7, "ago" = 8, "set" = 9, "out" = 10, "nov" = 11, "dez" = 12
    )

    mes_num <- meses_map[ultimo_mes]

    return(as.Date(paste0(ano, "-", sprintf("%02d", mes_num), "-01")))
  }

  dados_suja$Data <- sapply(dados_suja$Trimestre.Móvel, converter_trimestre_para_data)
  dados_suja$Data <- as.Date(dados_suja$Data, origin = "1970-01-01")

  # Identificando a coluna de categoria (SEMPRE usar nomes, não códigos)
  col_categoria <- NULL

  # Procurando por colunas específicas conhecidas (sempre sem ".Código.")
  colunas_categoria_possiveis <- c(
    "Condição.em.relação.à.força.de.trabalho.e.condição.de.ocupação",
    "Posição.na.ocupação",
    "Grupamento.de.atividade.no.trabalho.principal",
    "Grupamento.de.atividade",
    "Rendimento",
    "Sexo"
  )

  for (col in colunas_categoria_possiveis) {
    if (col %in% names(dados_suja)) {
      col_categoria <- col
      break
    }
  }

  # Se não encontrou, procura por padrões (excluindo colunas com "Código")
  if (is.na(col_categoria) || is.null(col_categoria)) {
    colunas_possiveis <- names(dados_suja)[grepl("Condição|Posição|Grupamento|Rendimento", names(dados_suja), ignore.case = TRUE)]
    # Filtra para excluir colunas de código
    colunas_sem_codigo <- colunas_possiveis[!grepl("Código", colunas_possiveis)]
    if (length(colunas_sem_codigo) > 0) {
      col_categoria <- colunas_sem_codigo[1]
    }
  }

  # Se ainda não encontrou e não é tabela de rendimento simples, usa a última coluna não-código
  if (is.na(col_categoria) || is.null(col_categoria)) {
    colunas_sem_codigo <- names(dados_suja)[!grepl("Código|Data|Trimestre|Variável|Valor|Nível|Unidade|Brasil", names(dados_suja))]
    if (length(colunas_sem_codigo) > 0) {
      col_categoria <- colunas_sem_codigo[1]
    } else {
      # Para tabela de rendimento, criar categoria única
      dados_suja$Categoria_Temp <- "Rendimento"
      col_categoria <- "Categoria_Temp"
    }
  }

  print(paste("Coluna de categoria identificada:", col_categoria))
  print("Categorias únicas encontradas:")
  print(unique(dados_suja[[col_categoria]]))

  # Aplicando mapeamento de categoria se fornecido
  if (!is.null(categoria_mapping)) {
    dados_suja$Categoria <- categoria_mapping[dados_suja[[col_categoria]]]
    dados_suja$Categoria[is.na(dados_suja$Categoria)] <- dados_suja[[col_categoria]][is.na(dados_suja$Categoria)]
  } else {
    dados_suja$Categoria <- dados_suja[[col_categoria]]
  }

  # Tratamento especial para tabelas sem categorias múltiplas (como rendimento)
  if (length(unique(dados_suja$Categoria)) == 1 || col_categoria == "Categoria_Temp") {
    # Para tabelas simples, criamos um dataframe direto
    dados_limpos <- dados_suja %>%
      select(Data, Valor) %>%
      rename(Rendimento = Valor) %>%
      arrange(Data)
  } else {
    # Transformando para formato largo para tabelas com múltiplas categorias
    dados_limpos <- dados_suja %>%
      select(Data, Categoria, Valor) %>%
      pivot_wider(names_from = Categoria, values_from = Valor) %>%
      arrange(Data)
  }

  # Removendo linhas com valores NA
  dados_limpos <- dados_limpos[complete.cases(dados_limpos), ]

  print("Dados transformados para formato largo:")
  print(head(dados_limpos))

  return(dados_limpos)
}

# Função auxiliar para extrair parâmetros de URL da API Sidra (para uso futuro)
extrair_parametros_sidra <- function(url_sidra) {
  # Exemplo de URL: https://apisidra.ibge.gov.br/values/t/6320/n1/all/v/4090/p/all/c11913/all
  parametros <- list()

  # Extrair tabela
  tabela_match <- regmatches(url_sidra, regexpr("t/[0-9]+", url_sidra))
  if (length(tabela_match) > 0) {
    parametros$tabela <- as.numeric(gsub("t/", "", tabela_match))
  }

  # Extrair variável
  variavel_match <- regmatches(url_sidra, regexpr("v/[0-9]+", url_sidra))
  if (length(variavel_match) > 0) {
    parametros$variavel <- as.numeric(gsub("v/", "", variavel_match))
  }

  return(parametros)
}

# Opção para usar dados da API ou CSVs já processados
usar_api <- TRUE # Usando apenas dados da API Sidra




#################### Primeiro ajuste: dados da PIA - TABELA 6318##########################

info_sidra(6318) # Verificando as informações da tabela 6318

if (usar_api) {
  # Mapeamento de categorias para nomes mais limpos
  categoria_mapping_6318 <- c(
    "Total" = "PIA",
    "Força de trabalho" = "ForcaTrab",
    "Força de trabalho - ocupada" = "PopOcupada",
    "Força de trabalho - desocupada" = "PopDesoc",
    "Fora da força de trabalho" = "NaoForcaTrab"
  )

  # Processando dados da API
  PIA_sem_ajuste <- processar_dados_sidra(6318, 1641, categoria_mapping_6318)
} else {
  # Carregando dados do CSV já processado
  PIA_sem_ajuste_suja <- read.csv("202505_6318.csv", sep = ";", dec = ",")

  # Caso haja colunas "NA", elas serão removidas
  PIA_sem_ajuste <- PIA_sem_ajuste_suja[, colSums(is.na(PIA_sem_ajuste_suja)) < nrow(PIA_sem_ajuste_suja)]

  colnames(PIA_sem_ajuste) <- c(
    "Data", "PIA", "ForcaTrab", "PopOcupada",
    "PopDesoc", "NaoForcaTrab"
  )

  # O formato de data está como dd/mm/yyyy. O R precisa trabalhar com YYYY-MM-DD. Ajustando:
  PIA_sem_ajuste$Data <- as.Date(PIA_sem_ajuste$Data, format = "%d/%m/%Y")
}

View(PIA_sem_ajuste)
head(PIA_sem_ajuste)
# Verificando os nomes das colunas após transformação
print("Nomes das colunas finais:")
print(names(PIA_sem_ajuste))

# Verificando a primeira data
print("Primeira data:")
print(PIA_sem_ajuste$Data[1])

# Criando um objeto time series (ts) para cada coluna (exceto Data)
columns <- names(PIA_sem_ajuste)[!names(PIA_sem_ajuste) %in% c("Data")]
print(columns)
str(PIA_sem_ajuste)

PIA_sem_ajuste <- as.data.frame(PIA_sem_ajuste) # Garantindo que seja um data frame
write_xlsx(PIA_sem_ajuste, "202505_6318_PIA_sem_ajuste.xlsx") # Exportando para Excel

# Montando ts para Grandes Atividades
PIA_ts <- lapply(columns, function(col) {
  ts(PIA_sem_ajuste[[col]],
    start = c(
      as.numeric(format(PIA_sem_ajuste$Data[1], "%Y")),
      as.numeric(format(PIA_sem_ajuste$Data[1], "%m"))
    ),
    frequency = 12
  )
})
names(PIA_ts) <- columns

# Atribui dinamicamente cada time series a uma variável
for (col in columns) {
  assign(col, PIA_ts[[col]])
}

print(names(PIA_ts))
str(PIA_ts)

# Verify the class of the time series objects
lapply(PIA_ts, class)

par(mfrow = c(length(columns), 1))
for (col in columns) {
  plot(PIA_ts[[col]], main = col, xlab = "Time", ylab = col, yaxt = "n")
  axis(2, at = pretty(PIA_ts[[col]]), labels = format(pretty(PIA_ts[[col]]), scientific = FALSE))
}

ls()

# Apply the SEASONAL x13-arima-seats with custom parameters
PIA_ts_ajuste <- lapply(PIA_ts, function(ts_obj) {
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
for (name in names(PIA_ts)) {
  assign(paste0(name, "_ajustado"), PIA_ts[[name]])
}

# Verify the adjusted variables are created
ls(pattern = "_ajustado")

# Print the class of each adjusted series
lapply(PIA_ts_ajuste, class)

# Extracting the final adjusted series
PIA_ts_ajuste_final <- lapply(PIA_ts_ajuste, final)

# Dynamically assign each final adjusted series to a variable with "_SA" suffix
for (name in names(PIA_ts_ajuste_final)) {
  assign(paste0(name, "_SA"), PIA_ts_ajuste_final[[name]])
}

lapply(PIA_ts_ajuste_final, class)
# Exporting the adjusted data to Excel

# Extract dates from one of the time series (assuming all have the same time index)
dates <- format(as.yearmon(time(PIA_ts_ajuste_final[[1]])), "%Y-%m")

# Create a data frame with all final adjusted series
PIA_SA <- data.frame(
  Data = dates,
  lapply(PIA_ts_ajuste_final, as.numeric) # Convert each time series to numeric
)

View(PIA_SA)

# Save the data frame to an Excel file
write_xlsx(PIA_SA, "202505_6318_PIA_SA.xlsx")





#################### CONFIGURAÇÃO: TODAS AS TABELAS USAM API SIDRA ####################

#################### Segundo ajuste: POSICAO - TABELA 6320   ##########################

if (usar_api) {
  # Mapeamento baseado nos nomes que você já usa
  categoria_mapping_6320 <- c(
    "Total" = "Ocupados",
    "Empregado" = "Empregado",
    "Empregado no setor privado" = "Privado",
    "Empregado no setor privado com carteira de trabalho assinada" = "PrivadoCLT",
    "Empregado no setor privado sem carteira de trabalho assinada" = "PrivadoSemCLT",
    "Trabalhador doméstico" = "Domestico",
    "Trabalhador doméstico com carteira de trabalho assinada" = "DomesticoCLT",
    "Trabalhador doméstico sem carteira de trabalho assinada" = "DomesticoSemCLT",
    "Empregado no setor público" = "Publico",
    "Empregado no setor público com carteira de trabalho assinada" = "PublicoCLT",
    "Empregado no setor público sem carteira de trabalho assinada" = "PublicoSemCLT",
    "Militar e bombeiro militar" = "Militar",
    "Empregador" = "Empregador",
    "Empregador com CNPJ" = "EmpregadorCNPJ",
    "Empregador sem CNPJ" = "EmpregadorSemCNPJ",
    "Conta própria" = "Autonomo",
    "Conta própria com CNPJ" = "AutonomoCNPJ",
    "Conta própria sem CNPJ" = "AutonomoSemCNPJ",
    "Trabalhador familiar auxiliar" = "Familiar"
  )

  # Processando dados da API com variável correta extraída da URL
  POSICAO_sem_ajuste <- processar_dados_sidra(6320, 4090, categoria_mapping_6320)
} else {
  # Código original para CSV (mantido como backup)
  POSICAO_sem_ajuste_suja <- read.csv("202505_6320.csv", sep = ";", dec = ",")
  POSICAO_sem_ajuste <- POSICAO_sem_ajuste_suja[, colSums(is.na(POSICAO_sem_ajuste_suja)) < nrow(POSICAO_sem_ajuste_suja)]

  colnames(POSICAO_sem_ajuste) <- c(
    "Data", "Ocupados", "Empregado", "Privado", "PrivadoCLT", "PrivadoSemCLT",
    "Domestico", "DomesticoCLT", "DomesticoSemCLT",
    "Publico", "PublicoCLT", "PublicoSemCLT", "Militar",
    "Empregador", "EmpregadorCNPJ", "EmpregadorSemCNPJ",
    "Autonomo", "AutonomoCNPJ", "AutonomoSemCNPJ", "Familiar"
  )

  POSICAO_sem_ajuste$Data <- as.Date(POSICAO_sem_ajuste$Data, format = "%d/%m/%Y")
}

View(POSICAO_sem_ajuste)
print("Nomes das colunas POSIÇÃO:")
print(names(POSICAO_sem_ajuste))
print("Primeira data POSIÇÃO:")
print(POSICAO_sem_ajuste$Data[1])

# Criando um objeto time series (ts) para cada coluna (exceto Data)
columns <- names(POSICAO_sem_ajuste)[!names(POSICAO_sem_ajuste) %in% c("Data")]
print(columns)
str(POSICAO_sem_ajuste)

POSICAO_sem_ajuste <- as.data.frame(POSICAO_sem_ajuste) # Garantindo que seja um data frame
write_xlsx(POSICAO_sem_ajuste, "202505_6320_POSICAO_sem_ajuste.xlsx") # Exportando para Excel

# Montando ts para Grandes Atividades
POSICAO_ts <- lapply(columns, function(col) {
  ts(POSICAO_sem_ajuste[[col]],
    start = c(
      as.numeric(format(POSICAO_sem_ajuste$Data[1], "%Y")),
      as.numeric(format(POSICAO_sem_ajuste$Data[1], "%m"))
    ),
    frequency = 12
  )
})
names(POSICAO_ts) <- columns

# Atribui dinamicamente cada time series a uma variável
for (col in columns) {
  assign(col, POSICAO_ts[[col]])
}

print(names(POSICAO_ts))
str(POSICAO_ts)

# Verify the class of the time series objects
lapply(POSICAO_ts, class)

par(mfrow = c(length(columns), 1))
for (col in columns) {
  plot(POSICAO_ts[[col]], main = col, xlab = "Time", ylab = col, yaxt = "n")
  axis(2, at = pretty(POSICAO_ts[[col]]), labels = format(pretty(POSICAO_ts[[col]]), scientific = FALSE))
}

ls()

# Apply the SEASONAL x13-arima-seats with custom parameters
POSICAO_ts_ajuste <- lapply(POSICAO_ts, function(ts_obj) {
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
for (name in names(POSICAO_ts)) {
  assign(paste0(name, "_ajustado"), POSICAO_ts[[name]])
}

# Verify the adjusted variables are created
ls(pattern = "_ajustado")

# Print the class of each adjusted series
lapply(POSICAO_ts_ajuste, class)

# Extracting the final adjusted series
POSICAO_ts_ajuste_final <- lapply(POSICAO_ts_ajuste, final)

# Dynamically assign each final adjusted series to a variable with "_SA" suffix
for (name in names(POSICAO_ts_ajuste_final)) {
  assign(paste0(name, "_SA"), POSICAO_ts_ajuste_final[[name]])
}

lapply(POSICAO_ts_ajuste_final, class)
# Exporting the adjusted data to Excel

# Extract dates from one of the time series (assuming all have the same time index)
dates <- format(as.yearmon(time(POSICAO_ts_ajuste_final[[1]])), "%Y-%m")

# Create a data frame with all final adjusted series
POSICAO_SA <- data.frame(
  Data = dates,
  lapply(POSICAO_ts_ajuste_final, as.numeric) # Convert each time series to numeric
)

View(POSICAO_SA)

# Save the data frame to an Excel file
write_xlsx(POSICAO_SA, "202505_6320_POSICAO_SA.xlsx")





#################### Terceiro ajuste: ATIVIDADE - TABELA 6323   ##########################

if (usar_api) {
  # Mapeamento baseado nos nomes que você já usa
  categoria_mapping_6323 <- c(
    "Total" = "Ocupados",
    "Agricultura, pecuária, produção florestal, pesca e aquicultura" = "Agro",
    "Indústria geral" = "Industria",
    "Construção" = "Construcao",
    "Comércio, reparação de veículos automotores e motocicletas" = "Comercio",
    "Transporte, armazenagem e correio" = "Transportes",
    "Alojamento e alimentação" = "Alojamento",
    "Informação, comunicação e atividades financeiras, imobiliárias, profissionais e administrativas" = "Informacao",
    "Administração pública, defesa e seguridade social, educação, saúde humana e serviços sociais" = "Adm",
    "Outros serviços" = "Outros",
    "Serviços domésticos" = "Domesticos"
  )

  # Processando dados da API com variável correta extraída da URL
  ATIVIDADE_sem_ajuste <- processar_dados_sidra(6323, 4090, categoria_mapping_6323)
} else {
  # Código original para CSV (mantido como backup)
  ATIVIDADE_sem_ajuste_suja <- read.csv("202505_6323.csv", sep = ";", dec = ",")
  ATIVIDADE_sem_ajuste <- ATIVIDADE_sem_ajuste_suja[, colSums(is.na(ATIVIDADE_sem_ajuste_suja)) < nrow(ATIVIDADE_sem_ajuste_suja)]

  colnames(ATIVIDADE_sem_ajuste) <- c(
    "Data", "Ocupados", "Agro", "Industria", "Construcao", "Comercio",
    "Transportes", "Alojamento", "Informacao",
    "Adm", "Outros", "Domesticos"
  )

  ATIVIDADE_sem_ajuste$Data <- as.Date(ATIVIDADE_sem_ajuste$Data, format = "%d/%m/%Y")
}

View(ATIVIDADE_sem_ajuste)
print("Nomes das colunas ATIVIDADE:")
print(names(ATIVIDADE_sem_ajuste))
print("Primeira data ATIVIDADE:")
print(ATIVIDADE_sem_ajuste$Data[1])

# Criando um objeto time series (ts) para cada coluna (exceto Data)
columns <- names(ATIVIDADE_sem_ajuste)[!names(ATIVIDADE_sem_ajuste) %in% c("Data")]
print(columns)
str(ATIVIDADE_sem_ajuste)

ATIVIDADE_sem_ajuste <- as.data.frame(ATIVIDADE_sem_ajuste) # Garantindo que seja um data frame
write_xlsx(ATIVIDADE_sem_ajuste, "202505_6323_ATIVIDADE_sem_ajuste.xlsx") # Exportando para Excel

# Montando ts para Grandes Atividades
ATIVIDADE_ts <- lapply(columns, function(col) {
  ts(ATIVIDADE_sem_ajuste[[col]],
    start = c(
      as.numeric(format(ATIVIDADE_sem_ajuste$Data[1], "%Y")),
      as.numeric(format(ATIVIDADE_sem_ajuste$Data[1], "%m"))
    ),
    frequency = 12
  )
})
names(ATIVIDADE_ts) <- columns

# Atribui dinamicamente cada time series a uma variável
for (col in columns) {
  assign(col, ATIVIDADE_ts[[col]])
}

print(names(ATIVIDADE_ts))
str(ATIVIDADE_ts)

# Verify the class of the time series objects
lapply(ATIVIDADE_ts, class)

par(mfrow = c(length(columns), 1))
for (col in columns) {
  plot(ATIVIDADE_ts[[col]], main = col, xlab = "Time", ylab = col, yaxt = "n")
  axis(2, at = pretty(ATIVIDADE_ts[[col]]), labels = format(pretty(ATIVIDADE_ts[[col]]), scientific = FALSE))
}

ls()

# Apply the SEASONAL x13-arima-seats with custom parameters
ATIVIDADE_ts_ajuste <- lapply(ATIVIDADE_ts, function(ts_obj) {
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
for (name in names(ATIVIDADE_ts)) {
  assign(paste0(name, "_ajustado"), ATIVIDADE_ts[[name]])
}

# Verify the adjusted variables are created
ls(pattern = "_ajustado")

# Print the class of each adjusted series
lapply(ATIVIDADE_ts_ajuste, class)

# Extracting the final adjusted series
ATIVIDADE_ts_ajuste_final <- lapply(ATIVIDADE_ts_ajuste, final)

# Dynamically assign each final adjusted series to a variable with "_SA" suffix
for (name in names(ATIVIDADE_ts_ajuste_final)) {
  assign(paste0(name, "_SA"), ATIVIDADE_ts_ajuste_final[[name]])
}

lapply(ATIVIDADE_ts_ajuste_final, class)
# Exporting the adjusted data to Excel

# Extract dates from one of the time series (assuming all have the same time index)
dates <- format(as.yearmon(time(ATIVIDADE_ts_ajuste_final[[1]])), "%Y-%m")

# Create a data frame with all final adjusted series
ATIVIDADE_SA <- data.frame(
  Data = dates,
  lapply(ATIVIDADE_ts_ajuste_final, as.numeric) # Convert each time series to numeric
)

View(ATIVIDADE_SA)

# Save the data frame to an Excel file
write_xlsx(ATIVIDADE_SA, "202505_6323_ATIVIDADE_SA.xlsx")




#################### Quarto ajuste: RENDIMENTO - TABELA 6387   ##########################

if (usar_api) {
  # Para rendimento, geralmente é uma variável simples
  categoria_mapping_6387 <- c(
    "Rendimento médio real habitual" = "Rendimento",
    "Rendimento médio real efetivo" = "Rendimento",
    "Total" = "Rendimento"
  )

  # Processando dados da API com variável correta extraída da URL
  RENDIMENTO_sem_ajuste <- processar_dados_sidra(6387, 5935, categoria_mapping_6387)
} else {
  # Código original para CSV (mantido como backup)
  RENDIMENTO_sem_ajuste_suja <- read.csv("202505_6387.csv", sep = ";", dec = ",")
  RENDIMENTO_sem_ajuste <- RENDIMENTO_sem_ajuste_suja[, colSums(is.na(RENDIMENTO_sem_ajuste_suja)) < nrow(RENDIMENTO_sem_ajuste_suja)]

  colnames(RENDIMENTO_sem_ajuste) <- c(
    "Data", "Rendimento"
  )

  RENDIMENTO_sem_ajuste$Data <- as.Date(RENDIMENTO_sem_ajuste$Data, format = "%d/%m/%Y")
}

View(RENDIMENTO_sem_ajuste)
print("Nomes das colunas RENDIMENTO:")
print(names(RENDIMENTO_sem_ajuste))
print("Primeira data RENDIMENTO:")
print(RENDIMENTO_sem_ajuste$Data[1])

# Criando um objeto time series (ts) para cada coluna (exceto Data)
columns <- names(RENDIMENTO_sem_ajuste)[!names(RENDIMENTO_sem_ajuste) %in% c("Data")]
print(columns)
str(RENDIMENTO_sem_ajuste)

RENDIMENTO_sem_ajuste <- as.data.frame(RENDIMENTO_sem_ajuste) # Garantindo que seja um data frame
write_xlsx(RENDIMENTO_sem_ajuste, "202505_6387_RENDIMENTO_sem_ajuste.xlsx") # Exportando para Excel

# Montando ts para Grandes RENDIMENTOs
RENDIMENTO_ts <- lapply(columns, function(col) {
  ts(RENDIMENTO_sem_ajuste[[col]],
    start = c(
      as.numeric(format(RENDIMENTO_sem_ajuste$Data[1], "%Y")),
      as.numeric(format(RENDIMENTO_sem_ajuste$Data[1], "%m"))
    ),
    frequency = 12
  )
})
names(RENDIMENTO_ts) <- columns

# Atribui dinamicamente cada time series a uma variável
for (col in columns) {
  assign(col, RENDIMENTO_ts[[col]])
}

print(names(RENDIMENTO_ts))
str(RENDIMENTO_ts)

# Verify the class of the time series objects
lapply(RENDIMENTO_ts, class)

par(mfrow = c(length(columns), 1))
for (col in columns) {
  plot(RENDIMENTO_ts[[col]], main = col, xlab = "Time", ylab = col, yaxt = "n")
  axis(2, at = pretty(RENDIMENTO_ts[[col]]), labels = format(pretty(RENDIMENTO_ts[[col]]), scientific = FALSE))
}

ls()

# Apply the SEASONAL x13-arima-seats with custom parameters
RENDIMENTO_ts_ajuste <- lapply(RENDIMENTO_ts, function(ts_obj) {
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
for (name in names(RENDIMENTO_ts)) {
  assign(paste0(name, "_ajustado"), RENDIMENTO_ts[[name]])
}

# Verify the adjusted variables are created
ls(pattern = "_ajustado")

# Print the class of each adjusted series
lapply(RENDIMENTO_ts_ajuste, class)

# Extracting the final adjusted series
RENDIMENTO_ts_ajuste_final <- lapply(RENDIMENTO_ts_ajuste, final)

# Dynamically assign each final adjusted series to a variable with "_SA" suffix
for (name in names(RENDIMENTO_ts_ajuste_final)) {
  assign(paste0(name, "_SA"), RENDIMENTO_ts_ajuste_final[[name]])
}

lapply(RENDIMENTO_ts_ajuste_final, class)
# Exporting the adjusted data to Excel

# Extract dates from one of the time series (assuming all have the same time index)
dates <- format(as.yearmon(time(RENDIMENTO_ts_ajuste_final[[1]])), "%Y-%m")

# Create a data frame with all final adjusted series
RENDIMENTO_SA <- data.frame(
  Data = dates,
  lapply(RENDIMENTO_ts_ajuste_final, as.numeric) # Convert each time series to numeric
)

View(RENDIMENTO_SA)

# Save the data frame to an Excel file
write_xlsx(RENDIMENTO_SA, "202505_6387_RENDIMENTO_SA.xlsx")
