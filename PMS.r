# ====================== Puxando dados da API do SIDRA - PMS
library(sidrar)
library(zoo)
library(writexl)
library(dplyr)
library(tidyr)
library(lubridate)

choose.files() # workaround para forçar a função choose.dir() a funcionar, pois pode bugar
diretorio <- choose.dir()
setwd(diretorio)
getwd()

# ==== Coletando dados da tabela 8688 - PMS

# Definindo o período dinâmico para a coleta de dados
data_inicio <- "202201"
data_fim <- format(Sys.Date(), "%Y%m")
periodo_completo <- paste0(data_inicio, "-", data_fim)

# Informar o usuário qual período está sendo usado
print(paste("Buscando dados para o período:", periodo_completo))

lista_de_classificacoes_total <- list(
  c11046 = 56726
)

# Fazer a chamada à API com o período dinâmico e fechado
tabela_8688 <- get_sidra(
  x = 8688, # Código da tabela (PMS - Índice de volume)
  period = periodo_completo, # Período dinâmico que criamos
  variable = 7168, # Variável "Índice de volume de serviços (2014=100)"
)

View(tabela_8688)


# Selecionando colunas relevantes
tabela_8688 <- tabela_8688 %>%
  select(
    data = "Mês (Código)",
    variable = "Atividades de serviços",
    tipos = "Tipos de índice (Código)",
    value = "Valor"
  ) %>%
  as_tibble()

View(tabela_8688)

head(tabela_8688)

# Removendo o índice na coluna "tipos" irrelevante (56725)
tabela_8688 <- tabela_8688 %>%
  filter(tipos != 56725)

# Define a coluna data propriamente e coloca os itens como colunas
pms_8688 <- tabela_8688 %>%
  mutate(
    data = ym(data)
  ) %>%
  pivot_wider(
    names_from = variable,
    values_from = value
  )

View(pms_8688)

# Transformando tabela 8688 para formato transposto (datas como colunas)
pms_8688_transposta <- tabela_8688 %>%
  mutate(
    data = ym(data)
  ) %>%
  pivot_wider(
    names_from = data,
    values_from = value
  )

# Omitido valores NA para evitar problemas na exportação
pms_8688_transposta <- na.omit(pms_8688_transposta)


# Exportando para Excel
timestamp <- format(Sys.time(), "%Y%m%d")
write_xlsx(
  list(
    "8688" = pms_8688_transposta
  ),
  path = paste0(timestamp, "_PMS.xlsx")
)

getwd()
