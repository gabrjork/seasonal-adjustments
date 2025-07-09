# ====================== Puxando dados da API do SIDRA - PIM-PF
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



# ==== Coletando dados da tabela 8880 - Headline da PMC com e sem ajuste sazonal
tabela_8880 <- get_sidra(8880, api = "/t/8880/n1/all/v/7170/p/last%2041/c11046/56734/d/v7170%205")
View(tabela_8880)

# Selecionando colunas relevantes
tabela_8880 <- tabela_8880 %>%
  select(
    data = "Mês (Código)",
    variable = "Tipos de índice",
    value = "Valor"
  ) %>%
  as_tibble()

View(tabela_8880)

pmc_8880 <- tabela_8880 %>%
  mutate(
    data = ym(data)
  ) %>%
  pivot_wider(
    names_from = variable,
    values_from = value
  )

View(pmc_8880)


# Transformando tabela 8880 para formato transposto (datas como colunas)
pmc_8880_transposta <- tabela_8880 %>%
  mutate(
    data = ym(data)
  ) %>%
  pivot_wider(
    names_from = data,
    values_from = value
  )

View(pmc_8880_transposta)


# ==== Coletando dados da tabela 8881 - Headline da PMC ampliada
tabela_8881 <- get_sidra(8881, api = "/t/8881/n1/all/v/7170/p/last%2041/c11046/56736/d/v7170%205")
View(tabela_8881)

# Selecionando colunas relevantes
tabela_8881 <- tabela_8881 %>%
  select(
    data = "Mês (Código)",
    variable = "Tipos de índice",
    value = "Valor"
  ) %>%
  as_tibble()

View(tabela_8881)

pmc_8881 <- tabela_8881 %>%
  mutate(
    data = ym(data)
  ) %>%
  pivot_wider(
    names_from = variable,
    values_from = value
  )

View(pmc_8881)

# Transformando tabela 8881 para formato transposto (datas como colunas)
pmc_8881_transposta <- tabela_8881 %>%
  mutate(
    data = ym(data)
  ) %>%
  pivot_wider(
    names_from = data,
    values_from = value
  )

View(pmc_8881_transposta)



# ==== Coletando dados da tabela 8883 - PIM-PF Atividades
tabela_8883 <- get_sidra(8883,
  api = "/t/8883/n1/all/v/7170/p/last%2041/c11046/56736/c85/all/d/v7170%205"
)
View(tabela_8883)

# Selecionando colunas relevantes
tabela_8883 <- tabela_8883 %>%
  select(
    data = "Mês (Código)",
    variable = "Atividades",
    value = "Valor"
  ) %>%
  as_tibble()

View(tabela_8883)

pmc_8883 <- tabela_8883 %>%
  mutate(
    data = ym(data)
  ) %>%
  pivot_wider(
    names_from = variable,
    values_from = value
  )

View(pmc_8883)

# Transformando tabela 8883 para formato transposto (datas como colunas)
pmc_8883_transposta <- tabela_8883 %>%
  mutate(
    data = ym(data)
  ) %>%
  pivot_wider(
    names_from = data,
    values_from = value
  )

pmc_8883_transposta <- na.omit(pmc_8883_transposta)
View(pmc_8883_transposta)


# Exportando para Excel
timestamp <- format(Sys.time(), "%Y%m%d")
write_xlsx(
  list(
    "8880" = pmc_8880_transposta,
    "8881" = pmc_8881_transposta,
    "8883" = pmc_8883_transposta
  ),
  path = paste0(timestamp, "_PMC.xlsx")
)
