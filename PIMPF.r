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

# ==== Coletando dados da tabela 8887 - PIM-PF Atividades
tabela_8887 <- get_sidra(8887, api = "/t/8887/n1/all/v/12607/p/last%2041/c543/129278,129283,129300,129301,129305/d/v12607%205")
View(tabela_8887)

# Selecionando colunas relevantes
tabela_8887 <- tabela_8887 %>%
  select(
    data = "Mês (Código)",
    variable = "Grandes categorias econômicas",
    value = "Valor"
  ) %>%
  as_tibble()

View(tabela_8887)

pimpf_8887 <- tabela_8887 %>%
  mutate(
    data = ym(data)
  ) %>%
  pivot_wider(
    names_from = variable,
    values_from = value
  )

View(pimpf_8887)

# Transformando tabela 8887 para formato transposto (datas como colunas)
pimpf_8887_transposta <- tabela_8887 %>%
  mutate(
    data = ym(data)
  ) %>%
  pivot_wider(
    names_from = data,
    values_from = value
  )

View(pimpf_8887_transposta)


# ==== Coletando dados da tabela 8888 - PIM-PF Atividades
tabela_8888 <- get_sidra(8888,
  api = "/t/8888/n1/all/v/12607/p/last%2041/c544/all/d/v12607%205"
)
View(tabela_8888)

# Selecionando colunas relevantes
tabela_8888 <- tabela_8888 %>%
  select(
    data = "Mês (Código)",
    variable = "Seções e atividades industriais (CNAE 2.0)",
    value = "Valor"
  ) %>%
  as_tibble()

View(tabela_8888)

pimpf_8888 <- tabela_8888 %>%
  mutate(
    data = ym(data)
  ) %>%
  pivot_wider(
    names_from = variable,
    values_from = value
  )

View(pimpf_8888)

# Transformando tabela 8888 para formato transposto (datas como colunas)
pimpf_8888_transposta <- tabela_8888 %>%
  mutate(
    data = ym(data)
  ) %>%
  pivot_wider(
    names_from = data,
    values_from = value
  )

View(pimpf_8888_transposta)


# Exportando para Excel
timestamp <- format(Sys.time(), "%Y%m%d")
write_xlsx(
  list(
    "8887" = pimpf_8887_transposta,
    "8888" = pimpf_8888_transposta
  ),
  path = paste0(timestamp, "_PIM_PF.xlsx")
)
