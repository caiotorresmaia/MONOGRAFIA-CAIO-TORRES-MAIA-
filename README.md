# ============================================================
# APÊNDICE — CÓDIGO COMPLETO DE REPLICAÇÃO
#
# Título: Efeitos Empíricos da Taxa Selic e da Carga Tributária
#         do Simples Nacional sobre a Natalidade e Mortalidade
#         de Micro e Pequenas Empresas no Brasil (2009–2021)
#
# Autor:  Caio Torres Maia
# Curso:  Ciências Econômicas — FEA-RP / USP
# Ano:    2026
#
# ─────────────────────────────────────────────────────────────
# INSTRUÇÕES DE REPLICAÇÃO
#
# 1. Baixe e descompacte a pasta do projeto, mantendo a
#    seguinte estrutura de subpastas:
#
#       MONOGRAFIA CAIO/
#       ├── RODAR_TUDO.R              <- script mestre
#       ├── dados_tratados/
#       │   └── Planilha_Mestre_TCC_MPEs_FINAL_v2_1.xlsx
#       #       ├── scripts/
#           ├── 00_setup.R
#           ├── 00b_session_info.R
#           ├── 01_importacao.R
#           ├── 02_descritiva.R
#           ├── 03_modelos.R
#           ├── 03b_diagnosticos.R
#           ├── 04_robustez.R
#           ├── 05_tabelas_finais.R
#           └── tabelas_corrigidas.R    
#       └── resultados/               <- criado automaticamente
#
# 2. No RStudio, defina o diretório de trabalho:
#       Session -> Set Working Directory -> Choose Directory...
#    Aponte para a pasta raiz "MONOGRAFIA CAIO".
#
# 3. Execute RODAR_TUDO.R (Ctrl+Shift+Enter ou clique "Source").
#    Todos os pacotes são instalados automaticamente.
#    Todos os resultados (gráficos e tabelas) são salvos em
#    resultados/.
#
# ─────────────────────────────────────────────────────────────
# VERSÕES UTILIZADAS
#
# R versão 4.3 ou superior.
# Versões exatas dos pacotes são salvas automaticamente em
# resultados/session_info.txt e resultados/versoes_pacotes.csv
# ao final da execução.
#
# ─────────────────────────────────────────────────────────────
# NOTA METODOLÓGICA
#
# Especificação principal: OLS em séries temporais com
# transformação log-log (coeficientes = elasticidades) e
# erros-padrão corrigidos pelo método de Newey-West (HAC),
# que controla simultaneamente autocorrelação serial e
# heterocedasticidade. Variáveis independentes principais
# entram defasadas um período (t-1) para garantir precedência
# temporal e reduzir risco de endogeneidade. Testes de raiz
# unitária (ADF, Phillips-Perron e KPSS) são realizados antes
# das estimações. Período amostral: 2009–2021 (13 obs.).
# ============================================================


# ============================================================
# SCRIPT 00 — SETUP: INSTALAÇÃO DE PACOTES
# Executar apenas uma vez na primeira utilização.
# ============================================================

cat("=== INSTALANDO PACOTES ===\n")
cat("Isso pode levar alguns minutos na primeira vez.\n\n")

pkgs <- c(
  "tidyverse",   # manipulação e visualização de dados
  "readxl",      # leitura de arquivos .xlsx
  "janitor",     # clean_names() — limpeza automática de nomes de colunas
  "tseries",     # teste ADF + Jarque-Bera
  "lmtest",      # testes diagnósticos: Breusch-Godfrey, RESET, Breusch-Pagan
  "sandwich",    # erros Newey-West (HAC)
  "stargazer",   # tabelas de resultados formatadas
  "ggplot2",     # gráficos
  "car",         # VIF e regressões parciais (avPlots)
  "urca",        # testes Phillips-Perron e KPSS de raiz unitária
  "vars",        # modelo VAR bivariado (extensão opcional)
  "scales",      # formatação de eixos nos gráficos
  "broom",       # organiza resultados de modelos em data frames
  "purrr"        # iteração funcional (map_dfr etc.)
)

pkgs_novos <- pkgs[!pkgs %in% installed.packages()[, "Package"]]
if (length(pkgs_novos) > 0) {
  cat("Instalando:", paste(pkgs_novos, collapse = ", "), "\n")
  install.packages(pkgs_novos)
} else {
  cat("Todos os pacotes já estão instalados.\n")
}
pkgs <- c(
  "tidyverse", "readxl", "janitor", "tseries", "lmtest", "sandwich",
  "stargazer", "ggplot2", "car", "urca", "vars", "scales", "broom", "purrr",
  # ADICIONADOS — geração das tabelas Word corrigidas:
  "modelsummary", "flextable", "officer", "tibble"
)
cat("\nVerificando instalação...\n")
ok <- sapply(pkgs, requireNamespace, quietly = TRUE)
if (all(ok)) {
  cat("OK — todos os", length(pkgs), "pacotes disponíveis.\n")
} else {
  cat("ATENCAO — pacotes com problema:",
      paste(pkgs[!ok], collapse = ", "), "\n")
  cat("Tente instalar manualmente: install.packages('nome_do_pacote')\n")
}


# ============================================================
# SCRIPT 01 — IMPORTAÇÃO E TRATAMENTO DOS DADOS
#
# O que faz:
# 1. Le a planilha xlsx (skip=4 pula cabecalho decorativo;
#    linha 5 é o cabecalho real das colunas).
# 2. Padroniza nomes com janitor::clean_names().
# 3. Detecta colunas por nome parcial (robusto a variacoes
#    de grafia no Excel).
# 4. Filtra amostra 2008-2021, converte tipos e verifica
#    qualidade dos dados.
# 5. Cria variáveis transformadas (logs e defasagens).
# 6. Salva dados_principal e dados_modelo em .rds.
# ============================================================

pkgs_necessarios <- c("tidyverse", "readxl", "janitor", "tseries",
                      "lmtest", "sandwich", "stargazer", "car", "urca")
pkgs_faltando <- pkgs_necessarios[!pkgs_necessarios %in%
                                    installed.packages()[, "Package"]]
if (length(pkgs_faltando) > 0) {
  install.packages(pkgs_faltando)
}

library(tidyverse)
library(readxl)
library(janitor)

cat("========================================\n")
cat("SCRIPT 01 — IMPORTACAO E LIMPEZA\n")
cat("========================================\n\n")

# --- Caminho da planilha (relativo à raiz do projeto) ---
caminho_planilha <- "dados_tratados/Planilha_Mestre_TCC_MPEs_FINAL_v2_1.xlsx"

if (!file.exists(caminho_planilha)) {
  stop(
    "\n  ARQUIVO NAO ENCONTRADO: ", caminho_planilha,
    "\n  Verifique:",
    "\n    1. O working directory está em 'MONOGRAFIA CAIO'?",
    "\n       Session -> Set Working Directory -> Choose Directory...",
    "\n    2. A planilha está em dados_tratados/ ?"
  )
}

cat("Planilha encontrada:", caminho_planilha, "\n")

# --- Leitura ---
# skip=4: pula as 4 primeiras linhas de cabecalho decorativo do Excel.
cat("Lendo planilha com skip=4...\n")
dados_brutos_raw <- read_excel(caminho_planilha, skip = 4)
cat("Dimensoes brutas:", nrow(dados_brutos_raw), "x",
    ncol(dados_brutos_raw), "\n\n")

# --- Padronizacao de nomes ---
# clean_names() converte para snake_case sem acentos.
# Ex.: "Selic Media (% a.a.)" -> "selic_media_a_a"
dados_brutos <- dados_brutos_raw %>%
  janitor::clean_names()

# --- Deteccao robusta de colunas por nome parcial ---
# Em vez de usar nomes exatos (frágeis), buscamos por substring.
detectar_coluna <- function(df, padrao) {
  nomes   <- names(df)
  matches <- nomes[grepl(padrao, nomes, ignore.case = TRUE)]
  if (length(matches) == 0) return(NA_character_)
  return(matches[1])
}

mapa_colunas <- list(
  ano            = "^ano$",
  selic          = "selic",
  ipca           = "ipca",
  selic_real     = "selic.*real|real.*selic|fisher",
  pib            = "pib.*var|var.*pib|pib.*real|crescimento",
  pib_nominal    = "pib.*nominal|nominal.*pib",
  arrec_sn       = "arrec",
  carga_sn       = "carga",
  iiebr          = "iiebr|iie_br|ii_e|incerteza|iie",
  d_recessao     = "recess",
  d_pandemia     = "pandemia|pandemi",
  estoque        = "estoque|cempre",
  entradas       = "^entrada",
  saidas         = "sa.da|saida",
  nascimentos    = "nascim",
  tx_natalidade  = "tx.*nasc|nasc.*tx|natalidade|tx_nasc",
  tx_entrada     = "tx.*entrada|entrada.*tx",
  tx_mortalidade = "tx.*sa.da|sa.da.*tx|mortalidade|tx_sa",
  fonte          = "^fonte$",
  status         = "^status$"
)

colunas_detectadas <- sapply(mapa_colunas, function(padrao) {
  detectar_coluna(dados_brutos, padrao)
})

cat("Deteccao de colunas:\n")
for (nm in names(colunas_detectadas)) {
  val <- colunas_detectadas[nm]
  status_det <- if (is.na(val)) "NAO ENCONTRADA"
                else paste0("-> '", val, "'")
  cat(sprintf("  %-20s %s\n", nm, status_det))
}
cat("\n")

# Verificar colunas criticas
criticas <- c("ano", "selic", "ipca", "pib", "carga_sn",
              "tx_natalidade", "tx_mortalidade")
faltando <- criticas[is.na(colunas_detectadas[criticas])]

if (length(faltando) > 0) {
  cat("ATENCAO: colunas criticas nao encontradas:",
      paste(faltando, collapse = ", "), "\n")
  cat("Ajuste o mapa_colunas acima conforme os nomes do Excel.\n")
  stop("Corrija o mapa_colunas antes de continuar.")
}

# Renomear colunas
renomear <- colunas_detectadas[!is.na(colunas_detectadas)]
renomear_inv <- setNames(as.character(renomear), names(renomear))
renomear_inv <- renomear_inv[names(renomear_inv) != renomear_inv]

if (length(renomear_inv) > 0) {
  dados_brutos <- dados_brutos %>% rename(!!!renomear_inv)
}

# --- Converter tipos e filtrar ---
dados_brutos$ano <- suppressWarnings(as.numeric(dados_brutos$ano))
dados_brutos     <- dados_brutos %>% filter(!is.na(ano))

cat("Anos disponíveis:", paste(sort(dados_brutos$ano), collapse = ", "),
    "\n\n")

# --- Amostra principal 2008-2021 ---
# 2007: Simples Nacional vigorou parcialmente (LC 123/2006).
# 2022+: Mudanca metodologica do CEMPRE/IBGE gera quebra de serie.
dados_principal <- dados_brutos %>%
  filter(ano >= 2008 & ano <= 2021) %>%
  arrange(ano)

cat("Amostra principal (2008-2021):", nrow(dados_principal),
    "observacoes\n\n")

# --- Converter variaveis numericas ---
vars_numericas <- c("selic", "ipca", "pib", "carga_sn", "iiebr",
                    "tx_natalidade", "tx_mortalidade",
                    "d_recessao", "d_pandemia")

for (v in vars_numericas) {
  if (v %in% names(dados_principal)) {
    dados_principal[[v]] <- suppressWarnings(
      as.numeric(dados_principal[[v]])
    )
  }
}

if ("selic_real" %in% names(dados_principal)) {
  dados_principal$selic_real <- suppressWarnings(
    as.numeric(dados_principal$selic_real)
  )
}

# --- Verificacao de qualidade ---
cat("NAs por variavel:\n")
vars_verificar <- c("tx_mortalidade", "tx_natalidade", "selic",
                    "ipca", "pib", "carga_sn", "iiebr",
                    "d_recessao", "d_pandemia")
vars_verificar <- vars_verificar[vars_verificar %in% names(dados_principal)]

for (v in vars_verificar) {
  n_na <- sum(is.na(dados_principal[[v]]))
  flag <- if (n_na > 0) " ATENCAO" else " OK"
  cat(sprintf("  %-20s %d NA%s\n", v, n_na, flag))
}

# --- Criacao de variaveis transformadas ---
#
# Por que log-log? Os coeficientes estimados são elasticidades:
# "Se Selic sobe 1%, a mortalidade muda B%."
#
# Por que lag(1)? Empresas nao fecham no mesmo período em que o
# juro sobe. A defasagem garante precedencia temporal e reduz
# o risco de endogeneidade simultanea.
#
# IIEBr pode ser negativo -> log(|x| + 0.001).
# Selic real pode ser negativa -> deslocamento para dominio positivo.

cat("\nCriando variaveis transformadas (logs e defasagens)...\n")

dados_modelo <- dados_principal %>%
  arrange(ano) %>%
  mutate(
    # Variáveis dependentes
    ln_mort          = log(tx_mortalidade),
    ln_nat           = log(tx_natalidade),

    # Independentes principais: defasadas 1 período (t-1)
    ln_selic_L1      = log(lag(selic, 1)),
    ln_carga_L1      = log(lag(carga_sn, 1)),

    # Controles
    ln_ipca          = log(ipca),
    ln_iiebr         = if ("iiebr" %in% names(.data)) {
      log(abs(.data[["iiebr"]]) + 0.001)
    } else {
      NA_real_
    },

    # PIB em nível (já é variacao %, log nao faz sentido econometrico)

    # Dummy: Reforma do Simples Nacional (LC 155/2016, vigencia 2018+)
    d_reforma        = as.integer(ano >= 2018),

    # Selic real (para analise de robustez)
    ln_selic_real_L1 = if ("selic_real" %in% names(.)) {
      sr <- lag(selic_real, 1)
      log(sr + abs(min(sr, na.rm = TRUE)) + 0.1)
    } else {
      NA_real_
    }
  ) %>%
  # Perde 2008 pelo lag — correto e esperado
  filter(ano >= 2009)

cat("Observacoes no modelo (apos lag):", nrow(dados_modelo), "\n")
cat("Anos no modelo:", paste(dados_modelo$ano, collapse = ", "), "\n\n")

# --- Salvar em .rds ---
dir.create("dados_tratados", showWarnings = FALSE)
dir.create("resultados",     showWarnings = FALSE)

saveRDS(dados_principal, "dados_tratados/dados_principal.rds")
saveRDS(dados_modelo,    "dados_tratados/dados_modelo.rds")
saveRDS(dados_brutos,    "dados_tratados/dados_brutos_limpos.rds")

cat("OK: Dados salvos em dados_tratados/\n")
cat("========================================\n")
cat("SCRIPT 01 CONCLUIDO\n")
cat("========================================\n\n")


# ============================================================
# SCRIPT 02 — ANÁLISE DESCRITIVA E GRÁFICOS
#
# Produz (em resultados/):
# grafico_01_demografico.png      — Natalidade e Mortalidade
# grafico_02_independentes.png    — Selic e Carga SN/PIB
# grafico_03_pib.png              — Variacao real do PIB
# grafico_04_correlacoes.png      — Heatmap de correlacoes
# grafico_D_comovimentacao.png    — Series padronizadas
# tabela_1_descritivas.txt        — Estatisticas descritivas
# tabela_2_correlacoes.csv        — Matriz de correlacoes
# ============================================================

library(tidyverse)
library(ggplot2)
library(stargazer)

cat("========================================\n")
cat("SCRIPT 02 — ANALISE DESCRITIVA\n")
cat("========================================\n\n")

dados_principal <- readRDS("dados_tratados/dados_principal.rds")
dados_modelo    <- readRDS("dados_tratados/dados_modelo.rds")

cat("Dados carregados:", nrow(dados_principal), "obs (2008-2021)\n\n")

# --- Grafico 1: Natalidade e Mortalidade das MPEs ---
# Linhas verticais marcam eventos macroeconomicos relevantes:
# recessao de 2015-16, reforma do Simples em 2018, pandemia em 2020.
g1 <- ggplot(dados_principal, aes(x = ano)) +
  geom_line(aes(y = tx_mortalidade, color = "Mortalidade"),
            linewidth = 1.3) +
  geom_point(aes(y = tx_mortalidade, color = "Mortalidade"),
             size = 2.5) +
  geom_line(aes(y = tx_natalidade, color = "Natalidade"),
            linewidth = 1.3, linetype = "dashed") +
  geom_point(aes(y = tx_natalidade, color = "Natalidade"),
             size = 2.5) +
  geom_vline(xintercept = 2015, linetype = "dotted",
             color = "gray40", linewidth = 0.8) +
  geom_vline(xintercept = 2018, linetype = "dotted",
             color = "gray40", linewidth = 0.8) +
  geom_vline(xintercept = 2020, linetype = "dotted",
             color = "gray40", linewidth = 0.8) +
  annotate("text", x = 2015.1,
           y = max(dados_principal$tx_mortalidade, na.rm=TRUE) * 0.97,
           label = "Recessao", hjust = 0, size = 3, color = "gray40") +
  annotate("text", x = 2018.1,
           y = max(dados_principal$tx_mortalidade, na.rm=TRUE) * 0.97,
           label = "Reforma SN", hjust = 0, size = 3, color = "gray40") +
  annotate("text", x = 2020.1,
           y = max(dados_principal$tx_mortalidade, na.rm=TRUE) * 0.97,
           label = "Pandemia", hjust = 0, size = 3, color = "gray40") +
  scale_color_manual(values = c("Mortalidade" = "#1a3a5c",
                                "Natalidade"  = "#2980b9")) +
  scale_x_continuous(breaks = 2008:2021) +
  labs(
    title    = "Taxas de Natalidade e Mortalidade das MPEs — Brasil (2008–2021)",
    subtitle = "Fonte: IBGE/CEMPRE. Taxas calculadas sobre estoque total de empresas ativas.",
    x = "Ano", y = "Taxa (%)", color = NULL
  ) +
  theme_minimal(base_size = 12) +
  theme(legend.position  = "bottom",
        axis.text.x      = element_text(angle = 45, hjust = 1),
        panel.grid.minor = element_blank())

ggsave("resultados/grafico_01_demografico.png",
       plot = g1, width = 10, height = 5.5, dpi = 150)
cat("OK: Grafico 1 salvo: grafico_01_demografico.png\n")

# --- Grafico 2: Selic e Carga Tributaria (eixo duplo) ---
escala_carga <- max(dados_principal$selic, na.rm=TRUE) /
                max(dados_principal$carga_sn, na.rm=TRUE)

g2 <- ggplot(dados_principal, aes(x = ano)) +
  geom_line(aes(y = selic, color = "Selic (% a.a.)"),
            linewidth = 1.3) +
  geom_point(aes(y = selic, color = "Selic (% a.a.)"), size = 2.5) +
  geom_line(aes(y = carga_sn * escala_carga, color = "Carga SN/PIB (%)"),
            linewidth = 1.3, linetype = "dashed") +
  geom_point(aes(y = carga_sn * escala_carga, color = "Carga SN/PIB (%)"),
             size = 2.5) +
  scale_y_continuous(
    name     = "Selic (% a.a.)",
    sec.axis = sec_axis(~ . / escala_carga, name = "Carga SN / PIB (%)")
  ) +
  scale_color_manual(values = c("Selic (% a.a.)"  = "#c0392b",
                                "Carga SN/PIB (%)" = "#27ae60")) +
  scale_x_continuous(breaks = 2008:2021) +
  labs(
    title    = "Selic e Carga Tributaria do Simples Nacional — Brasil (2008–2021)",
    subtitle = "Fonte: BCB/SGS (Selic) e Receita Federal / IBGE (Carga SN/PIB).",
    x = "Ano", color = NULL
  ) +
  theme_minimal(base_size = 12) +
  theme(legend.position  = "bottom",
        axis.text.x      = element_text(angle = 45, hjust = 1),
        panel.grid.minor = element_blank())

ggsave("resultados/grafico_02_independentes.png",
       plot = g2, width = 10, height = 5.5, dpi = 150)
cat("OK: Grafico 2 salvo: grafico_02_independentes.png\n")

# --- Grafico 3: Variacao Real do PIB ---
g3 <- dados_principal %>%
  mutate(cor_barra = ifelse(pib >= 0, "Crescimento", "Recessao")) %>%
  ggplot(aes(x = ano, y = pib, fill = cor_barra)) +
  geom_col(width = 0.7, alpha = 0.85) +
  geom_hline(yintercept = 0, color = "black", linewidth = 0.5) +
  scale_fill_manual(values = c("Crescimento" = "#2980b9",
                               "Recessao"    = "#c0392b")) +
  scale_x_continuous(breaks = 2008:2021) +
  labs(
    title    = "Variacao Real do PIB — Brasil (2008–2021)",
    subtitle = "Fonte: IBGE/SCN via IPEA Data. Valores em % ao ano.",
    x = "Ano", y = "Variacao Real do PIB (%)", fill = NULL
  ) +
  theme_minimal(base_size = 12) +
  theme(legend.position  = "bottom",
        axis.text.x      = element_text(angle = 45, hjust = 1),
        panel.grid.minor = element_blank())

ggsave("resultados/grafico_03_pib.png",
       plot = g3, width = 10, height = 5.5, dpi = 150)
cat("OK: Grafico 3 salvo: grafico_03_pib.png\n")

# --- Grafico 4: Matriz de Correlacoes (heatmap) ---
vars_cor_nomes <- c("tx_mortalidade", "tx_natalidade",
                    "selic", "carga_sn", "pib", "ipca", "iiebr")
vars_cor <- dados_principal %>%
  dplyr::select(dplyr::any_of(vars_cor_nomes))

matriz_cor <- cor(vars_cor, use = "complete.obs")

nomes_legiveis_todos <- c(
  "tx_mortalidade" = "Tx. Mortalidade",
  "tx_natalidade"  = "Tx. Natalidade",
  "selic"          = "Selic",
  "carga_sn"       = "Carga SN/PIB",
  "pib"            = "PIB (var. %)",
  "ipca"           = "IPCA",
  "iiebr"          = "IIEBr"
)
nomes_legiveis <- nomes_legiveis_todos[
  names(nomes_legiveis_todos) %in% names(vars_cor)
]

cor_long <- as.data.frame(matriz_cor) %>%
  tibble::rownames_to_column("var1") %>%
  tidyr::pivot_longer(-var1, names_to = "var2", values_to = "r") %>%
  mutate(
    var1 = dplyr::coalesce(nomes_legiveis[var1], var1),
    var2 = dplyr::coalesce(nomes_legiveis[var2], var2)
  )

g4 <- ggplot(cor_long, aes(x = var1, y = var2, fill = r)) +
  geom_tile(color = "white", linewidth = 0.5) +
  geom_text(aes(label = round(r, 2)),
            size = 3.5, color = "black", fontface = "bold") +
  scale_fill_gradient2(
    low = "#c0392b", mid = "white", high = "#1a3a5c",
    midpoint = 0, limits = c(-1, 1), name = "r"
  ) +
  labs(
    title    = "Matriz de Correlacoes — Brasil (2008–2021)",
    subtitle = "Pearson. Vermelho = negativo. Azul = positivo.",
    x = NULL, y = NULL
  ) +
  theme_minimal(base_size = 11) +
  theme(axis.text.x     = element_text(angle = 45, hjust = 1),
        panel.grid      = element_blank(),
        legend.position = "right")

ggsave("resultados/grafico_04_correlacoes.png",
       plot = g4, width = 8, height = 7, dpi = 150)
cat("OK: Grafico 4 salvo: grafico_04_correlacoes.png\n")

# --- Grafico D: Co-movimentacao padronizada Selic x Mortalidade ---
# Argumento visual contra regressao espuria: as series se movem
# juntas no tempo, nao por coincidencia de tendencias comuns.
dados_norm <- dados_principal %>%
  mutate(
    selic_norm = (selic - mean(selic, na.rm=TRUE)) /
                  sd(selic, na.rm=TRUE),
    mort_norm  = (tx_mortalidade - mean(tx_mortalidade, na.rm=TRUE)) /
                  sd(tx_mortalidade, na.rm=TRUE)
  )

gD <- ggplot(dados_norm, aes(x = ano)) +
  geom_line(aes(y = mort_norm, color = "Mortalidade (padronizada)"),
            linewidth = 1.3) +
  geom_point(aes(y = mort_norm, color = "Mortalidade (padronizada)"),
             size = 2.5) +
  geom_line(aes(y = selic_norm, color = "Selic (padronizada)"),
            linewidth = 1.3, linetype = "dashed") +
  geom_point(aes(y = selic_norm, color = "Selic (padronizada)"),
             size = 2.5) +
  geom_hline(yintercept = 0, color = "gray70",
             linewidth = 0.5, linetype = "dotted") +
  scale_color_manual(values = c(
    "Mortalidade (padronizada)" = "#1a3a5c",
    "Selic (padronizada)"       = "#c0392b"
  )) +
  scale_x_continuous(breaks = 2008:2021) +
  labs(
    title    = "Co-movimentacao entre Selic e Mortalidade das MPEs (2008–2021)",
    subtitle = "Series padronizadas (media=0, dp=1). Argumento visual contra regressao espuria.",
    x = "Ano", y = "Desvios-padrao em relacao a media", color = NULL,
    caption = "Fonte: BCB/SGS e IBGE/CEMPRE. Elaboracao propria."
  ) +
  theme_minimal(base_size = 12) +
  theme(legend.position  = "bottom",
        axis.text.x      = element_text(angle = 45, hjust = 1),
        panel.grid.minor = element_blank(),
        plot.title       = element_text(face = "bold"),
        plot.caption     = element_text(color = "gray50", size = 9))

ggsave("resultados/grafico_D_comovimentacao.png",
       plot = gD, width = 10, height = 5.5, dpi = 180)
cat("OK: Grafico D salvo: grafico_D_comovimentacao.png\n")

# --- Tabela 1: Estatisticas Descritivas ---
vars_desc_nomes <- c("tx_mortalidade", "tx_natalidade",
                     "selic", "ipca", "carga_sn", "pib", "iiebr")
vars_desc <- dados_principal %>%
  dplyr::select(dplyr::any_of(vars_desc_nomes))

labels_desc_map <- c(
  "tx_mortalidade" = "Tx. Mortalidade (%)",
  "tx_natalidade"  = "Tx. Natalidade (%)",
  "selic"          = "Selic (% a.a.)",
  "ipca"           = "IPCA (% a.a.)",
  "carga_sn"       = "Carga SN/PIB (%)",
  "pib"            = "PIB Var. Real (%)",
  "iiebr"          = "IIEBr (var. media %)"
)
labels_desc <- unname(
  labels_desc_map[names(labels_desc_map) %in% names(vars_desc)]
)

stargazer(
  as.data.frame(vars_desc),
  type             = "text",
  title            = "Tabela 1 — Estatisticas Descritivas (2008–2021)",
  digits           = 3,
  summary.stat     = c("n", "mean", "sd", "min", "max"),
  covariate.labels = labels_desc,
  out = "resultados/tabela_1_descritivas.txt"
)
cat("OK: Tabela 1 salva: tabela_1_descritivas.txt\n")

# --- Tabela 2: Matriz de Correlacoes (triangulo inferior) ---
mat     <- round(cor(vars_cor, use = "complete.obs"), 3)
mat_tri <- mat
mat_tri[upper.tri(mat_tri)] <- NA

mat_char <- matrix("", nrow = nrow(mat_tri), ncol = ncol(mat_tri))
for (i in seq_len(nrow(mat_tri))) {
  for (j in seq_len(ncol(mat_tri))) {
    if (!is.na(mat_tri[i, j]))
      mat_char[i, j] <- as.character(mat_tri[i, j])
  }
}
labels_cor_presentes <- nomes_legiveis_todos[
  names(nomes_legiveis_todos) %in% names(vars_cor)
]
rownames(mat_char) <- paste0("(", seq_along(labels_cor_presentes), ") ",
                              unname(labels_cor_presentes))
colnames(mat_char) <- paste0("(", seq_along(labels_cor_presentes), ")")

write.csv(mat_char, "resultados/tabela_2_correlacoes.csv")
cat("OK: Tabela 2 salva: tabela_2_correlacoes.csv\n")

cat("\n========================================\n")
cat("SCRIPT 02 CONCLUIDO\n")
cat("========================================\n\n")


# ============================================================
# SCRIPT 03 — TESTES ADF E MODELOS OLS PRINCIPAIS
#
# Equacao estimada (forma log-log):
#
#   ln(Y_t) = a + B1*ln(Selic_{t-1})
#             + B2*ln(CargaSN_{t-1})
#             + g1*PIB_t + g2*ln(IPCA_t)
#             + g3*ln(IIEBr_t)
#             + d1*D_Recessao + d2*D_Pandemia
#             + d3*D_Reforma + e_t
#
# onde Y = taxa de mortalidade ou taxa de natalidade das MPEs.
#
# Produz (em resultados/):
# tabela_adf.csv                 — Testes ADF em nivel e diferenca
# grafico_07_estacionariedade.png
# grafico_05_ajuste_mortalidade.png
# grafico_06_ajuste_natalidade.png
# tabela_3_regressao_principal.txt
# tabela_4_diagnosticos.csv
# dados_tratados/modelos_ols.rds — modelos OLS salvos
# ============================================================

library(tseries)
library(lmtest)
library(sandwich)
library(stargazer)

cat("========================================\n")
cat("SCRIPT 03 — TESTES ADF E MODELOS OLS\n")
cat("========================================\n\n")

dados_modelo    <- readRDS("dados_tratados/dados_modelo.rds")
dados_principal <- readRDS("dados_tratados/dados_principal.rds")

cat("Dados carregados:", nrow(dados_modelo), "obs (2009-2021)\n\n")

# Flags de variaveis opcionais — definidas antes do ADF
tem_recessao <- "d_recessao" %in% names(dados_modelo) &&
                !all(is.na(dados_modelo$d_recessao))
tem_iiebr    <- "ln_iiebr"   %in% names(dados_modelo) &&
                !all(is.na(dados_modelo$ln_iiebr))
tem_ipca     <- "ln_ipca"    %in% names(dados_modelo) &&
                !all(is.na(dados_modelo$ln_ipca))

# --- Bloco 1: Testes ADF em nivel ---
# H0: serie TEM raiz unitaria (nao estacionaria)
# p < 0,05 -> rejeita H0 -> serie estacionaria -> usar em nivel
# p > 0,05 -> nao rejeita H0 -> considerar diferenciacao
# NOTA: com n = 13, o ADF tem baixo poder. Valores borderline
# (p entre 0,05 e 0,15) sao tratados como estacionarios quando
# os graficos nao mostram tendencia persistente. A correcao
# Newey-West mitiga o risco de inferencia invalida.

cat("========================================\n")
cat("BLOCO 1 — TESTES ADF EM NIVEL\n")
cat("H0: raiz unitaria  |  p<0,05 -> estacionaria\n")
cat("========================================\n\n")

series_nivel <- list(
  "ln(Mortalidade)"   = na.omit(dados_modelo$ln_mort),
  "ln(Natalidade)"    = na.omit(dados_modelo$ln_nat),
  "ln(Selic t-1)"     = na.omit(dados_modelo$ln_selic_L1),
  "ln(Carga SN t-1)"  = na.omit(dados_modelo$ln_carga_L1),
  "PIB (var. real %)" = na.omit(dados_modelo$pib)
)
if (tem_ipca)
  series_nivel[["ln(IPCA)"]]  <- na.omit(dados_modelo$ln_ipca)
if (tem_iiebr)
  series_nivel[["ln(IIEBr)"]] <- na.omit(dados_modelo$ln_iiebr)

adf_nivel <- data.frame(
  Variavel = character(), Estatistica = numeric(),
  P_valor  = numeric(),   Conclusao   = character(),
  stringsAsFactors = FALSE
)

for (nm in names(series_nivel)) {
  serie <- series_nivel[[nm]]
  if (length(serie) >= 5) {
    t  <- adf.test(serie)
    p  <- round(t$p.value, 4)
    s  <- round(t$statistic, 3)
    cl <- ifelse(p <= 0.05, "Estacionaria (usar em nivel)",
          ifelse(p <= 0.10, "Borderline (usar em nivel c/ cautela)",
                            "Nao estacionaria (considerar diferenca)"))
    adf_nivel <- rbind(adf_nivel,
      data.frame(Variavel=nm, Estatistica=s, P_valor=p, Conclusao=cl))
  }
}
print(adf_nivel, row.names = FALSE)

# ADF nas primeiras diferencas
cat("\n========================================\n")
cat("BLOCO 1B — ADF NAS PRIMEIRAS DIFERENCAS\n")
cat("p<0,05 aqui -> serie e I(1)\n")
cat("========================================\n\n")

series_diff <- list(
  "Dln(Mortalidade)"  = na.omit(diff(dados_modelo$ln_mort)),
  "Dln(Natalidade)"   = na.omit(diff(dados_modelo$ln_nat)),
  "Dln(Selic t-1)"    = na.omit(diff(na.omit(dados_modelo$ln_selic_L1))),
  "Dln(Carga SN t-1)" = na.omit(diff(na.omit(dados_modelo$ln_carga_L1))),
  "DPIB"              = na.omit(diff(dados_modelo$pib)),
  "Dln(IPCA)"         = na.omit(diff(dados_modelo$ln_ipca))
)

adf_diff <- data.frame(
  Variavel = character(), Estatistica = numeric(),
  P_valor  = numeric(),   Conclusao   = character(),
  stringsAsFactors = FALSE
)

for (nm in names(series_diff)) {
  serie <- series_diff[[nm]]
  if (length(serie) >= 4) {
    t  <- adf.test(serie)
    p  <- round(t$p.value, 4)
    s  <- round(t$statistic, 3)
    cl <- ifelse(p <= 0.05, "Estacionaria em diferenca — I(1)",
          ifelse(p <= 0.10, "Borderline", "Ainda nao estacionaria"))
    adf_diff <- rbind(adf_diff,
      data.frame(Variavel=nm, Estatistica=s, P_valor=p, Conclusao=cl))
  }
}
print(adf_diff, row.names = FALSE)

write.csv(rbind(
  cbind(Tipo = "Nivel",        adf_nivel),
  cbind(Tipo = "1a Diferenca", adf_diff)
), "resultados/tabela_adf.csv", row.names = FALSE)
cat("\nOK: Tabela ADF salva: resultados/tabela_adf.csv\n\n")

# --- Bloco 2: Inspecao visual das series ---
dados_long <- dados_modelo %>%
  dplyr::select(ano, ln_mort, ln_nat, ln_selic_L1,
                ln_carga_L1, pib, ln_ipca) %>%
  tidyr::pivot_longer(-ano, names_to = "variavel",
                      values_to = "valor") %>%
  mutate(variavel = dplyr::recode(variavel,
    "ln_mort"     = "ln(Mortalidade)",
    "ln_nat"      = "ln(Natalidade)",
    "ln_selic_L1" = "ln(Selic t-1)",
    "ln_carga_L1" = "ln(Carga SN t-1)",
    "pib"         = "PIB (var. %)",
    "ln_ipca"     = "ln(IPCA)"
  ))

g_adf <- ggplot(dados_long, aes(x = ano, y = valor)) +
  geom_line(color = "#1a3a5c", linewidth = 1.0) +
  geom_point(color = "#1a3a5c", size = 1.8) +
  geom_smooth(method = "lm", se = FALSE,
              color = "#c0392b", linewidth = 0.7, linetype = "dashed") +
  facet_wrap(~variavel, scales = "free_y", ncol = 2) +
  scale_x_continuous(breaks = c(2009, 2013, 2017, 2021)) +
  labs(
    title    = "Inspecao Visual das Series — Tendencia e Estacionariedade",
    subtitle = "Linha vermelha = tendencia linear. Serie estacionaria: sem tendencia persistente.",
    x = "Ano", y = "Valor"
  ) +
  theme_minimal(base_size = 11) +
  theme(panel.grid.minor = element_blank(),
        strip.text       = element_text(face = "bold"))

ggsave("resultados/grafico_07_estacionariedade.png",
       plot = g_adf, width = 10, height = 8, dpi = 150)
cat("OK: Grafico 7 salvo\n\n")

# --- Bloco 3: Modelos OLS completos ---
# Formulas construidas dinamicamente conforme variaveis disponíveis.
vars_controle <- c("pib")
if (tem_ipca)     vars_controle <- c(vars_controle, "ln_ipca")
if (tem_iiebr)    vars_controle <- c(vars_controle, "ln_iiebr")
if (tem_recessao) vars_controle <- c(vars_controle, "d_recessao")
vars_controle <- c(vars_controle, "d_pandemia", "d_reforma")

formula_mort_completa <- as.formula(
  paste("ln_mort ~ ln_selic_L1 + ln_carga_L1 +",
        paste(vars_controle, collapse = " + "))
)
formula_nat_completa <- as.formula(
  paste("ln_nat ~ ln_selic_L1 + ln_carga_L1 +",
        paste(vars_controle, collapse = " + "))
)

mod_mort_c <- lm(formula_mort_completa, data = dados_modelo)
mod_nat_c  <- lm(formula_nat_completa,  data = dados_modelo)

cat("--- MORTALIDADE (completo, com Newey-West) ---\n")
print(coeftest(mod_mort_c, vcov = NeweyWest(mod_mort_c)))
cat("\n--- NATALIDADE (completo, com Newey-West) ---\n")
print(coeftest(mod_nat_c,  vcov = NeweyWest(mod_nat_c)))

# --- Bloco 4: Modelos parcimoniososos (preferidos) ---
# Mantem apenas variaveis teoricamente centrais e relevantes.
# Modelo mais simples = interpretacao mais clara.
# Natalidade inclui PIB (ciclo economico afeta abertura de negocios).
cat("\n========================================\n")
cat("BLOCO 4 — MODELOS PARCIMONIOSOSOS (PREFERIDOS)\n")
cat("========================================\n\n")

mod_mort_p <- lm(
  ln_mort ~ ln_selic_L1 + ln_carga_L1 + d_pandemia + d_reforma,
  data = dados_modelo
)
mod_nat_p <- lm(
  ln_nat ~ ln_selic_L1 + ln_carga_L1 + pib + d_pandemia + d_reforma,
  data = dados_modelo
)

cat("--- MORTALIDADE (parcimonia, com Newey-West) ---\n")
print(coeftest(mod_mort_p, vcov = NeweyWest(mod_mort_p)))
cat("R2:", round(summary(mod_mort_p)$r.squared, 4),
    "| R2 Ajustado:", round(summary(mod_mort_p)$adj.r.squared, 4), "\n")

cat("\n--- NATALIDADE (parcimonia, com Newey-West) ---\n")
print(coeftest(mod_nat_p, vcov = NeweyWest(mod_nat_p)))
cat("R2:", round(summary(mod_nat_p)$r.squared, 4),
    "| R2 Ajustado:", round(summary(mod_nat_p)$adj.r.squared, 4), "\n\n")

# Salvar modelos para uso nos scripts seguintes
saveRDS(list(
  mod_mort_c = mod_mort_c,
  mod_nat_c  = mod_nat_c,
  mod_mort_p = mod_mort_p,
  mod_nat_p  = mod_nat_p
), "dados_tratados/modelos_ols.rds")
cat("OK: Modelos salvos em dados_tratados/modelos_ols.rds\n\n")

# --- Bloco 5: Testes de diagnostico ---
# Breusch-Godfrey: H0 = sem autocorrelacao. p > 0,05 -> ok
# RESET Ramsey:    H0 = modelo bem especificado. p > 0,05 -> ok
# Breusch-Pagan:   H0 = homocedasticidade. p > 0,05 -> ok
#                  (Newey-West corrige mesmo se rejeitado)

diagnosticos <- function(modelo, nome) {
  cat("---", nome, "---\n")
  bg  <- bgtest(modelo, order = 1)
  res <- resettest(modelo, power = 2:3)
  bp  <- bptest(modelo)
  cat(sprintf("  Breusch-Godfrey (autocorr.):  p = %.4f  %s\n",
      bg$p.value,
      ifelse(bg$p.value  > 0.05, "OK sem autocorr.", "possivel autocorr.")))
  cat(sprintf("  RESET Ramsey (especificacao): p = %.4f  %s\n",
      res$p.value,
      ifelse(res$p.value > 0.05, "OK bem especificado", "revisar")))
  cat(sprintf("  Breusch-Pagan (heteroced.):   p = %.4f  %s\n",
      bp$p.value,
      ifelse(bp$p.value  > 0.05, "OK homocedasticidade", "NW corrige")))
  invisible(list(bg=bg, reset=res, bp=bp))
}

d_mort <- diagnosticos(mod_mort_p, "MORTALIDADE (parcimonia)")
cat("\n")
d_nat  <- diagnosticos(mod_nat_p,  "NATALIDADE (parcimonia)")

diag_df <- data.frame(
  Modelo  = c(rep("Mortalidade", 3), rep("Natalidade", 3)),
  Teste   = rep(c("Breusch-Godfrey", "RESET Ramsey", "Breusch-Pagan"), 2),
  P_valor = c(
    round(d_mort$bg$p.value, 4),    round(d_mort$reset$p.value, 4),
    round(d_mort$bp$p.value, 4),    round(d_nat$bg$p.value, 4),
    round(d_nat$reset$p.value, 4),  round(d_nat$bp$p.value, 4)
  ),
  Conclusao = c(
    ifelse(d_mort$bg$p.value    > 0.05, "OK", "Atencao"),
    ifelse(d_mort$reset$p.value > 0.05, "OK", "Atencao"),
    ifelse(d_mort$bp$p.value    > 0.05, "OK", "Atencao"),
    ifelse(d_nat$bg$p.value     > 0.05, "OK", "Atencao"),
    ifelse(d_nat$reset$p.value  > 0.05, "OK", "Atencao"),
    ifelse(d_nat$bp$p.value     > 0.05, "OK", "Atencao")
  )
)
write.csv(diag_df, "resultados/tabela_4_diagnosticos.csv",
          row.names = FALSE)
cat("\nOK: Diagnosticos salvos: tabela_4_diagnosticos.csv\n")

# --- Bloco 6: Graficos observado vs. ajustado ---
dados_plot <- dados_modelo %>%
  mutate(
    mort_ajust = exp(fitted(mod_mort_p)),
    nat_ajust  = exp(fitted(mod_nat_p))
  )

g_fit_m <- ggplot(dados_plot, aes(x = ano)) +
  geom_line(aes(y = tx_mortalidade, color = "Observado"),
            linewidth = 1.2) +
  geom_line(aes(y = mort_ajust, color = "Ajustado"),
            linewidth = 1.2, linetype = "dashed") +
  geom_point(aes(y = tx_mortalidade, color = "Observado"), size = 2.5) +
  scale_color_manual(values = c("Observado" = "#1a3a5c",
                                "Ajustado"  = "#c0392b")) +
  scale_x_continuous(breaks = 2009:2021) +
  labs(
    title    = "Mortalidade: Observado vs. Ajustado pelo Modelo",
    subtitle = "OLS com erros Newey-West — Amostra 2009–2021",
    x = "Ano", y = "Taxa de Mortalidade (%)", color = NULL
  ) +
  theme_minimal(base_size = 12) +
  theme(legend.position  = "bottom",
        axis.text.x      = element_text(angle = 45, hjust = 1),
        panel.grid.minor = element_blank())

ggsave("resultados/grafico_05_ajuste_mortalidade.png",
       plot = g_fit_m, width = 10, height = 5.5, dpi = 150)

g_fit_n <- ggplot(dados_plot, aes(x = ano)) +
  geom_line(aes(y = tx_natalidade, color = "Observado"),
            linewidth = 1.2) +
  geom_line(aes(y = nat_ajust, color = "Ajustado"),
            linewidth = 1.2, linetype = "dashed") +
  geom_point(aes(y = tx_natalidade, color = "Observado"), size = 2.5) +
  scale_color_manual(values = c("Observado" = "#2980b9",
                                "Ajustado"  = "#27ae60")) +
  scale_x_continuous(breaks = 2009:2021) +
  labs(
    title    = "Natalidade: Observado vs. Ajustado pelo Modelo",
    subtitle = "OLS com erros Newey-West — Amostra 2009–2021",
    x = "Ano", y = "Taxa de Natalidade (%)", color = NULL
  ) +
  theme_minimal(base_size = 12) +
  theme(legend.position  = "bottom",
        axis.text.x      = element_text(angle = 45, hjust = 1),
        panel.grid.minor = element_blank())

ggsave("resultados/grafico_06_ajuste_natalidade.png",
       plot = g_fit_n, width = 10, height = 5.5, dpi = 150)
cat("OK: Graficos de ajuste salvos\n")

# --- Bloco 7: Tabela principal (stargazer) ---
stargazer(
  mod_mort_p, mod_nat_p,
  type  = "text",
  title = "Tabela 3 — OLS com Erros Newey-West HAC",
  dep.var.labels   = c("ln(Tx. Mortalidade)", "ln(Tx. Natalidade)"),
  covariate.labels = c(
    "ln(Selic) t-1", "ln(Carga SN/PIB) t-1",
    "PIB Var. Real", "D. Pandemia (2020-21)",
    "D. Reforma SN (2018+)", "Constante"
  ),
  se = list(
    sqrt(diag(NeweyWest(mod_mort_p))),
    sqrt(diag(NeweyWest(mod_nat_p)))
  ),
  star.cutoffs  = c(0.10, 0.05, 0.01),
  digits        = 3,
  no.space      = TRUE,
  omit.stat     = c("ser"),
  add.lines = list(
    c("Erros Padrao", "Newey-West HAC", "Newey-West HAC"),
    c("Periodo",      "2009-2021",      "2009-2021"),
    c("Observacoes",  "13",             "13")
  ),
  notes        = "*p<0,10; **p<0,05; ***p<0,01. Erros padrao HAC entre parenteses.",
  notes.append = FALSE,
  out = "resultados/tabela_3_regressao_principal.txt"
)
cat("OK: Tabela 3 salva: tabela_3_regressao_principal.txt\n")

cat("\n========================================\n")
cat("SCRIPT 03 CONCLUIDO\n")
cat("========================================\n\n")


# ============================================================
# SCRIPT 03b — DIAGNOSTICOS AVANCADOS
#
# Complementa o Script 03 com:
# - Testes PP e KPSS (alem do ADF)
# - ACF/PACF e Ljung-Box dos residuos
# - Q-Q plot e Jarque-Bera (normalidade)
# - Residuos no tempo e residuos x ajustados
# - Cook's Distance (observacoes influentes)
# - VIF (multicolinearidade)
# ============================================================

library(tseries)
library(lmtest)
library(sandwich)
library(ggplot2)
library(car)
library(urca)
library(purrr)

cat("========================================\n")
cat("SCRIPT 03b — DIAGNOSTICOS AVANCADOS\n")
cat("========================================\n\n")

dados_modelo <- readRDS("dados_tratados/dados_modelo.rds")
modelos      <- readRDS("dados_tratados/modelos_ols.rds")

mod_mort_p <- modelos$mod_mort_p
mod_nat_p  <- modelos$mod_nat_p
mod_mort_c <- modelos$mod_mort_c
mod_nat_c  <- modelos$mod_nat_c

# --- Bloco 1: ADF + PP + KPSS ---
# Diagnostico cruzado:
# ADF nao rejeita + KPSS rejeita -> quase certa raiz unitaria
# ADF rejeita + KPSS nao rejeita -> quase certa estacionariedade
# Ambos ambiguos -> usar graficos + conhecimento teorico

cat("BLOCO 1 — RAIZ UNITARIA: ADF + PP + KPSS\n\n")

series_ru <- list(
  "ln(Mortalidade)"  = na.omit(dados_modelo$ln_mort),
  "ln(Natalidade)"   = na.omit(dados_modelo$ln_nat),
  "ln(Selic t-1)"    = na.omit(dados_modelo$ln_selic_L1),
  "ln(Carga SN t-1)" = na.omit(dados_modelo$ln_carga_L1),
  "PIB (var. real)"  = na.omit(dados_modelo$pib),
  "ln(IPCA)"         = na.omit(dados_modelo$ln_ipca)
)

tbl_ru <- data.frame(
  Variavel = character(), ADF_p = numeric(),
  PP_p = character(), KPSS_critico = character(),
  Diagnostico = character(), stringsAsFactors = FALSE
)

for (nm in names(series_ru)) {
  s <- series_ru[[nm]]
  if (length(s) < 5) next

  adf_r  <- tryCatch(adf.test(s), error = function(e) NULL)
  pp_r   <- tryCatch({
    u <- ur.pp(s, type = "Z-tau", model = "constant")
    list(stat = u@teststat[1], crit5 = u@cval[2])
  }, error = function(e) NULL)
  kpss_r <- tryCatch({
    u <- ur.kpss(s, type = "mu", lags = "short")
    list(stat = u@teststat[1], crit5 = u@cval[2])
  }, error = function(e) NULL)

  adf_p    <- if (!is.null(adf_r)) round(adf_r$p.value, 4) else NA
  pp_p     <- if (!is.null(pp_r))
    ifelse(pp_r$stat < pp_r$crit5, "< 0.05", "> 0.05") else NA
  kpss_cr  <- if (!is.null(kpss_r))
    paste0(round(kpss_r$crit5, 4),
           ifelse(kpss_r$stat > kpss_r$crit5,
                  " [REJEITA H0]", " [nao rejeita]")) else NA

  adf_est  <- if (!is.na(adf_p)) (adf_p <= 0.10) else FALSE
  pp_est   <- if (!is.null(pp_r)) (pp_r$stat < pp_r$crit5) else NA
  kpss_est <- if (!is.null(kpss_r)) (kpss_r$stat <= kpss_r$crit5) else NA

  diag <- if (!is.na(pp_est) && !is.na(kpss_est)) {
    if (adf_est && pp_est && kpss_est)       "Estacionaria (todos)"
    else if (!adf_est && !pp_est && !kpss_est) "Raiz unitaria (todos)"
    else "Resultado misto — usar em nivel c/ cautela"
  } else "Inconclusivo"

  tbl_ru <- rbind(tbl_ru, data.frame(
    Variavel = nm, ADF_p = adf_p, PP_p = pp_p,
    KPSS_critico = kpss_cr, Diagnostico = diag,
    stringsAsFactors = FALSE
  ))
}

print(tbl_ru, row.names = FALSE)
write.csv(tbl_ru, "resultados/tabela_raiz_unitaria.csv",
          row.names = FALSE)
cat("\nOK: Tabela salva: tabela_raiz_unitaria.csv\n\n")

# --- Bloco 2: ACF/PACF dos residuos + Ljung-Box ---
res_mort <- residuals(mod_mort_p)
res_nat  <- residuals(mod_nat_p)

plot_acf_pacf <- function(resid, titulo, nome_arquivo) {
  n        <- length(resid)
  max_lag  <- min(n - 2, 8)
  banda    <- 1.96 / sqrt(n)
  acf_vals <- acf(resid,  lag.max = max_lag, plot = FALSE)$acf[-1]
  pacf_vals<- pacf(resid, lag.max = max_lag, plot = FALSE)$acf

  df_all <- rbind(
    data.frame(Lag = 1:length(acf_vals),  Valor = acf_vals,  Tipo = "ACF"),
    data.frame(Lag = 1:length(pacf_vals), Valor = pacf_vals, Tipo = "PACF")
  )

  g <- ggplot(df_all, aes(x = Lag, y = Valor)) +
    geom_hline(yintercept = 0, color = "black", linewidth = 0.4) +
    geom_hline(yintercept =  banda, color = "#c0392b",
               linetype = "dashed", linewidth = 0.7) +
    geom_hline(yintercept = -banda, color = "#c0392b",
               linetype = "dashed", linewidth = 0.7) +
    geom_segment(aes(xend = Lag, yend = 0),
                 color = "#1a3a5c", linewidth = 1.2) +
    geom_point(color = "#1a3a5c", size = 2.5) +
    facet_wrap(~Tipo, ncol = 2) +
    scale_x_continuous(breaks = 1:max_lag) +
    labs(
      title    = paste("ACF e PACF dos Residuos —", titulo),
      subtitle = "Linha vermelha = banda 95% (+/-1,96/raiz(n)). Pico fora = autocorrelacao.",
      x = "Defasagem (lag)", y = "Correlacao"
    ) +
    theme_minimal(base_size = 12) +
    theme(panel.grid.minor = element_blank(),
          strip.text = element_text(face = "bold"))

  ggsave(paste0("resultados/", nome_arquivo),
         plot = g, width = 10, height = 5, dpi = 150)
  cat("OK: Salvo:", nome_arquivo, "\n")
}

plot_acf_pacf(res_mort, "Mortalidade",
              "grafico_08a_acf_residuos_mortalidade.png")
plot_acf_pacf(res_nat,  "Natalidade",
              "grafico_08b_acf_residuos_natalidade.png")

lb_mort <- Box.test(res_mort, lag = 1, type = "Ljung-Box")
lb_nat  <- Box.test(res_nat,  lag = 1, type = "Ljung-Box")
cat(sprintf("\nLjung-Box — Mortalidade: p = %.4f  %s\n",
    lb_mort$p.value,
    ifelse(lb_mort$p.value > 0.05, "OK sem autocorrelacao",
           "autocorrelacao detectada")))
cat(sprintf("Ljung-Box — Natalidade:  p = %.4f  %s\n\n",
    lb_nat$p.value,
    ifelse(lb_nat$p.value > 0.05, "OK sem autocorrelacao",
           "autocorrelacao detectada")))

# --- Bloco 3: Q-Q plot + Jarque-Bera ---
# ATENCAO: com n=13, ambos os testes tem baixo poder.
df_qq <- data.frame(
  residuos = c(scale(res_mort), scale(res_nat)),
  Modelo   = rep(c("Mortalidade", "Natalidade"),
                 c(length(res_mort), length(res_nat)))
)

g_qq <- ggplot(df_qq, aes(sample = residuos)) +
  stat_qq(color = "#1a3a5c", size = 3, alpha = 0.9) +
  stat_qq_line(color = "#c0392b", linewidth = 1.0, linetype = "dashed") +
  facet_wrap(~Modelo, ncol = 2) +
  labs(
    title    = "Q-Q Plot — Normalidade dos Residuos (Padronizados)",
    subtitle = "Pontos na diagonal = distribuicao normal.",
    x = "Quantis Teoricos (Normal Padrao)", y = "Quantis Amostrais",
    caption = "Nota: com n=13, o teste tem baixo poder. Avaliacao visual essencial."
  ) +
  theme_minimal(base_size = 12) +
  theme(panel.grid.minor = element_blank(),
        strip.text = element_text(face = "bold"))

ggsave("resultados/grafico_09_qq_residuos.png",
       plot = g_qq, width = 9, height = 5, dpi = 150)
cat("OK: Grafico 9 salvo\n")

jb_mort <- tryCatch(jarque.bera.test(res_mort), error = function(e) NULL)
jb_nat  <- tryCatch(jarque.bera.test(res_nat),  error = function(e) NULL)

if (!is.null(jb_mort))
  cat(sprintf("Jarque-Bera — Mortalidade: p = %.4f  %s\n",
      jb_mort$p.value,
      ifelse(jb_mort$p.value > 0.05,
             "OK normalidade nao rejeitada", "nao normalidade")))
if (!is.null(jb_nat))
  cat(sprintf("Jarque-Bera — Natalidade:  p = %.4f  %s\n\n",
      jb_nat$p.value,
      ifelse(jb_nat$p.value > 0.05,
             "OK normalidade nao rejeitada", "nao normalidade")))

# --- Bloco 5: Cook's Distance ---
# Limiar: 4/(n-k-1). Obs. acima podem distorcer coeficientes.
# Com n=13, um unico outlier pode alterar o sinal de um coeficiente.

plotar_cook <- function(modelo, titulo, nome_arquivo) {
  n      <- nobs(modelo)
  k      <- length(coef(modelo)) - 1
  limiar <- 4 / (n - k - 1)

  df_inf <- data.frame(
    ano      = dados_modelo$ano[seq_len(n)],
    cook     = cooks.distance(modelo)
  ) %>% mutate(influente = cook > limiar)

  g_cook <- ggplot(df_inf, aes(x = ano, y = cook, fill = influente)) +
    geom_col(alpha = 0.85, width = 0.6) +
    geom_hline(yintercept = limiar, linetype = "dashed",
               color = "#c0392b", linewidth = 0.8) +
    geom_text(aes(label = ifelse(influente, ano, "")),
              vjust = -0.4, size = 3.2, color = "#c0392b",
              fontface = "bold") +
    scale_fill_manual(values = c("FALSE" = "#1a3a5c",
                                 "TRUE"  = "#e74c3c"), guide = "none") +
    scale_x_continuous(breaks = dados_modelo$ano[seq_len(n)]) +
    labs(
      title    = paste("Cook's Distance —", titulo),
      subtitle = "Barras vermelhas = obs. potencialmente influentes.",
      x = "Ano", y = "Cook's Distance"
    ) +
    theme_minimal(base_size = 12) +
    theme(axis.text.x = element_text(angle = 45, hjust = 1),
          panel.grid.minor = element_blank())

  ggsave(paste0("resultados/", nome_arquivo),
         plot = g_cook, width = 9, height = 5, dpi = 150)
  cat("OK: Salvo:", nome_arquivo, "\n")
  obs_inf <- df_inf$ano[df_inf$influente]
  if (length(obs_inf) > 0)
    cat("  ATENCAO: obs. influentes:", paste(obs_inf, collapse=", "), "\n")
  else
    cat("  Nenhuma obs. acima do limiar.\n")
}

plotar_cook(mod_mort_p, "Mortalidade", "grafico_12a_cook_mortalidade.png")
plotar_cook(mod_nat_p,  "Natalidade",  "grafico_12b_cook_natalidade.png")

# --- Bloco 7: VIF ---
# VIF = 1: sem multicolinearidade. > 5: moderada. > 10: severa.
calcular_vif <- function(modelo, nome) {
  vif_vals <- tryCatch(car::vif(modelo), error = function(e) NULL)
  if (is.null(vif_vals)) return(invisible(NULL))
  df_vif <- data.frame(
    Modelo   = nome,
    Variavel = names(vif_vals),
    VIF      = round(as.numeric(vif_vals), 3),
    Diagnostico = dplyr::case_when(
      as.numeric(vif_vals) < 2  ~ "Sem multicolinearidade",
      as.numeric(vif_vals) < 5  ~ "Moderada (toleravel)",
      as.numeric(vif_vals) < 10 ~ "Alta — interpretar c/ cautela",
      TRUE                      ~ "Severa"
    ),
    stringsAsFactors = FALSE
  )
  cat(nome, ":\n")
  print(df_vif[, -1], row.names = FALSE)
  cat("\n")
  return(df_vif)
}

vif_mort <- calcular_vif(mod_mort_p, "Mortalidade")
vif_nat  <- calcular_vif(mod_nat_p,  "Natalidade")
write.csv(rbind(vif_mort, vif_nat), "resultados/tabela_vif.csv",
          row.names = FALSE)
cat("OK: Tabela VIF salva\n\n")

cat("========================================\n")
cat("SCRIPT 03b CONCLUIDO\n")
cat("========================================\n\n")


# ============================================================
# SCRIPT 04 — ANALISE DE ROBUSTEZ
#
# Especificacoes testadas:
# 1. Base (2009-2021, parcimoniosa)
# 2. Sem pandemia (exclui 2020-2021)
# 3. Selic real no lugar da Selic nominal
# 4. Sem dummy de reforma
# 5. Completo (com IPCA, IIEBr, D.Recessao)
#
# Se os coeficientes mantem sinal e magnitude entre
# especificacoes, os resultados sao robustos (nao dependem
# de escolhas metodologicas especificas).
# ============================================================

library(lmtest)
library(sandwich)
library(stargazer)

cat("========================================\n")
cat("SCRIPT 04 — TESTES DE ROBUSTEZ\n")
cat("========================================\n\n")

dados_modelo <- readRDS("dados_tratados/dados_modelo.rds")
modelos      <- readRDS("dados_tratados/modelos_ols.rds")

mod_mort_p <- modelos$mod_mort_p
mod_nat_p  <- modelos$mod_nat_p
mod_mort_c <- modelos$mod_mort_c
mod_nat_c  <- modelos$mod_nat_c

# Especificacao 2: Sem pandemia (2009-2019)
# A pandemia gerou simultaneamente fechamentos por crise e
# formalizacoes via MEI por desemprego — efeito liquido ambiguo.
cat("--- Especificacao 2: Sem pandemia (2009-2019) ---\n")
dados_sem_pan <- dados_modelo %>% filter(ano <= 2019)

mod_mort_r1 <- lm(
  ln_mort ~ ln_selic_L1 + ln_carga_L1 + d_pandemia + d_reforma,
  data = dados_sem_pan
)
mod_nat_r1 <- lm(
  ln_nat ~ ln_selic_L1 + ln_carga_L1 + pib + d_pandemia + d_reforma,
  data = dados_sem_pan
)
cat("Mortalidade (sem pandemia):\n")
print(coeftest(mod_mort_r1, vcov = NeweyWest(mod_mort_r1)))
cat("Natalidade (sem pandemia):\n")
print(coeftest(mod_nat_r1,  vcov = NeweyWest(mod_nat_r1)))

# Especificacao 3: Selic real
# Testa se o custo real (descontado inflacao) ou nominal e o
# que importa para a decisao do empreendedor.
cat("\n--- Especificacao 3: Selic real ---\n")
tem_selic_real <- "ln_selic_real_L1" %in% names(dados_modelo) &&
                  !all(is.na(dados_modelo$ln_selic_real_L1))

if (tem_selic_real) {
  mod_mort_r2 <- lm(
    ln_mort ~ ln_selic_real_L1 + ln_carga_L1 + d_pandemia + d_reforma,
    data = dados_modelo
  )
  mod_nat_r2 <- lm(
    ln_nat ~ ln_selic_real_L1 + ln_carga_L1 + pib + d_pandemia + d_reforma,
    data = dados_modelo
  )
  cat("Mortalidade (Selic real):\n")
  print(coeftest(mod_mort_r2, vcov = NeweyWest(mod_mort_r2)))
  cat("Natalidade (Selic real):\n")
  print(coeftest(mod_nat_r2,  vcov = NeweyWest(mod_nat_r2)))
} else {
  cat("ATENCAO: ln_selic_real_L1 nao disponivel. Usando modelo base.\n")
  mod_mort_r2 <- mod_mort_p
  mod_nat_r2  <- mod_nat_p
}

# Especificacao 4: Sem dummy de reforma
# A reforma do Simples de 2018 pode captar efeitos estruturais
# independentes da carga tributaria.
cat("\n--- Especificacao 4: Sem dummy de reforma ---\n")
mod_mort_r3 <- lm(
  ln_mort ~ ln_selic_L1 + ln_carga_L1 + d_pandemia,
  data = dados_modelo
)
mod_nat_r3 <- lm(
  ln_nat ~ ln_selic_L1 + ln_carga_L1 + pib + d_pandemia,
  data = dados_modelo
)
cat("Mortalidade (sem D.Reforma):\n")
print(coeftest(mod_mort_r3, vcov = NeweyWest(mod_mort_r3)))
cat("Natalidade (sem D.Reforma):\n")
print(coeftest(mod_nat_r3,  vcov = NeweyWest(mod_nat_r3)))

# Tabelas de robustez
stargazer(
  mod_mort_p, mod_mort_r1, mod_mort_r2, mod_mort_r3, mod_mort_c,
  type  = "text",
  title = "Robustez — Equacao de Mortalidade",
  dep.var.labels = "ln(Tx. Mortalidade)",
  column.labels  = c("Base", "Sem Pandemia", "Selic Real",
                     "Sem D.Reforma", "Completo"),
  covariate.labels = c(
    "ln(Selic) t-1", "ln(Selic Real) t-1", "ln(Carga SN) t-1",
    "PIB Var. Real", "ln(IPCA)", "ln(IIEBr)", "D. Recessao",
    "D. Pandemia", "D. Reforma SN", "Constante"
  ),
  se = list(
    sqrt(diag(NeweyWest(mod_mort_p))),  sqrt(diag(NeweyWest(mod_mort_r1))),
    sqrt(diag(NeweyWest(mod_mort_r2))), sqrt(diag(NeweyWest(mod_mort_r3))),
    sqrt(diag(NeweyWest(mod_mort_c)))
  ),
  star.cutoffs = c(0.10, 0.05, 0.01), digits = 3, no.space = TRUE,
  omit.stat    = c("f", "ser"),
  notes        = "*p<0,10; **p<0,05; ***p<0,01. Erros Newey-West HAC.",
  notes.append = FALSE,
  out = "resultados/tabela_robustez_mortalidade.txt"
)
cat("OK: Tabela robustez mortalidade salva\n")

stargazer(
  mod_nat_p, mod_nat_r1, mod_nat_r2, mod_nat_r3, mod_nat_c,
  type  = "text",
  title = "Robustez — Equacao de Natalidade",
  dep.var.labels = "ln(Tx. Natalidade)",
  column.labels  = c("Base", "Sem Pandemia", "Selic Real",
                     "Sem D.Reforma", "Completo"),
  covariate.labels = c(
    "ln(Selic) t-1", "ln(Selic Real) t-1", "ln(Carga SN) t-1",
    "PIB Var. Real", "ln(IPCA)", "ln(IIEBr)", "D. Recessao",
    "D. Pandemia", "D. Reforma SN", "Constante"
  ),
  se = list(
    sqrt(diag(NeweyWest(mod_nat_p))),  sqrt(diag(NeweyWest(mod_nat_r1))),
    sqrt(diag(NeweyWest(mod_nat_r2))), sqrt(diag(NeweyWest(mod_nat_r3))),
    sqrt(diag(NeweyWest(mod_nat_c)))
  ),
  star.cutoffs = c(0.10, 0.05, 0.01), digits = 3, no.space = TRUE,
  omit.stat    = c("f", "ser"),
  notes        = "*p<0,10; **p<0,05; ***p<0,01. Erros Newey-West HAC.",
  notes.append = FALSE,
  out = "resultados/tabela_robustez_natalidade.txt"
)
cat("OK: Tabela robustez natalidade salva\n")

cat("\n========================================\n")
cat("SCRIPT 04 CONCLUIDO\n")
cat("========================================\n\n")


# ============================================================
# SCRIPT 05 — TABELAS FINAIS E RELATORIO CONSOLIDADO
#
# Consolida todos os resultados, gera relatorio em texto
# com avaliacao automatica das hipoteses e interpretacao
# das elasticidades estimadas.
# ============================================================

library(lmtest)
library(sandwich)
library(stargazer)

cat("========================================\n")
cat("SCRIPT 05 — TABELAS FINAIS\n")
cat("========================================\n\n")

dados_principal <- readRDS("dados_tratados/dados_principal.rds")
dados_modelo    <- readRDS("dados_tratados/dados_modelo.rds")
modelos         <- readRDS("dados_tratados/modelos_ols.rds")

mod_mort_p <- modelos$mod_mort_p
mod_nat_p  <- modelos$mod_nat_p

# Tabela 3 — OLS principal (versao final)
stargazer(
  mod_mort_p, mod_nat_p,
  type  = "text",
  title = "Tabela 3 — Impacto da Selic e Carga Tributaria sobre MPEs (OLS — Newey-West HAC)",
  dep.var.labels   = c("ln(Tx. Mortalidade)", "ln(Tx. Natalidade)"),
  covariate.labels = c(
    "ln(Selic) t-1", "ln(Carga SN/PIB) t-1",
    "PIB Var. Real", "D. Pandemia (2020-21)",
    "D. Reforma SN (2018+)", "Constante"
  ),
  se = list(
    sqrt(diag(NeweyWest(mod_mort_p))),
    sqrt(diag(NeweyWest(mod_nat_p)))
  ),
  star.cutoffs  = c(0.10, 0.05, 0.01),
  digits        = 3,
  no.space      = TRUE,
  omit.stat     = c("ser"),
  add.lines = list(
    c("Erros Padrao", "Newey-West HAC", "Newey-West HAC"),
    c("Periodo",      "2009-2021",      "2009-2021"),
    c("N",            "13",             "13")
  ),
  notes        = "*p<0,10; **p<0,05; ***p<0,01. Elasticidades (forma log-log).",
  notes.append = FALSE,
  out = "resultados/tabela_3_regressao_principal.txt"
)
cat("OK: tabela_3_regressao_principal.txt\n")

# Relatorio consolidado com avaliacao das hipoteses
sink("resultados/relatorio_final.txt")

cat("=======================================================\n")
cat("RELATORIO FINAL — TCC CAIO TORRES MAIA | FEA-RP / USP\n")
cat("Impacto da Selic e Carga Tributaria sobre MPEs no Brasil\n")
cat("Periodo: 2009-2021 | Metodo: OLS com Newey-West HAC\n")
cat("=======================================================\n\n")

coefs_m <- coeftest(mod_mort_p, vcov = NeweyWest(mod_mort_p))
coefs_n <- coeftest(mod_nat_p,  vcov = NeweyWest(mod_nat_p))

b_selic_m <- coefs_m["ln_selic_L1", "Estimate"]
p_selic_m <- coefs_m["ln_selic_L1", "Pr(>|t|)"]
b_carga_m <- coefs_m["ln_carga_L1", "Estimate"]
p_carga_m <- coefs_m["ln_carga_L1", "Pr(>|t|)"]
b_selic_n <- coefs_n["ln_selic_L1", "Estimate"]
p_selic_n <- coefs_n["ln_selic_L1", "Pr(>|t|)"]
b_carga_n <- coefs_n["ln_carga_L1", "Estimate"]
p_carga_n <- coefs_n["ln_carga_L1", "Pr(>|t|)"]

cat("RESULTADOS DAS HIPOTESES:\n\n")

cat(sprintf("H1 (Selic -> Mortalidade):  B = %+.3f, p = %.3f -> %s\n",
    b_selic_m, p_selic_m,
    if (p_selic_m < 0.10 && b_selic_m > 0) "CONFIRMADA (B>0, p<0,10)"
    else if (p_selic_m < 0.10) "SINAL CONTRARIO"
    else "NAO CONFIRMADA (p>=0,10)"))

cat(sprintf("H1 (Selic -> Natalidade):   B = %+.3f, p = %.3f -> %s\n",
    b_selic_n, p_selic_n,
    if (p_selic_n < 0.10 && b_selic_n < 0) "CONFIRMADA (B<0, p<0,10)"
    else if (p_selic_n < 0.10) "SINAL CONTRARIO"
    else "NAO CONFIRMADA (p>=0,10)"))

cat(sprintf("H2 (Carga -> Mortalidade):  B = %+.3f, p = %.3f -> %s\n",
    b_carga_m, p_carga_m,
    if (p_carga_m < 0.10 && b_carga_m > 0) "CONFIRMADA (B>0, p<0,10)"
    else if (p_carga_m < 0.10) "EFEITO PROTETOR DO SIMPLES (B<0)"
    else "NAO CONFIRMADA (p>=0,10)"))

cat(sprintf("H2 (Carga -> Natalidade):   B = %+.3f, p = %.3f -> %s\n",
    b_carga_n, p_carga_n,
    if (p_carga_n < 0.10 && b_carga_n < 0) "CONFIRMADA (B<0, p<0,10)"
    else if (p_carga_n < 0.10) "SINAL CONTRARIO"
    else "NAO CONFIRMADA (p>=0,10)"))

cat("\n\nINTERPRETACAO DAS ELASTICIDADES (forma log-log):\n")
cat(sprintf("  Um aumento de 1%% na Selic esta associado a %.2f%% na mortalidade.\n",
    b_selic_m))
cat(sprintf("  Um aumento de 1%% na Carga SN/PIB esta associado a %.2f%% na mortalidade.\n",
    b_carga_m))

cat(sprintf("\nMortalidade — R2 = %.3f  |  R2 Ajustado = %.3f\n",
    summary(mod_mort_p)$r.squared, summary(mod_mort_p)$adj.r.squared))
cat(sprintf("Natalidade  — R2 = %.3f  |  R2 Ajustado = %.3f\n",
    summary(mod_nat_p)$r.squared, summary(mod_nat_p)$adj.r.squared))

cat("\n=======================================================\n")
cat(sprintf("Gerado em: %s\n", format(Sys.time(), "%Y-%m-%d %H:%M:%S")))
sink()

cat("OK: relatorio_final.txt\n")

cat("\n========================================\n")
cat("SCRIPT 05 CONCLUIDO\n")
cat("========================================\n\n")

# ============================================================
# SCRIPT 06 — TABELAS WORD CORRIGIDAS (.docx)
#
# Gera as versões finais das Tabelas 2, 3, 4 e 6 em formato
# Word (.docx), usando os modelos reais (mod_mort_p / mod_nat_p)
# carregados de modelos_ols.rds.
#
# PROBLEMA CORRIGIDO: o script anterior (tabelas_monografia.R)
# utilizava dados simulados (set.seed(42)), o que produzia
# coeficientes inconsistentes com o texto da monografia.
#
# Saída: resultados/Tabelas_Monografia/
#   tabela2_correlacoes.docx
#   tabela3_regressao_principal.docx
#   tabela4_diagnosticos.docx
#   tabela6_vif.docx
# ============================================================
library(modelsummary)
library(flextable)
library(officer)
library(tibble)
library(car)

cat("========================================\n")
cat("SCRIPT 06 — TABELAS WORD CORRIGIDAS\n")
cat("========================================\n\n")

dados_modelo <- readRDS("dados_tratados/dados_modelo.rds")
modelos      <- readRDS("dados_tratados/modelos_ols.rds")
mod_mort_p   <- modelos$mod_mort_p
mod_nat_p    <- modelos$mod_nat_p

OUT <- "resultados/Tabelas_Monografia/"
dir.create(OUT, recursive = TRUE, showWarnings = FALSE)

hac <- function(m) NeweyWest(m, prewhite = FALSE, adjust = TRUE)

salvar_docx <- function(ft, arquivo, nota = NULL) {
  doc <- read_docx()
  doc <- body_add_flextable(doc, ft)
  if (!is.null(nota)) doc <- body_add_par(doc, nota, style = "Normal")
  print(doc, target = arquivo)
  cat("Salvo:", arquivo, "\n")
}

coef_map <- c(
  "(Intercept)"    = "Constante",
  "ln_selic_L1"    = "ln(Selic) t-1",
  "ln_carga_L1"    = "ln(Carga SN/PIB) t-1",
  "ln_carga_sn_L1" = "ln(Carga SN/PIB) t-1",
  "pib"            = "PIB Var. Real",
  "d_pandemia"     = "D. Pandemia (2020-21)",
  "d_reforma"      = "D. Reforma SN (2018+)"
)

gof_map <- tribble(
  ~raw,            ~clean,        ~fmt,
  "r.squared",     "R2",          3,
  "adj.r.squared", "R2 Ajustado", 3,
  "nobs",          "N",           0
)

estrelas <- c("*" = 0.10, "**" = 0.05, "***" = 0.01)

# Tabela 2 — Correlações
col_carga <- intersect(c("ln_carga_L1", "ln_carga_sn_L1"), names(dados_modelo))[1]
col_mort  <- intersect(c("ln_mort", "ln_mortalidade"),      names(dados_modelo))[1]
col_nat   <- intersect(c("ln_nat",  "ln_natalidade"),        names(dados_modelo))[1]

vars_corr <- dados_modelo[complete.cases(dados_modelo[,
                                                      c(col_mort, col_nat, "ln_selic_L1", col_carga, "pib")]), ]
vars_corr <- setNames(
  vars_corr[, c(col_mort, col_nat, "ln_selic_L1", col_carga, "pib")],
  c("ln(Mortalidade)", "ln(Natalidade)", "ln(Selic) t-1",
    "ln(Carga SN) t-1", "PIB (var.%)")
)
ft2 <- datasummary_correlation(vars_corr, output = "flextable") %>%
  theme_zebra() %>% autofit() %>% fontsize(size = 9, part = "all")
salvar_docx(ft2, paste0(OUT, "tabela2_correlacoes.docx"),
            nota = paste0("Correlações de Pearson das variáveis em logaritmo natural. N = ",
                          nrow(vars_corr), " obs. (2009-2021)."))

# Tabela 3 — Regressão principal
ft3 <- modelsummary(
  list("Mortalidade" = mod_mort_p, "Natalidade" = mod_nat_p),
  vcov     = list(hac(mod_mort_p), hac(mod_nat_p)),
  coef_map = coef_map, gof_map = gof_map, stars = estrelas,
  output   = "flextable"
) %>% theme_zebra() %>% autofit() %>% fontsize(size = 9, part = "all")
salvar_docx(ft3, paste0(OUT, "tabela3_regressao_principal.docx"),
            nota = "Erros-padrão HAC (Newey-West) entre parênteses. *** p<0,01, ** p<0,05, * p<0,10.")

# Tabela 4 — Diagnósticos
diag_modelo <- function(modelo, nome) {
  res <- residuals(modelo)
  bg  <- bgtest(modelo, order = 1)
  lb  <- Box.test(res, lag = 2, type = "Ljung-Box")
  rs  <- resettest(modelo, power = 2:3)
  bp  <- bptest(modelo)
  jb  <- suppressWarnings(jarque.bera.test(res))
  data.frame(
    Modelo = nome,
    Teste  = c("Breusch-Godfrey", "Ljung-Box",
               "RESET Ramsey", "Breusch-Pagan", "Jarque-Bera"),
    Estatistica = round(c(bg$statistic, lb$statistic,
                          rs$statistic, bp$statistic,
                          jb$statistic), 3),
    p_valor = round(c(bg$p.value, lb$p.value,
                      rs$p.value, bp$p.value, jb$p.value), 3),
    Resultado = c(
      ifelse(bg$p.value  > 0.05, "Nao rejeita H0", "Rejeita H0 (!)"),
      ifelse(lb$p.value  > 0.05, "Nao rejeita H0", "Rejeita H0 (!)"),
      ifelse(rs$p.value  > 0.05, "Nao rejeita H0", "Rejeita H0 (!)"),
      ifelse(bp$p.value  > 0.05, "Nao rejeita H0", "Rejeita H0 (!)"),
      ifelse(jb$p.value  > 0.05, "Nao rejeita H0", "Rejeita H0 (!)")
    ), check.names = FALSE
  )
}
tab_diag <- rbind(diag_modelo(mod_mort_p, "Mortalidade"),
                  diag_modelo(mod_nat_p,  "Natalidade"))
ft4 <- flextable(tab_diag) %>%
  merge_v(j = "Modelo") %>% theme_zebra() %>%
  autofit() %>% fontsize(size = 9, part = "all")
salvar_docx(ft4, paste0(OUT, "tabela4_diagnosticos.docx"),
            nota = "H0 = hipótese nula satisfeita. Nível de significância: 5%.")

# Tabela 6 — VIF
vif_m <- data.frame(Variavel = names(vif(mod_mort_p)),
                    `VIF Mortalidade` = round(vif(mod_mort_p), 3),
                    check.names = FALSE)
vif_n <- data.frame(Variavel = names(vif(mod_nat_p)),
                    `VIF Natalidade`  = round(vif(mod_nat_p), 3),
                    check.names = FALSE)
rotulos <- c("ln_selic_L1" = "ln(Selic) t-1",
             "ln_carga_L1" = "ln(Carga SN/PIB) t-1",
             "pib"         = "PIB (var. %)",
             "d_pandemia"  = "D. Pandemia",
             "d_reforma"   = "D. Reforma SN")
tab_vif <- merge(vif_m, vif_n, by = "Variavel", all = TRUE)
tab_vif$Variavel <- ifelse(tab_vif$Variavel %in% names(rotulos),
                           rotulos[tab_vif$Variavel], tab_vif$Variavel)
ft6 <- flextable(tab_vif) %>%
  theme_zebra() %>% autofit() %>% fontsize(size = 9, part = "all")
salvar_docx(ft6, paste0(OUT, "tabela6_vif.docx"),
            nota = "VIF < 5 = multicolinearidade tolerável. Modelos parcimoniosos.")

cat("\n========================================\n")
cat("SCRIPT 06 CONCLUIDO\n")
cat("Tabelas salvas em: resultados/Tabelas_Monografia/\n")
cat("========================================\n\n")
# ============================================================
# SCRIPT 00b — REGISTRO DO AMBIENTE DE REPLICACAO
#
# Padrao exigido pelas principais revistas de economia
# (AEA, Econometrica) para pacotes de replicacao.
# Salva versao do R e de todos os pacotes utilizados em:
# resultados/session_info.txt     — formato texto completo
# resultados/versoes_pacotes.csv  — tabela limpa
# ============================================================

cat("========================================\n")
cat("SCRIPT 00b — AMBIENTE DE REPLICACAO\n")
cat("========================================\n\n")

pkgs_analise <- c(
  "tidyverse", "readxl", "janitor",
  "tseries", "urca", "lmtest", "sandwich",
  "stargazer", "car", "ggplot2", "purrr"
)

for (p in pkgs_analise) {
  suppressPackageStartupMessages(library(p, character.only = TRUE))
}

# session_info.txt — formato texto completo
sink("resultados/session_info.txt")
cat("=======================================================\n")
cat("INFORMACOES DO AMBIENTE DE REPLICACAO\n")
cat("TCC: Caio Torres Maia | FEA-RP / USP | 2026\n")
cat("=======================================================\n\n")
cat("Data/hora da execucao:", format(Sys.time(), "%Y-%m-%d %H:%M:%S"),
    "\n\n")
print(sessionInfo())
sink()
cat("OK: Arquivo gerado: resultados/session_info.txt\n")

# versoes_pacotes.csv — tabela limpa para o apendice
pkgs_info <- installed.packages()[pkgs_analise, c("Package", "Version")]
pkgs_df   <- as.data.frame(pkgs_info, stringsAsFactors = FALSE)
pkgs_df$R_version <- paste0(R.version$major, ".", R.version$minor)

write.csv(pkgs_df, "resultados/versoes_pacotes.csv", row.names = FALSE)
cat("OK: Arquivo gerado: resultados/versoes_pacotes.csv\n\n")

cat("--- VERSAO DO R ---\n")
cat("R", paste0(R.version$major, ".", R.version$minor), "\n")
cat("Plataforma:", R.version$platform, "\n\n")

cat("--- PACOTES UTILIZADOS ---\n")
print(pkgs_df[, c("Package", "Version")], row.names = FALSE)

cat("\n========================================\n")
cat("ANALISE COMPLETA ENCERRADA COM SUCESSO\n")
cat("Todos os resultados em: resultados/\n")
cat("========================================\n")
