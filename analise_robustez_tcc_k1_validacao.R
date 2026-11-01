# =============================================================================
# TCC CEFET-MG / RAIS — ANÁLISE DE ROBUSTEZ DA CLUSTERIZAÇÃO
# Arquivo: analise_robustez_tcc.R
# Versão: 1.0.0
# Data: 2026-10-15
#
# Objetivo:
#   Avaliar a sensibilidade dos resultados de clusterização a escolhas analíticas
#   plausíveis, sem substituir o cenário principal do TCC.
#
# Regras-base preservadas do script principal:
#   * coortes de conclusão de 2021 a 2024;
#   * cenário estrito: ano_referencia > ano_conclusao;
#   * remun_media_nom arredondada para reais e corrigida para preços de 2024;
#   * fatores IPCA do projeto para 2021–2024;
#   * consolidação por egresso-ano: soma de remuneração, idade mediana,
#     tempo de emprego médio, nível máximo e modas determinísticas;
#   * Estudos 2 e 3 no nível individual, após consolidar pessoa-ano.
#
# Robustez implementada:
#   Estudo 2 (K-means): k=2,...,6; especificação-base e remuneração winsorizada
#   (1º/99º percentis); estabilidade entre inicializações por ARI, se mclust
#   estiver instalado.
#   Estudo 3 (K-prototypes): k=2,...,6 e lambda=0.5, 1 e 2; comparação de custo
#   interno e silhueta aproximada na mesma amostra fixa.
#
# Importante:
#   * Não interpreta clusters como efeitos causais.
#   * A winsorização é apenas uma análise alternativa, não substitui o resultado
#     principal e deve ser reportada como tal.
#   * A silhueta mista segue a métrica empregada no script revisado do Estudo 3:
#     distância euclidiana quadrática nas variáveis numéricas + lambda para
#     discordância ocupacional. Essa métrica deve ser interpretada com cautela.
#   * Não exporta IDs individuais nem microdados derivados.
# =============================================================================

options(stringsAsFactors = FALSE, scipen = 999)
set.seed(1234)

# ---- 0. Configuração ---------------------------------------------------------
ARQUIVO_BASE <- "base_cefet_mg_rais_2021_2024_egressos_com_1_curso_ate_2024.xlsx"
ABA_BASE <- 1
DIR_SAIDA <- file.path("output_clusterizacao", "analise_robustez")
ANO_BASE <- 2024L
COORTE_MIN <- 2021L
COORTE_MAX <- 2024L
ANOS_COM_FATOR <- 2021:2024
FATORES_IPCA <- data.frame(
  ano_referencia = 2021:2024,
  fator_2024 = c(1.1929197, 1.0916165, 1.0436746, 1.0000000)
)
SEED <- 1234L
NSTART_KMEANS <- 50L
NSTART_KPROTO <- 30L
K_AVALIADOS <- 2:6  # Estudo 2 (K-means)
K_AVALIADOS_KPROTO <- 1:6  # Estudo 3 (K-prototypes), inclui k=1 para validação
LAMBDAS <- c(0.5, 1, 2)
N_AMOSTRA_SILHUETA <- 1000L
# TRUE instala pacotes ausentes; FALSE interrompe e lista os pacotes necessários.
INSTALAR_PACOTES_AUSENTES <- TRUE

# ---- 1. Pacotes --------------------------------------------------------------
PACOTES <- c("readxl", "dplyr", "tidyr", "ggplot2", "cluster", "clustMixType")
faltantes <- PACOTES[!vapply(PACOTES, requireNamespace, quietly = TRUE,
                            FUN.VALUE = logical(1))]
if (length(faltantes) && INSTALAR_PACOTES_AUSENTES) {
  install.packages(faltantes, repos = "https://cloud.r-project.org",
                   dependencies = TRUE)
}
faltantes <- PACOTES[!vapply(PACOTES, requireNamespace, quietly = TRUE,
                            FUN.VALUE = logical(1))]
if (length(faltantes)) {
  stop("Instale os pacotes antes de executar: ",
       paste(faltantes, collapse = ", "))
}
suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(tidyr)
  library(ggplot2)
  library(cluster)
  library(clustMixType)
})
TEM_MCLUST <- requireNamespace("mclust", quietly = TRUE)

# ---- 2. Localizar base e preparar saídas -------------------------------------
resolver_arquivo <- function(arq) {
  candidatos <- c(arq, file.path("..", arq),
                  file.path("data", arq), file.path("dados", arq))
  encontrados <- candidatos[file.exists(candidatos)]
  if (length(encontrados)) return(normalizePath(encontrados[1]))
  stop("Não encontrei a base: ", arq,
       "\nColoque o XLSX na pasta de trabalho, em ./data ou ./dados, ",
       "ou edite ARQUIVO_BASE no início do script.")
}
CAMINHO_BASE <- resolver_arquivo(ARQUIVO_BASE)
dir.create(DIR_SAIDA, recursive = TRUE, showWarnings = FALSE)
DIR_TABELAS <- file.path(DIR_SAIDA, "tabelas")
DIR_FIGURAS <- file.path(DIR_SAIDA, "figuras")
DIR_AUDITORIA <- file.path(DIR_SAIDA, "auditoria")
invisible(lapply(c(DIR_TABELAS, DIR_FIGURAS, DIR_AUDITORIA),
                 dir.create, recursive = TRUE, showWarnings = FALSE))

ARQUIVO_LOG <- file.path(DIR_AUDITORIA,
                         paste0("analise_robustez_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".log"))
con_log <- file(ARQUIVO_LOG, open = "wt", encoding = "UTF-8")
log_msg <- function(..., nivel = "INFO") {
  msg <- paste0("[", format(Sys.time(), "%Y-%m-%d %H:%M:%S"), "] [",
                nivel, "] ", paste(..., collapse = ""))
  cat(msg, "\n")
  cat(msg, "\n", file = con_log)
}
log_msg("Início da análise de robustez.")
log_msg("Base: ", CAMINHO_BASE)
log_msg("Diretório de saída: ", normalizePath(DIR_SAIDA, mustWork = FALSE))
log_msg("Versão do R: ", R.version.string)
log_msg("mclust disponível para ARI: ", TEM_MCLUST)

exportar_csv <- function(x, nome) {
  write.csv(x, file.path(DIR_TABELAS, nome), row.names = FALSE,
            na = "", fileEncoding = "UTF-8")
}
exigir_colunas <- function(dados, cols) {
  ausentes <- setdiff(cols, names(dados))
  if (length(ausentes)) stop("Colunas ausentes no XLSX: ",
                             paste(ausentes, collapse = ", "))
}

# ---- 3. Leitura, filtros e consolidação conforme script principal ------------
abas <- excel_sheets(CAMINHO_BASE)
if (length(abas) < ABA_BASE) stop("ABA_BASE não existe no arquivo.")
dados <- read_excel(CAMINHO_BASE, sheet = ABA_BASE, .name_repair = "unique")
colunas <- c("id_aluno", "id_curso", "ano_conclusao", "remun_media_nom",
             "tempo_emprego_m", "ano_referencia", "desc_ocupacao", "idade",
             "campus", "nivel_categoria", "valor_nivel_categoria")
exigir_colunas(dados, colunas)

modo_deterministico <- function(x) {
  x <- trimws(as.character(x))
  x <- x[!is.na(x) & nzchar(x)]
  if (!length(x)) return(NA_character_)
  tab <- table(x)
  sort(names(tab)[tab == max(tab)])[1]
}
safe_mediana <- function(x) if (all(is.na(x))) NA_real_ else median(x, na.rm = TRUE)
safe_media <- function(x) if (all(is.na(x))) NA_real_ else mean(x, na.rm = TRUE)
safe_maximo <- function(x) if (all(is.na(x))) NA_real_ else max(x, na.rm = TRUE)

base <- dados %>%
  mutate(
    ano_referencia = as.integer(ano_referencia),
    ano_conclusao = as.integer(ano_conclusao),
    remun_media_nom = as.numeric(remun_media_nom),
    tempo_emprego_m = as.numeric(tempo_emprego_m),
    idade = as.numeric(idade),
    valor_nivel_categoria = as.numeric(valor_nivel_categoria),
    remun_nominal_reais = round(remun_media_nom),
    desc_ocupacao = trimws(as.character(desc_ocupacao)),
    campus = trimws(as.character(campus)),
    nivel_categoria = trimws(as.character(nivel_categoria))
  ) %>%
  filter(!is.na(id_aluno), !is.na(ano_conclusao),
         ano_conclusao >= COORTE_MIN, ano_conclusao <= COORTE_MAX,
         !is.na(remun_media_nom), remun_media_nom > 0) %>%
  group_by(id_aluno) %>%
  filter(n_distinct(id_curso, na.rm = TRUE) <= 1) %>%
  ungroup() %>%
  filter(ano_referencia %in% ANOS_COM_FATOR, ano_referencia > ano_conclusao) %>%
  left_join(FATORES_IPCA, by = "ano_referencia") %>%
  mutate(remuneracao_reais = remun_nominal_reais * fator_2024)

if (anyNA(base$fator_2024)) {
  stop("Há registros sem fator IPCA. Verifique anos e fatores; nenhum fator foi imputado.")
}
if (!nrow(base)) stop("Nenhum registro restou após os filtros estritos.")

log_msg("Registros RAIS após filtros estritos: ", nrow(base))
log_msg("Egressos após filtros estritos: ", dplyr::n_distinct(base$id_aluno))
log_msg("Anos observados: ", paste(sort(unique(base$ano_referencia)), collapse = ", "))

pessoa_ano <- base %>%
  group_by(id_aluno, ano_referencia) %>%
  summarise(
    ano_conclusao = min(ano_conclusao, na.rm = TRUE),
    remuneracao_reais = sum(remuneracao_reais, na.rm = TRUE),
    idade = safe_mediana(idade),
    tempo_emprego_m = safe_media(tempo_emprego_m),
    valor_nivel_categoria = safe_maximo(valor_nivel_categoria),
    campus = modo_deterministico(campus),
    nivel_categoria = modo_deterministico(nivel_categoria),
    desc_ocupacao = modo_deterministico(desc_ocupacao),
    n_registros_fonte = n(),
    .groups = "drop"
  )

# Replica a elegibilidade individual usada para o Estudo 2 no script de validação.
base_individual <- pessoa_ano %>%
  filter(!is.na(remuneracao_reais), remuneracao_reais > 0,
         !is.na(idade), !is.na(tempo_emprego_m),
         !is.na(valor_nivel_categoria),
         valor_nivel_categoria %in% c(1, 2, 3))

ind <- base_individual %>%
  group_by(id_aluno) %>%
  summarise(
    remuneracao_reais = mean(remuneracao_reais, na.rm = TRUE),
    idade_media = mean(idade, na.rm = TRUE),
    tempo_emprego_medio = mean(tempo_emprego_m, na.rm = TRUE),
    nivel_formacao = max(valor_nivel_categoria, na.rm = TRUE),
    n_anos_observados = n_distinct(ano_referencia),
    .groups = "drop"
  ) %>%
  filter(is.finite(remuneracao_reais), remuneracao_reais > 0,
         is.finite(idade_media), is.finite(tempo_emprego_medio),
         is.finite(nivel_formacao))

ind3_meta <- base_individual %>%
  group_by(id_aluno) %>%
  summarise(desc_ocupacao = modo_deterministico(desc_ocupacao),
            .groups = "drop")
ind3 <- ind %>% left_join(ind3_meta, by = "id_aluno") %>%
  filter(!is.na(desc_ocupacao), nzchar(desc_ocupacao))

if (nrow(ind) < 10) stop("A amostra individual do Estudo 2 é muito pequena.")
if (nrow(ind3) < 10) stop("A amostra do Estudo 3 é muito pequena.")

# Auditoria agregada, sem exportar identificadores individuais.
auditoria <- data.frame(
  indicador = c("registros_apos_filtros_estritos", "egressos_com_registros",
                "pessoa_anos", "egressos_estudo2", "egressos_estudo3",
                "anos_referencia", "cenario"),
  valor = c(nrow(base), n_distinct(base$id_aluno), nrow(pessoa_ano),
            nrow(ind), nrow(ind3),
            paste(sort(unique(base$ano_referencia)), collapse = ","),
            "estrito: ano_referencia > ano_conclusao")
)
exportar_csv(auditoria, "auditoria_amostra.csv")

# ---- 4. Funções auxiliares de clusterização ----------------------------------
padronizar <- function(x) {
  x <- as.data.frame(x)
  x[] <- lapply(x, function(v) as.numeric(scale(v)))
  x
}
calcular_ari <- function(a, b) {
  if (TEM_MCLUST) return(mclust::adjustedRandIndex(a, b))
  NA_real_
}
silhueta_media <- function(X, cl) {
  if (length(unique(cl)) < 2 || length(unique(cl)) >= nrow(X)) return(NA_real_)
  mean(cluster::silhouette(cl, dist(X))[, "sil_width"])
}
winsorizar <- function(x, probs = c(0.01, 0.99)) {
  limites <- quantile(x, probs = probs, na.rm = TRUE, names = FALSE, type = 7)
  pmin(pmax(x, limites[1]), limites[2])
}
perfil_clusters <- function(dados, cl, nome_modelo) {
  dados %>%
    mutate(cluster = factor(cl)) %>%
    group_by(cluster) %>%
    summarise(
      n = n(),
      percentual = 100 * n() / nrow(dados),
      remuneracao_media = mean(remuneracao_reais, na.rm = TRUE),
      remuneracao_mediana = median(remuneracao_reais, na.rm = TRUE),
      idade_media = mean(idade_media, na.rm = TRUE),
      tempo_emprego_medio = mean(tempo_emprego_medio, na.rm = TRUE),
      nivel_formacao_medio = mean(nivel_formacao, na.rm = TRUE),
      .groups = "drop"
    ) %>% mutate(especificacao = nome_modelo, .before = 1)
}

# ---- 5. Estudo 2 — K-means: k e sensibilidade a valores extremos --------------
vars2 <- c("remuneracao_reais", "idade_media", "tempo_emprego_medio", "nivel_formacao")
ind2_base <- ind %>% select(all_of(vars2))
ind2_wins <- ind2_base %>%
  mutate(remuneracao_reais = winsorizar(remuneracao_reais))

rodar_kmeans <- function(dados_num, nome_spec) {
  X <- padronizar(dados_num)
  resultados <- list()
  perfis <- list()
  for (k in K_AVALIADOS) {
    if (nrow(X) <= k) next
    set.seed(SEED + k)
    fit <- kmeans(X, centers = k, nstart = NSTART_KMEANS, iter.max = 100)
    resultados[[length(resultados) + 1L]] <- data.frame(
      estudo = "Estudo 2", especificacao = nome_spec, k = k,
      n = nrow(X), custo_intra = fit$tot.withinss,
      silhueta_media = silhueta_media(X, fit$cluster),
      tamanho_minimo_cluster = min(table(fit$cluster)),
      tamanho_maximo_cluster = max(table(fit$cluster))
    )
    if (k == 2) {
      perfis[[length(perfis) + 1L]] <- perfil_clusters(ind, fit$cluster, nome_spec)
    }
  }
  list(resultados = bind_rows(resultados), perfis = bind_rows(perfis))
}
res2_base <- rodar_kmeans(ind2_base, "base_padronizada")
res2_wins <- rodar_kmeans(ind2_wins, "remuneracao_winsorizada_1_99")
tab_kmeans <- bind_rows(res2_base$resultados, res2_wins$resultados)
exportar_csv(tab_kmeans, "estudo2_sensibilidade_kmeans.csv")
exportar_csv(bind_rows(res2_base$perfis, res2_wins$perfis),
             "estudo2_perfis_k2_especificacoes.csv")

# Estabilidade de inicialização para k=2 e k=3 (ARI entre partições de sementes).
estabilidade <- list()
for (k in c(2L, 3L)) {
  if (nrow(ind) <= k) next
  X <- padronizar(ind2_base)
  fits <- vector("list", 10)
  for (j in seq_along(fits)) {
    set.seed(SEED + 1000L * k + j)
    fits[[j]] <- kmeans(X, centers = k, nstart = 1, iter.max = 100)$cluster
  }
  pares <- combn(seq_along(fits), 2)
  ari <- apply(pares, 2, function(ix) calcular_ari(fits[[ix[1]]], fits[[ix[2]]]))
  estabilidade[[length(estabilidade) + 1L]] <- data.frame(
    estudo = "Estudo 2", k = k, n_execucoes = length(fits),
    ari_medio = if (all(is.na(ari))) NA_real_ else mean(ari, na.rm = TRUE),
    ari_minimo = if (all(is.na(ari))) NA_real_ else min(ari, na.rm = TRUE),
    ari_maximo = if (all(is.na(ari))) NA_real_ else max(ari, na.rm = TRUE),
    metodo = if (TEM_MCLUST) "ARI entre 10 inicializacoes aleatorias" else
      "ARI nao calculado: instale mclust"
  )
}
exportar_csv(bind_rows(estabilidade), "estudo2_estabilidade_inicializacoes.csv")

# Gráfico de custo K-means (não usar sozinho para decidir k).
p_km <- ggplot(tab_kmeans, aes(x = k, y = custo_intra, linetype = especificacao,
                               group = especificacao)) +
  geom_line() + geom_point(size = 2) +
  scale_x_continuous(breaks = K_AVALIADOS) +
  labs(title = "Estudo 2 — sensibilidade do custo K-means",
       x = "Número de clusters (k)", y = "Soma dos quadrados intracluster",
       linetype = "Especificação") + theme_minimal()
ggsave(file.path(DIR_FIGURAS, "estudo2_custo_kmeans.png"), p_km,
       width = 8, height = 5, dpi = 300)

# ---- 6. Estudo 3 — K-prototypes: k e lambda -------------------------------
vars_num3 <- c("remuneracao_reais", "idade_media", "tempo_emprego_medio", "nivel_formacao")
X3 <- ind3 %>%
  mutate(across(all_of(vars_num3), ~ as.numeric(scale(.x))),
         desc_ocupacao = factor(desc_ocupacao)) %>%
  select(all_of(vars_num3), desc_ocupacao)
if (any(!vapply(X3[vars_num3], function(x) all(is.finite(x)), logical(1)))) {
  stop("Há valores não finitos nas variáveis numéricas do Estudo 3.")
}
if (any(vapply(X3[vars_num3], function(x) sd(x) == 0, logical(1)))) {
  stop("Ao menos uma variável numérica do Estudo 3 tem variância zero.")
}

# Amostra determinística para silhueta, fixa entre especificações.
set.seed(SEED + 777L)
idx_sil <- sample(seq_len(nrow(X3)), min(N_AMOSTRA_SILHUETA, nrow(X3)))
silhueta_mista <- function(X, clusters, vars_num, lambda) {
  num <- as.matrix(X[, vars_num, drop = FALSE])
  categoria <- as.character(X$desc_ocupacao)
  D <- as.matrix(dist(num))^2
  D <- D + lambda * outer(categoria, categoria, FUN = "!=")
  if (length(unique(clusters)) < 2 || length(unique(clusters)) >= nrow(X)) return(NA_real_)
  mean(cluster::silhouette(clusters, as.dist(D))[, "sil_width"])
}

resultados3 <- list()
perfis3 <- list()
# Distribuicao da variavel categorica por cluster: exportar todas as categorias
# para k=2 em cada lambda, sem identificadores individuais.
ocupacoes_clusters3 <- list()
for (lambda_atual in LAMBDAS) {
  for (k in K_AVALIADOS_KPROTO) {
    if (nrow(X3) <= k) next
    log_msg("Estudo 3: ajustando k=", k, ", lambda=", lambda_atual)
    set.seed(SEED + 100L * k + as.integer(lambda_atual * 10))
    fit <- clustMixType::kproto(X3, k = k, lambda = lambda_atual,
                                nstart = NSTART_KPROTO)
    cl <- fit$cluster
    # O custo é comparável apenas dentro de especificações com a mesma métrica/lambda.
    sil <- tryCatch(
      silhueta_mista(X3[idx_sil, , drop = FALSE], cl[idx_sil],
                     vars_num3, lambda_atual),
      error = function(e) NA_real_
    )
    resultados3[[length(resultados3) + 1L]] <- data.frame(
      estudo = "Estudo 3", k = k, lambda = lambda_atual, n = nrow(X3),
      custo_intra = as.numeric(fit$tot.withinss),
      silhueta_mista_amostra = sil,
      tamanho_minimo_cluster = min(table(cl)),
      tamanho_maximo_cluster = max(table(cl))
    )
    if (k == 2) {
      perfis3[[length(perfis3) + 1L]] <- ind3 %>%
        mutate(cluster = factor(cl)) %>%
        group_by(cluster) %>%
        summarise(
          n = n(),
          percentual = 100 * n() / nrow(ind3),
          remuneracao_media = mean(remuneracao_reais, na.rm = TRUE),
          remuneracao_mediana = median(remuneracao_reais, na.rm = TRUE),
          idade_media = mean(idade_media, na.rm = TRUE),
          tempo_emprego_medio = mean(tempo_emprego_medio, na.rm = TRUE),
          nivel_formacao_medio = mean(nivel_formacao, na.rm = TRUE),
          .groups = "drop"
        ) %>% mutate(lambda = lambda_atual, .before = 1)

      ocupacoes_clusters3[[length(ocupacoes_clusters3) + 1L]] <- ind3 %>%
        mutate(
          lambda = lambda_atual,
          cluster = as.character(cl),
          desc_ocupacao = trimws(as.character(desc_ocupacao))
        ) %>%
        filter(!is.na(desc_ocupacao), nzchar(desc_ocupacao)) %>%
        count(lambda, cluster, desc_ocupacao, name = "n_ocupacao") %>%
        group_by(lambda, cluster) %>%
        mutate(
          n_cluster = sum(n_ocupacao),
          percentual_no_cluster = 100 * n_ocupacao / n_cluster
        ) %>%
        ungroup() %>%
        mutate(percentual_amostra = 100 * n_ocupacao / nrow(ind3)) %>%
        arrange(lambda, cluster, desc(n_ocupacao), desc_ocupacao)
    }
  }
}
tab_kproto <- bind_rows(resultados3)
exportar_csv(tab_kproto, "estudo3_sensibilidade_k_lambda.csv")
exportar_csv(bind_rows(perfis3), "estudo3_perfis_k2_por_lambda.csv")
tab_ocupacoes_clusters3 <- bind_rows(ocupacoes_clusters3)
exportar_csv(tab_ocupacoes_clusters3, "estudo3_desc_ocupacao_por_cluster_k2_lambda.csv")
# Tabela auxiliar com as dez categorias mais frequentes de cada cluster.
tab_top_ocupacoes3 <- tab_ocupacoes_clusters3 %>%
  group_by(lambda, cluster) %>%
  slice_max(order_by = n_ocupacao, n = 10, with_ties = FALSE) %>%
  ungroup() %>%
  arrange(lambda, cluster, desc(n_ocupacao))
exportar_csv(tab_top_ocupacoes3, "estudo3_top10_desc_ocupacao_por_cluster_k2_lambda.csv")

p_kp <- ggplot(tab_kproto, aes(x = k, y = silhueta_mista_amostra,
                               group = factor(lambda), linetype = factor(lambda))) +
  geom_line() + geom_point(size = 2) +
  scale_x_continuous(breaks = K_AVALIADOS) +
  labs(title = "Estudo 3 — sensibilidade da silhueta a k e lambda",
       x = "Número de clusters (k)", y = "Silhueta média (amostra fixa)",
       linetype = "Lambda") + theme_minimal()
ggsave(file.path(DIR_FIGURAS, "estudo3_silhueta_lambda.png"), p_kp,
       width = 8, height = 5, dpi = 300)

# ---- 7. Relatório de execução e notas metodológicas --------------------------
writeLines(c(
  "ANÁLISE DE ROBUSTEZ — TCC CEFET-MG / RAIS",
  paste0("Execução: ", format(Sys.time(), "%Y-%m-%d %H:%M:%S %Z")),
  paste0("Arquivo de entrada: ", CAMINHO_BASE),
  paste0("N registros estritos: ", nrow(base)),
  paste0("N egresso-ano: ", nrow(pessoa_ano)),
  paste0("N indivíduos Estudo 2: ", nrow(ind)),
  paste0("N indivíduos Estudo 3: ", nrow(ind3)),
  "",
  "Cenário principal preservado: ano_referencia > ano_conclusao.",
  "Fatores IPCA usados: 2021=1.1929197; 2022=1.0916165; 2023=1.0436746; 2024=1.0000000.",
  "Estudo 2: comparação de k=2,...,6 e análise alternativa com remuneração winsorizada nos percentis 1 e 99.",
  "Estudo 2: estabilidade entre inicializações calculada por ARI apenas se o pacote mclust estiver disponível.",
  "Estudo 3: comparação de k=2,...,6 e lambda=0.5, 1, 2; silhueta calculada em amostra fixa.",
  "Os rótulos numéricos dos clusters são arbitrários e podem trocar entre execuções; compare perfis, não apenas o número do cluster.",
  "Custos de K-means não devem ser comparados diretamente com custos de K-prototypes nem entre lambdas diferentes.",
  "A winsorização é uma especificação alternativa e não altera os resultados principais do TCC.",
  "A análise é descritiva; não identifica efeito causal da formação sobre remuneração.",
  "Saídas agregadas: tabelas CSV, gráficos PNG, auditoria e log. Nenhum identificador individual é exportado.",
  "",
  "Pacotes e versões:",
  paste(names(utils::sessionInfo()$otherPkgs),
        vapply(utils::sessionInfo()$otherPkgs, function(x) x$Version, character(1)),
        sep = " = "),
  paste0("R = ", R.version.string)
), file.path(DIR_AUDITORIA, "relatorio_metodologico.txt"), useBytes = TRUE)

writeLines(capture.output(sessionInfo()),
           file.path(DIR_AUDITORIA, "sessionInfo.txt"), useBytes = TRUE)
log_msg("Análise concluída. Saídas em: ", normalizePath(DIR_SAIDA, mustWork = FALSE))
close(con_log)
cat("\nANÁLISE DE ROBUSTEZ CONCLUÍDA\n")
cat("Base: ", CAMINHO_BASE, "\n")
cat("Saídas: ", normalizePath(DIR_SAIDA, mustWork = FALSE), "\n")
cat("Consulte auditoria/relatorio_metodologico.txt e os CSVs em tabelas/.\n")
