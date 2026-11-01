# =============================================================================
# TCC - CEFET-MG / RAIS
# SCRIPT FINAL DE CLUSTERIZAÇÃO
# =============================================================================
# Autor: Evandrino Gomes Barros
# Projeto: Trabalho Final de Curso - Especialização em Análise de Dados como
#           Método de Apoio às Políticas Públicas - UFRRJ
#
# OBJETIVO
# --------
# Reproduzir, documentar e auditar os quatro estudos de clusterização
# apresentados no TCC:
#
#   Estudo 1 - Perfis gerais de inserção profissional
#              K-means; unidade = observação anual; remuneração em log.
#
#   Estudo 2 - Formação e inserção profissional no nível individual
#              K-means; unidade = indivíduo; formação = Técnico/Graduação/
#              Pós-graduação codificada em 1/2/3.
#
#   Estudo 3 - Formação, inserção e descrição ocupacional
#              K-prototypes; unidade = indivíduo; descrições ocupacionais determinadas na nova base;
#              lambda = 1; silhueta em amostra determinística de 1.500.
#
#   Estudo 4 - Perfis institucionais dos campi
#              K-means; unidade = campus; dez campi; k = 2.
#
# O script também:
#   * registra logs em arquivo e no console;
#   * salva a sessão e versões dos pacotes;
#   * valida a estrutura da base de entrada;
#   * cria bases analíticas intermediárias;
#   * gera tabelas, gráficos e objetos RDS;
#   * executa auditorias de contagem e de consistência;
#   * registra verificações de consistência da nova análise, sem comparar com
#     resultados do TCC anterior, pois a base foi alterada;
#   * NÃO publica CPF, CNPJ, matrícula, razão social ou outros identificadores.
#
# IMPORTANTE SOBRE REPRODUTIBILIDADE E UNIDADE DE ANÁLISE
# ------------------------------------------------------
# A auditoria verifica a consistência interna da nova análise e não compara
# resultados com valores do TCC anterior, pois a base foi alterada.
# REGRA FIXA: cenário estrito em todas as execuções.
# Somente registros com ano_referencia > ano_conclusao são elegíveis.
# O ano da conclusão e os anos anteriores são excluídos da clusterização.
# A análise é descritiva e não identifica, por si só, efeito causal da formação sobre o salário.
# O Estudo 1 usa unidades egresso-ano; os Estudos 2 e 3 consolidam os anos por
# egresso; o Estudo 4 agrega indicadores por campus.
#
# Regra de remuneração: reconstrução mensal por vínculo ativo; somam-se apenas
# vínculos simultâneos e cada mês com emprego entra uma vez no denominador.
# Idade é a mediana dos registros no egresso-ano; tempo de emprego informado na
# RAIS é resumido pela média entre vínculos; formação é o maior nível observado;
# campus e ocupação são a moda determinística, com desempate alfabético.
#
# Versão do script: 3.4.0 — remuneração reconstruída por mês ativo e cenário estrito pós-conclusão
# Data de consolidação: 2026-10-11
# =============================================================================


# =============================================================================
# 0. CONFIGURAÇÃO DO USUÁRIO
# =============================================================================

# Estrutura recomendada do repositório:
#
#   base_cefet_mg_rais_2021_2024_egressos_com_1_curso_ate_2024.xlsx
#   scripts/clusterizacao_final.R
#   output/
#
# Execute este script a partir da raiz do repositório.

ARQUIVO_BASE <- "base_cefet_mg_rais_2021_2024_egressos_com_1_curso_ate_2024.xlsx"

# Cenário fixo, sem possibilidade de seleção por variável de ambiente.
# Não altere para inclusivo: a regra metodológica definida é estrita.
CENARIO_POS_CONCLUSAO <- "estrito"

DIR_SAIDA <- file.path("output_clusterizacao", "cenario_estrito")

# Para publicação no GitHub, mantenha FALSE.
# O padrão é NÃO exportar microdados individuais.
PUBLICAR_MICRODADOS <- FALSE

# Ano de referência para os preços reais.
ANO_BASE <- 2024L
ANOS_ANALISE <- 2021:2024  # anos usados para mapear os fatores IPCA já definidos abaixo
ANOS_POS_ESPERADOS <- 2022:2024  # anos possíveis após conclusão para coortes 2021-2024 na janela RAIS 2020-2024

# Regra fixa: analisar somente remunerações de anos posteriores à conclusão (ano_referencia > ano_conclusao).
ANALISAR_APENAS_POS_CONCLUSAO <- TRUE

DESCRICAO_CENARIO_POS <- "Cenário estrito: somente ano_referencia > ano_conclusao; exclui o ano da conclusão e os anos anteriores."

# Seeds utilizadas para tornar a execução reproduzível.
SEED_KMEANS <- 1234L
SEED_KPROTO <- 1234L
SEED_SILHUETA <- 2024L

# Número de inicializações.
NSTART_KMEANS <- 50L
NSTART_KPROTO_EXPLORACAO <- 30L
NSTART_KPROTO_FINAL <- 50L

# Parâmetros do K-prototypes.
LAMBDA_PRINCIPAL <- 1
LAMBDA_SENSIBILIDADE <- c(0.5, 1, 2)

# Silhueta do Estudo 3.
N_AMOSTRA_SILHUETA_ESTUDO3 <- 1500L

# Faixas de k avaliadas nos estudos.
K_ESTUDO1 <- 2:6
K_ESTUDO2 <- 2:6
K_ESTUDO3 <- 2:6
K_ESTUDO4 <- 2:5


# =============================================================================
# 1. PACOTES
# =============================================================================

PACOTES <- c(
  "readxl",
  "dplyr",
  "tidyr",
  "ggplot2",
  "cluster",
  "clustMixType"
)

faltantes <- PACOTES[
  !vapply(
    PACOTES,
    requireNamespace,
    quietly = TRUE,
    FUN.VALUE = logical(1)
  )
]

if (length(faltantes) > 0) {
  stop(
    paste0(
      "Pacotes ausentes: ",
      paste(faltantes, collapse = ", "),
      "\nInstale-os antes de executar o script com:\n",
      "install.packages(c(",
      paste(sprintf('"%s"', faltantes), collapse = ", "),
      "))"
    )
  )
}

suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(tidyr)
  library(ggplot2)
  library(cluster)
  library(clustMixType)
})


# =============================================================================
# 2. DIRETÓRIOS, LOG E CONTROLE DA EXECUÇÃO
# =============================================================================

dir.create(
  DIR_SAIDA,
  recursive = TRUE,
  showWarnings = FALSE
)

DIR_DADOS <- file.path(
  DIR_SAIDA,
  "bases_auxiliares"
)

DIR_TABELAS <- file.path(
  DIR_SAIDA,
  "tabelas"
)

DIR_FIGURAS <- file.path(
  DIR_SAIDA,
  "figuras"
)

DIR_MODELOS <- file.path(
  DIR_SAIDA,
  "modelos"
)

DIR_AUDITORIA <- file.path(
  DIR_SAIDA,
  "auditoria"
)

invisible(
  lapply(
    c(
      DIR_DADOS,
      DIR_TABELAS,
      DIR_FIGURAS,
      DIR_MODELOS,
      DIR_AUDITORIA
    ),
    dir.create,
    recursive = TRUE,
    showWarnings = FALSE
  )
)

ARQUIVO_LOG <- file.path(
  DIR_AUDITORIA,
  paste0(
    "clusterizacao_",
    format(Sys.time(), "%Y%m%d_%H%M%S"),
    ".log"
  )
)

log_msg <- function(..., nivel = "INFO") {

  mensagem <- paste0(
    "[",
    format(Sys.time(), "%Y-%m-%d %H:%M:%S"),
    "] [",
    nivel,
    "] ",
    paste(..., collapse = "")
  )

  cat(
    mensagem,
    "\n"
  )

  con_log <- file(
    ARQUIVO_LOG,
    open = "at",
    encoding = "UTF-8"
  )
  on.exit(close(con_log), add = TRUE)

  cat(
    mensagem,
    "\n",
    file = con_log
  )
}
log_msg(
  "Início da execução do script final de clusterização."
)

log_msg(
  "Arquivo de entrada: ",
  normalizePath(
    ARQUIVO_BASE,
    mustWork = FALSE
  )
)

log_msg(
  "Diretório de saída: ",
  normalizePath(
    DIR_SAIDA,
    mustWork = FALSE
  )
)

log_msg(
  "PUBLICAR_MICRODADOS = ",
  PUBLICAR_MICRODADOS
)


# =============================================================================
# 3. FUNÇÕES DE APOIO
# =============================================================================

# -----------------------------------------------------------------------------
# 3.1. Verificação de colunas
# -----------------------------------------------------------------------------

exigir_colunas <- function(
    dados,
    colunas,
    nome_objeto = "base"
) {

  ausentes <- setdiff(
    colunas,
    names(dados)
  )

  if (length(ausentes) > 0) {

    stop(
      paste0(
        "A base ",
        nome_objeto,
        " não contém as colunas obrigatórias: ",
        paste(ausentes, collapse = ", ")
      )
    )
  }

  invisible(TRUE)
}


# -----------------------------------------------------------------------------
# 3.2. Exportação de CSV
# -----------------------------------------------------------------------------

exportar_csv <- function(
    x,
    nome,
    digits = NULL
) {

  caminho <- file.path(
    DIR_TABELAS,
    paste0(
      nome,
      ".csv"
    )
  )

  y <- x

  if (!is.null(digits)) {

    y <- y %>%
      mutate(
        across(
          where(is.numeric),
          ~ round(
            .x,
            digits
          )
        )
      )
  }

  write.csv(
    y,
    caminho,
    row.names = FALSE,
    fileEncoding = "UTF-8"
  )

  log_msg(
    "Tabela exportada: ",
    caminho
  )

  invisible(caminho)
}


# -----------------------------------------------------------------------------
# 3.3. Exportação de objetos RDS
# -----------------------------------------------------------------------------

exportar_rds <- function(
    x,
    nome
) {

  caminho <- file.path(
    DIR_MODELOS,
    paste0(
      nome,
      ".rds"
    )
  )

  saveRDS(
    x,
    caminho
  )

  log_msg(
    "Objeto RDS exportado: ",
    caminho
  )

  invisible(caminho)
}


# -----------------------------------------------------------------------------
# 3.4. Registro de auditoria
# -----------------------------------------------------------------------------

auditoria <- list()

registrar_auditoria <- function(
    item,
    valor_calculado,
    valor_esperado = NA,
    tolerancia = 1e-6,
    observacao = ""
) {

  ok <- if (
    is.na(valor_esperado)
  ) {

    NA

} else if (
  is.character(valor_calculado) ||
  is.character(valor_esperado)
) {
  
  isTRUE(
    identical(
      as.character(valor_calculado),
      as.character(valor_esperado)
    )
  )
  
} else {
  
  isTRUE(
    abs(
      valor_calculado -
        valor_esperado
    ) <= tolerancia
  )
}

  auditoria[[length(auditoria) + 1L]] <<-
    data.frame(
      item = item,
      valor_calculado =
        as.character(
          valor_calculado
        ),
      valor_esperado =
        as.character(
          valor_esperado
        ),
      tolerancia =
        tolerancia,
      status =
        ifelse(
          is.na(ok),
          "NAO_APLICAVEL",
          ifelse(
            ok,
            "OK",
            "DIVERGENCIA"
          )
        ),
      observacao =
        observacao,
      stringsAsFactors = FALSE
    )

  if (
    !is.na(ok) &&
      !ok
  ) {

    log_msg(
      "DIVERGÊNCIA: ",
      item,
      " | calculado=",
      valor_calculado,
      " | esperado=",
      valor_esperado,
      " | ",
      observacao,
      nivel = "WARN"
    )

  } else {

    log_msg(
      "Auditoria: ",
      item,
      " | calculado=",
      valor_calculado,
      " | esperado=",
      valor_esperado,
      " | status=",
      ifelse(
        is.na(ok),
        "NAO_APLICAVEL",
        ifelse(
          ok,
          "OK",
          "DIVERGENCIA"
        )
      )
    )
  }

  invisible(ok)
}


# -----------------------------------------------------------------------------
# 3.5. Padronização Z
# -----------------------------------------------------------------------------

padronizar <- function(
    dados,
    variaveis
) {

  dados %>%
    mutate(
      across(
        all_of(variaveis),
        ~ as.numeric(
          scale(.x)
        )
      )
    )
}


# -----------------------------------------------------------------------------
# 3.6. Silhueta média para K-means
# -----------------------------------------------------------------------------

silhueta_kmeans <- function(
    X,
    clusters
) {

  d <- dist(
    X
  )

  sil <- silhouette(
    clusters,
    d
  )

  mean(
    sil[, "sil_width"]
  )
}


# -----------------------------------------------------------------------------
# 3.7. Distância mista usada no Estudo 3
# -----------------------------------------------------------------------------
#
# A especificação reproduz a lógica utilizada no script revisado do Estudo 3:
#
# D = distância euclidiana quadrática das variáveis numéricas
#     + lambda * I(descrição ocupacional diferente)
#
# A silhueta é calculada sobre a matriz de distâncias da amostra.

calcular_silhueta_mista <- function(
    X,
    clusters,
    vars_num,
    lambda = 1
) {

  matriz_num <- as.matrix(
    X[
      ,
      vars_num,
      drop = FALSE
    ]
  )

  categoria <- as.character(
    X$desc_ocupacao
  )

  D <- as.matrix(
    dist(
      matriz_num
    )
  )^2

  D <- D +
    lambda *
    outer(
      categoria,
      categoria,
      FUN = "!="
    )

  sil <- silhouette(
    clusters,
    as.dist(
      D
    )
  )

  mean(
    sil[, "sil_width"]
  )
}


# -----------------------------------------------------------------------------
# 3.8. Relatório de estrutura da base
# -----------------------------------------------------------------------------

relatorio_estrutura <- function(
    dados,
    nome
) {

  estrutura <- data.frame(
    variavel = names(
      dados
    ),
    classe = vapply(
      dados,
      function(x)
        paste(
          class(x),
          collapse = "/"
        ),
      character(1)
    ),
    n_na = vapply(
      dados,
      function(x)
        sum(
          is.na(x)
        ),
      numeric(1)
    ),
    n_unicos = vapply(
      dados,
      function(x)
        dplyr::n_distinct(x),
      numeric(1)
    ),
    stringsAsFactors = FALSE
  )

  write.csv(
    estrutura,
    file.path(
      DIR_AUDITORIA,
      paste0(
        nome,
        "_estrutura.csv"
      )
    ),
    row.names = FALSE,
    fileEncoding = "UTF-8"
  )

  estrutura
}


# -----------------------------------------------------------------------------
# 3.9. Gráfico de seleção de k
# -----------------------------------------------------------------------------

plotar_selecao_k <- function(
    tabela,
    nome,
    titulo,
    yvar,
    ylab
) {

  p <- ggplot(
    tabela,
    aes(
      x = k,
      y = .data[[yvar]]
    )
  ) +
    geom_line() +
    geom_point(
      size = 2
    ) +
    scale_x_continuous(
      breaks = tabela$k
    ) +
    labs(
      title = titulo,
      x = "Número de clusters (k)",
      y = ylab
    ) +
    theme_minimal()

  caminho <- file.path(
    DIR_FIGURAS,
    paste0(
      nome,
      ".png"
    )
  )

  ggsave(
    caminho,
    p,
    width = 7,
    height = 4.5,
    dpi = 300
  )

  log_msg(
    "Figura exportada: ",
    caminho
  )

  invisible(p)
}


# =============================================================================
# 4. LEITURA DA BASE
# =============================================================================

if (!file.exists(
  ARQUIVO_BASE
)) {

  stop(
    "Arquivo de entrada não encontrado: ",
    ARQUIVO_BASE
  )
}

log_msg(
  "Lendo a planilha de entrada."
)

abas <- excel_sheets(
  ARQUIVO_BASE
)

log_msg(
  "Abas disponíveis: ",
  paste(
    abas,
    collapse = ", "
  )
)

# A base fonte nova possui uma aba principal com ano_conclusao e egressos com curso único.
DADOS <- read_excel(
  ARQUIVO_BASE,
  sheet = abas[1],
  .name_repair = "unique"
)

log_msg(
  "Base carregada: ",
  nrow(DADOS),
  " linhas e ",
  ncol(DADOS),
  " colunas."
)

relatorio_estrutura(
  DADOS,
  "01_base_original"
)


# =============================================================================
# 5. VALIDAÇÃO DA ESTRUTURA
# =============================================================================

COLUNAS_OBRIGATORIAS <- c(
  "id_aluno",
  "id_curso",
  "ano_conclusao",
  "remun_media_nom",
  "dt_admissao",
  "dt_desligamento",
  "vinculo_ativo_3112",
  "tempo_emprego_m",
  "ano_referencia",
  "desc_ocupacao",
  "idade",
  "campus",
  "nivel_categoria",
  "valor_nivel_categoria"
)

exigir_colunas(
  DADOS,
  COLUNAS_OBRIGATORIAS,
  "base original"
)

log_msg(
  "Todas as colunas obrigatórias estão presentes."
)


# =============================================================================
# 6. LIMPEZA E NORMALIZAÇÃO DA BASE
# =============================================================================
#
# Na nova base, remun_media_nom já está em reais, dividida por 100 na etapa
# anterior e arredondada ao inteiro mais próximo. NÃO dividir novamente por 100.
# A remuneração nominal é atualizada para preços médios de 2024.
#
# Fatores utilizados no TCC:
#   2020 = 1,2919518
#   2021 = 1,1929197
#   2022 = 1,0916165
#   2023 = 1,0436746
#   2024 = 1,0000000

FATORES_IPCA <- data.frame(
  ano = ANOS_ANALISE,
  fator_2024 = c(
    1.1929197,
    1.0916165,
    1.0436746,
    1.0000000
  )
)

write.csv(
  FATORES_IPCA,
  file.path(
    DIR_AUDITORIA,
    "fatores_atualizacao_ipca.csv"
  ),
  row.names = FALSE,
  fileEncoding = "UTF-8"
)

base_limpa <- DADOS %>%
  mutate(
    ano_referencia =
      as.integer(
        ano_referencia
      ),
    remun_media_nom =
      as.numeric(
        remun_media_nom
      ),
    tempo_emprego_m =
      as.numeric(
        tempo_emprego_m
      ),
    idade =
      as.numeric(
        idade
      ),
    valor_nivel_categoria =
      as.numeric(
        valor_nivel_categoria
      ),
    ano_conclusao = as.integer(ano_conclusao),
    remun_nominal_reais = round(remun_media_nom)
  ) %>%
  left_join(
    FATORES_IPCA,
    by = c(
      "ano_referencia" = "ano"
    )
  ) %>%
  mutate(
    remuneracao_reais =
      remun_nominal_reais *
      fator_2024
  ) %>%
  mutate(
    desc_ocupacao =
      trimws(
        as.character(
          desc_ocupacao
        )
      ),
    campus =
      trimws(
        as.character(
          campus
        )
      ),
    nivel_categoria =
      trimws(
        as.character(
          nivel_categoria
        )
      ),
    campus = case_when(
      campus == "Campus I" ~ "Campus I - Belo Horizonte",
      campus == "Campus II" ~ "Campus II - Belo Horizonte",
      TRUE ~ campus
    )
  ) %>%
  filter(
    !is.na(id_aluno),
    !is.na(ano_conclusao),
    ano_conclusao >= 2021,
    ano_conclusao <= 2024,
    !is.na(remun_media_nom),
    remun_media_nom > 0
  ) %>%
  group_by(id_aluno) %>%
  filter(n_distinct(id_curso, na.rm = TRUE) <= 1) %>%
  ungroup()

if (any(is.na(base_limpa$campus) | trimws(base_limpa$campus) == "")) {
  stop("Há campus ausente/vazio. Atribuição manual exige confirmação com a base acadêmica; o script não inventará campus.")
}

if (ANALISAR_APENAS_POS_CONCLUSAO) {
  base_limpa <- base_limpa %>%
    filter(ano_referencia > ano_conclusao)
}

log_msg("Cenário pós-conclusão: ", CENARIO_POS_CONCLUSAO, " | ", DESCRICAO_CENARIO_POS)
log_msg("Observação metodológica: cenário estrito; como a RAIS é anual, o ano da conclusão é excluído para evitar misturar remunerações anteriores e posteriores à conclusão.")

# Função determinística para selecionar a categoria mais frequente.
# Em caso de empate, escolhe alfabeticamente para garantir reprodutibilidade.
moda_deterministica <- function(x) {
  x <- trimws(as.character(x))
  x <- x[!is.na(x) & nzchar(x)]
  if (length(x) == 0L) return(NA_character_)
  frequencias <- table(x)
  candidatas <- names(frequencias)[frequencias == max(frequencias)]
  sort(candidatas)[1]
}

# Reconstrução aproximada da remuneração mensal por pessoa-ano a partir dos
# vínculos ativos em cada mês. remun_media_nom é a média mensal do vínculo nos
# meses trabalhados (13º excluído na definição estatística da RAIS); multiplicá-la
# pelos meses ativos estima a massa salarial do vínculo no ano.
# Como dt_admissao/dt_desligamento podem vir como datas Excel, texto ou "NULL",
# a função normaliza os formatos sem presumir que "NULL" seja uma data válida.
normalizar_data_rais <- function(x) {
  if (inherits(x, "Date")) return(x)
  if (inherits(x, "POSIXt")) return(as.Date(x))
  if (is.numeric(x)) return(as.Date(ifelse(is.na(x) | x <= 0, NA_real_, x), origin = "1899-12-30"))
  z <- trimws(as.character(x))
  z[z %in% c("", "NULL", "null", "NA", "N/A", "0")] <- NA_character_
  out <- as.Date(rep(NA_character_, length(z)), format = "%Y-%m-%d")
  for (fmt in c("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d")) {
    ix <- is.na(out) & !is.na(z)
    if (any(ix)) {
      parsed <- suppressWarnings(as.Date(z[ix], format = fmt))
      out[ix] <- parsed
    }
  }
  out
}

base_vinculos <- base_limpa %>%
  mutate(
    .admissao = normalizar_data_rais(dt_admissao),
    .desligamento_informado = normalizar_data_rais(dt_desligamento),
    .ativo_3112 = tolower(trimws(as.character(vinculo_ativo_3112))) %in% c("1", "sim", "s", "true", "ativo"),
    .inicio_ano = as.Date(sprintf("%d-01-01", ano_referencia)),
    .fim_ano = as.Date(sprintf("%d-12-31", ano_referencia)),
    .inicio = pmax(.admissao, .inicio_ano, na.rm = TRUE),
    .fim = if_else(is.na(.desligamento_informado) & .ativo_3112, .fim_ano, pmin(.desligamento_informado, .fim_ano, na.rm = TRUE)),
    .inicio_mes = as.integer(format(.inicio, "%Y")) * 12L + as.integer(format(.inicio, "%m")),
    .fim_mes = as.integer(format(.fim, "%Y")) * 12L + as.integer(format(.fim, "%m")),
    .meses_vinculo = if_else(!is.na(.inicio) & !is.na(.fim) & .fim >= .inicio,
                             pmax(0L, .fim_mes - .inicio_mes + 1L), 0L)
  )

# Registros sem admissão válida ou sem desligamento interpretável e sem marcação
# de ativo em 31/12 são auditados e interrompem o cálculo: não se inventa duração.
problemas_datas <- base_vinculos %>%
  filter(is.na(.admissao) | (is.na(.desligamento_informado) & !.ativo_3112))
write.csv(
  problemas_datas %>% select(id_aluno, ano_referencia, dt_admissao, dt_desligamento, vinculo_ativo_3112),
  file.path(DIR_AUDITORIA, "vinculos_com_datas_incompletas.csv"),
  row.names = FALSE, fileEncoding = "UTF-8"
)
if (nrow(problemas_datas) > 0L) {
  stop("Há vínculos sem data de admissão ou sem desligamento interpretável e sem ativo_3112=1. Consulte auditoria/vinculos_com_datas_incompletas.csv; não é seguro estimar remuneração sem revisar esses casos.")
}

# Expande cada vínculo para os meses-calendário em que esteve ativo. Se dois
# vínculos coexistem, ambos contribuem para a remuneração daquele mês; o mês,
# porém, entra uma única vez no denominador da média da pessoa.
linhas_mes <- lapply(seq_len(nrow(base_vinculos)), function(i) {
  r <- base_vinculos[i, ]
  if (is.na(r$.meses_vinculo) || r$.meses_vinculo < 1L) return(NULL)
  m <- seq.int(r$.inicio_mes, r$.fim_mes)
  data.frame(
    id_aluno = r$id_aluno,
    ano_referencia = as.integer(r$ano_referencia),
    mes_indice = m,
    remuneracao_mensal_vinculo = as.numeric(r$remuneracao_reais),
    remuneracao_mensal_nominal_vinculo = as.numeric(r$remun_nominal_reais),
    stringsAsFactors = FALSE
  )
})
remuneracao_mes <- bind_rows(linhas_mes) %>%
  group_by(id_aluno, ano_referencia, mes_indice) %>%
  summarise(
    remuneracao_mes_real = sum(remuneracao_mensal_vinculo, na.rm = TRUE),
    remuneracao_mes_nominal = sum(remuneracao_mensal_nominal_vinculo, na.rm = TRUE),
    n_vinculos_ativos_mes = n(),
    .groups = "drop"
  )
remuneracao_pessoa_ano <- remuneracao_mes %>%
  group_by(id_aluno, ano_referencia) %>%
  summarise(
    meses_com_emprego = n_distinct(mes_indice),
    massa_salarial_estimada_real = sum(remuneracao_mes_real, na.rm = TRUE),
    massa_salarial_estimada_nominal = sum(remuneracao_mes_nominal, na.rm = TRUE),
    remuneracao_reais = mean(remuneracao_mes_real, na.rm = TRUE),
    remuneracao_nominal_reais = mean(remuneracao_mes_nominal, na.rm = TRUE),
    remuneracao_reais_12m = massa_salarial_estimada_real / 12,
    remuneracao_nominal_12m = massa_salarial_estimada_nominal / 12,
    meses_com_sobreposicao = sum(n_vinculos_ativos_mes > 1L),
    max_vinculos_simultaneos = max(n_vinculos_ativos_mes),
    .groups = "drop"
  )

# Demais atributos são consolidados como antes; remuneração é substituída pela
# medida temporal acima. Idade usa mediana, formação o maior nível, campus e
# ocupação a moda determinística. Tempo de emprego permanece descritivo.
base_pessoa_ano <- base_vinculos %>%
  group_by(id_aluno, ano_referencia) %>%
  summarise(
    ano_conclusao = min(ano_conclusao, na.rm = TRUE),
    fator_2024 = first(fator_2024),
    idade = if (all(is.na(idade))) NA_real_ else median(idade, na.rm = TRUE),
    tempo_emprego_m = if (all(is.na(tempo_emprego_m))) NA_real_ else mean(tempo_emprego_m, na.rm = TRUE),
    n_niveis_formacao_distintos = n_distinct(valor_nivel_categoria[!is.na(valor_nivel_categoria)]),
    n_campi_distintos = n_distinct(campus[!is.na(campus) & trimws(campus) != ""]),
    n_ocupacoes_distintas = n_distinct(desc_ocupacao[!is.na(desc_ocupacao) & trimws(desc_ocupacao) != ""]),
    valor_nivel_categoria = if (all(is.na(valor_nivel_categoria))) NA_real_ else max(valor_nivel_categoria, na.rm = TRUE),
    campus = moda_deterministica(campus),
    nivel_categoria = moda_deterministica(nivel_categoria),
    desc_ocupacao = moda_deterministica(desc_ocupacao),
    n_registros_fonte = n(),
    n_vinculos_pessoa_ano = n(),
    n_vinculos_unicos = n() == 1L,
    .groups = "drop"
  ) %>%
  left_join(remuneracao_pessoa_ano, by = c("id_aluno", "ano_referencia")) %>%
  mutate(
    remuneracao_reais = if_else(meses_com_emprego > 0L, remuneracao_reais, NA_real_),
    remuneracao_nominal_reais = if_else(meses_com_emprego > 0L, remuneracao_nominal_reais, NA_real_),
    remuneracao_reais_12m = if_else(meses_com_emprego > 0L, remuneracao_reais_12m, NA_real_),
    remuneracao_nominal_12m = if_else(meses_com_emprego > 0L, remuneracao_nominal_12m, NA_real_)
  )

write.csv(
  base_pessoa_ano %>% select(ano_referencia, meses_com_emprego, n_vinculos_pessoa_ano,
                             n_vinculos_unicos, meses_com_sobreposicao,
                             max_vinculos_simultaneos, remuneracao_reais,
                             remuneracao_reais_12m),
  file.path(DIR_AUDITORIA, "auditoria_remuneracao_pessoa_ano.csv"),
  row.names = FALSE, fileEncoding = "UTF-8"
)
# Arquivo de sensibilidade: pessoa-anos que contêm exatamente um vínculo-fonte.
# Serve para repetir a preparação e os modelos em uma amostra sem múltiplos
# vínculos no mesmo ano; as estimativas deste arquivo não substituem a principal.
write.csv(
  base_pessoa_ano %>% filter(n_vinculos_unicos) %>%
    select(-id_aluno),
  file.path(DIR_AUDITORIA, "sensibilidade_pessoa_ano_vinculo_unico.csv"),
  row.names = FALSE, fileEncoding = "UTF-8"
)
write.csv(
  base_pessoa_ano %>% select(ano_referencia, meses_com_emprego,
                             remuneracao_reais, remuneracao_reais_12m,
                             massa_salarial_estimada_real),
  file.path(DIR_AUDITORIA, "sensibilidade_remuneracao_media_vs_12m.csv"),
  row.names = FALSE, fileEncoding = "UTF-8"
)


# Auditoria detalhada da contagem por etapa.
# Esta trilha replica, em sequência, os filtros efetivamente aplicados acima,
# permitindo distinguir exclusões por filtro da redução causada pela consolidação.
adicionar_etapa_contagem <- function(etapa, entrada, saida, natureza, descricao) {
  data.frame(
    etapa = etapa,
    registros_entrada = as.integer(entrada),
    registros_saida = as.integer(saida),
    reducao_na_etapa = as.integer(entrada - saida),
    percentual_da_base_original = round(100 * saida / nrow(DADOS), 2),
    natureza = natureza,
    descricao = descricao,
    stringsAsFactors = FALSE
  )
}

contagem_filtragem <- data.frame(
  etapa = "01_base_original",
  registros_entrada = NA_integer_,
  registros_saida = as.integer(nrow(DADOS)),
  reducao_na_etapa = 0L,
  percentual_da_base_original = 100,
  natureza = "Base inicial",
  descricao = "Linhas da primeira aba da planilha fonte.",
  stringsAsFactors = FALSE
)

base_contagem <- DADOS %>%
  mutate(
    ano_referencia = as.integer(ano_referencia),
    ano_conclusao = as.integer(ano_conclusao),
    remun_media_nom = as.numeric(remun_media_nom)
  )

n_entrada <- nrow(base_contagem)
base_contagem <- base_contagem %>% filter(!is.na(id_aluno))
contagem_filtragem <- bind_rows(contagem_filtragem, adicionar_etapa_contagem(
  "02_id_aluno_preenchido", n_entrada, nrow(base_contagem), "Filtro",
  "Exige identificador do egresso preenchido."
))

n_entrada <- nrow(base_contagem)
base_contagem <- base_contagem %>% filter(!is.na(ano_conclusao))
contagem_filtragem <- bind_rows(contagem_filtragem, adicionar_etapa_contagem(
  "03_ano_conclusao_preenchido", n_entrada, nrow(base_contagem), "Filtro",
  "Exige ano de conclusão preenchido."
))

n_entrada <- nrow(base_contagem)
base_contagem <- base_contagem %>% filter(ano_conclusao >= 2021, ano_conclusao <= 2024)
contagem_filtragem <- bind_rows(contagem_filtragem, adicionar_etapa_contagem(
  "04_coorte_2021_2024", n_entrada, nrow(base_contagem), "Filtro",
  "Mantém egressos com ano de conclusão entre 2021 e 2024."
))

n_entrada <- nrow(base_contagem)
base_contagem <- base_contagem %>% filter(!is.na(remun_media_nom))
contagem_filtragem <- bind_rows(contagem_filtragem, adicionar_etapa_contagem(
  "05_remuneracao_preenchida", n_entrada, nrow(base_contagem), "Filtro",
  "Exige remun_media_nom não ausente."
))

n_entrada <- nrow(base_contagem)
base_contagem <- base_contagem %>% filter(remun_media_nom > 0)
contagem_filtragem <- bind_rows(contagem_filtragem, adicionar_etapa_contagem(
  "06_remuneracao_positiva", n_entrada, nrow(base_contagem), "Filtro",
  "Exige remun_media_nom maior que zero."
))

n_entrada <- nrow(base_contagem)
base_contagem <- base_contagem %>%
  group_by(id_aluno) %>%
  filter(n_distinct(id_curso, na.rm = TRUE) <= 1) %>%
  ungroup()
contagem_filtragem <- bind_rows(contagem_filtragem, adicionar_etapa_contagem(
  "07_um_curso_por_egresso", n_entrada, nrow(base_contagem), "Filtro",
  "Mantém egressos com no máximo um id_curso distinto."
))

n_entrada <- nrow(base_contagem)
if (ANALISAR_APENAS_POS_CONCLUSAO) {
  base_contagem <- base_contagem %>% filter(ano_referencia > ano_conclusao)
  descricao_pos <- "Cenário estrito: mantém somente ano_referencia > ano_conclusao; exclui o próprio ano de conclusão e os anos anteriores."
} else {
  descricao_pos <- "Filtro pós-conclusão desativado; nenhum registro removido nesta etapa."
}
contagem_filtragem <- bind_rows(contagem_filtragem, adicionar_etapa_contagem(
  "08_pos_conclusao", n_entrada, nrow(base_contagem), "Filtro",
  descricao_pos
))

# A consolidação reduz linhas por agrupamento, mas não é exclusão de registros:
# os valores de vínculos são resumidos em uma linha por egresso-ano.
contagem_filtragem <- bind_rows(contagem_filtragem, adicionar_etapa_contagem(
  "09_consolidacao_pessoa_ano", nrow(base_contagem), nrow(base_pessoa_ano),
  "Consolidação (não é filtro)",
  "Uma linha por id_aluno e ano_referencia; remuneração média calculada mês a mês, somando vínculos simultâneos e contando cada mês empregado uma vez."
))

write.csv(
  contagem_filtragem,
  file.path(DIR_AUDITORIA, "contagem_filtragem_base.csv"),
  row.names = FALSE,
  fileEncoding = "UTF-8"
)

for (i_contagem in seq_len(nrow(contagem_filtragem))) {
  log_msg(
    "Contagem por etapa: ", contagem_filtragem$etapa[i_contagem],
    " | entrada=", contagem_filtragem$registros_entrada[i_contagem],
    " | saída=", contagem_filtragem$registros_saida[i_contagem],
    " | redução=", contagem_filtragem$reducao_na_etapa[i_contagem]
  )
}

log_msg(
  "Consolidação pessoa-ano: ", nrow(base_limpa),
  " registros RAIS reduzidos a ", nrow(base_pessoa_ano),
  " unidades egresso-ano."
)

registrar_auditoria(
  "unidades_pessoa_ano",
  nrow(base_pessoa_ano),
  observacao = "Uma linha por id_aluno e ano_referencia após consolidar vínculos e reconstruir remuneração mensal por meses ativos."
)

registrar_auditoria(
  "duplicatas_chave_pessoa_ano",
  anyDuplicated(base_pessoa_ano[c("id_aluno", "ano_referencia")]),
  0,
  observacao = "A chave egresso-ano deve ser única após a consolidação."
)

registrar_auditoria(
  "pessoa_anos_com_multiplos_campi",
  sum(base_pessoa_ano$n_campi_distintos > 1, na.rm = TRUE),
  0,
  observacao = "Campus deve ser consistente para o egresso-ano; em conflito, a moda determinística é mantida e o caso deve ser revisado."
)

registrar_auditoria(
  "pessoa_anos_com_multiplas_ocupacoes",
  sum(base_pessoa_ano$n_ocupacoes_distintas > 1, na.rm = TRUE),
  observacao = "Vínculos podem ter ocupações diferentes; a ocupação do egresso-ano é a mais frequente, com desempate alfabético."
)

anos_encontrados <- sort(
  unique(
    base_limpa$ano_referencia[
      !is.na(
        base_limpa$ano_referencia
      )
    ]
  )
)

log_msg(
  "Anos encontrados na base: ",
  paste(
    anos_encontrados,
    collapse = ", "
  )
)

registrar_auditoria(
  "anos_analise_pos_conclusao",
  paste(anos_encontrados, collapse = ","),
  paste(ANOS_POS_ESPERADOS, collapse = ","),
  observacao = "Base analítica estrita: somente ano_referencia > ano_conclusao; para coortes 2021-2024 e RAIS até 2024, os anos possíveis são 2022-2024."
)

n_sem_fator <- sum(
  is.na(
    base_limpa$fator_2024
  ) &
    !is.na(
      base_limpa$ano_referencia
    )
)

registrar_auditoria(
  "registros_sem_fator_ipca",
  n_sem_fator,
  0,
  observacao =
    "Todos os anos utilizados devem possuir fator de atualização."
)

# Base limpa sanitizada:
# não contém CPF, CNPJ, matrícula, razão social ou id_aluno.
COLUNAS_PUBLICAVEIS_BASE <- c(
  "remun_nominal_reais",
  "fator_2024",
  "remuneracao_reais",
  "tempo_emprego_m",
  "idade",
  "ano_referencia",
  "desc_ocupacao",
  "campus",
  "nivel_categoria",
  "valor_nivel_categoria"
)

base_limpa_publicavel <-
  base_limpa %>%
  select(
    any_of(
      COLUNAS_PUBLICAVEIS_BASE
    )
  )

saveRDS(
  base_limpa_publicavel,
  file.path(
    DIR_DADOS,
    "base_limpa_publicavel.rds"
  )
)

log_msg(
  "Base limpa sanitizada salva."
)

if (PUBLICAR_MICRODADOS) {

  base_trabalho <- base_limpa

  saveRDS(
    base_trabalho,
    file.path(
      DIR_DADOS,
      "base_trabalho_LOCAL_NAO_PUBLICAR.rds"
    )
  )

  log_msg(
    "Base individual local salva. NÃO publicar no GitHub.",
    nivel = "WARN"
  )
}


# =============================================================================
# 7. BASE INDIVIDUAL CONSOLIDADA A PARTIR DE PESSOA-ANO
# =============================================================================
# Cada egresso-ano tem o mesmo peso no cálculo da média individual ao longo
# do período. n_registros conserva a quantidade de registros RAIS de origem.
# Formação individual = maior nível observado no período.

base_para_individual <- base_pessoa_ano %>%
  filter(
    !is.na(remuneracao_reais),
    remuneracao_reais > 0,
    !is.na(idade),
    !is.na(tempo_emprego_m),
    !is.na(valor_nivel_categoria),
    valor_nivel_categoria %in% c(1, 2, 3)
  )

ind_num <- base_para_individual %>%
  group_by(id_aluno) %>%
  summarise(
    remuneracao_reais = mean(remuneracao_reais, na.rm = TRUE),
    idade_media = mean(idade, na.rm = TRUE),
    tempo_emprego_medio = mean(tempo_emprego_m, na.rm = TRUE),
    nivel_formacao = max(valor_nivel_categoria, na.rm = TRUE),
    n_anos_observados = n_distinct(ano_referencia),
    n_pessoa_anos = n(),
    n_registros = sum(n_registros_fonte, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  mutate(nivel_formacao = as.numeric(nivel_formacao))

log_msg(
  "Indivíduos na base individual consolidada: ",
  nrow(ind_num)
)

registrar_auditoria(
  "individuos_estudos_2_e_3",
  nrow(ind_num),
  observacao = "Indivíduos calculados a partir de unidades pessoa-ano consolidadas."
)

saveRDS(
  ind_num,
  file.path(DIR_DADOS, "base_individual_consolidada.rds")
)

# =============================================================================
# 8. ESTUDO 1 - PERFIS GERAIS DE INSERÇÃO PROFISSIONAL
# =============================================================================

log_msg(
  "============================================================"
)

log_msg(
  "INÍCIO DO ESTUDO 1"
)

log_msg(
  "============================================================"
)

# Unidade: observação anual.
#
# Variáveis:
#   * log da remuneração média real;
#   * idade;
#   * tempo de emprego.
#
# A remuneração é transformada em logaritmo antes da padronização.

estudo1 <- base_pessoa_ano %>%
  filter(
    !is.na(
      id_aluno
    ),
    !is.na(
      remuneracao_reais
    ),
    remuneracao_reais > 0,
    !is.na(
      idade
    ),
    !is.na(
      tempo_emprego_m
    )
  ) %>%
  mutate(
    log_remuneracao =
      log(
        remuneracao_reais
      )
  ) %>%
  select(
    id_aluno,
    ano_referencia,
    remuneracao_reais,
    log_remuneracao,
    idade,
    tempo_emprego_m
  )

log_msg(
  "Observações do Estudo 1: ",
  nrow(estudo1)
)

registrar_auditoria(
  "estudo1_n_observacoes",
  nrow(estudo1),
  observacao = "Nova base pós-conclusão; unidade egresso-ano com remuneração reconstruída por mês ativo."
)

vars_estudo1 <- c(
  "log_remuneracao",
  "idade",
  "tempo_emprego_m"
)

X1 <- estudo1 %>%
  select(
    all_of(
      vars_estudo1
    )
  ) %>%
  padronizar(
    vars_estudo1
  )

# WSS para k = 1,...,8.
wss1 <- data.frame(
  k = 1:8,
  WSS = NA_real_
)

for (i in seq_len(
  nrow(wss1)
)) {

  k_atual <- wss1$k[i]

  set.seed(
    SEED_KMEANS +
      k_atual
  )

  modelo <- kmeans(
    X1,
    centers = k_atual,
    nstart = NSTART_KMEANS,
    iter.max = 100
  )

  wss1$WSS[i] <-
    modelo$tot.withinss
}

exportar_csv(
  wss1,
  "estudo1_wss",
  digits = 2
)

plotar_selecao_k(
  wss1,
  "estudo1_cotovelo",
  "Método do cotovelo - Estudo 1",
  "WSS",
  "WSS"
)

# Silhueta para k = 2,...,6.
sil1 <- data.frame(
  k = K_ESTUDO1,
  silhueta = NA_real_
)

for (i in seq_len(
  nrow(sil1)
)) {

  k_atual <- sil1$k[i]

  set.seed(
    SEED_KMEANS +
      k_atual
  )

  modelo <- kmeans(
    X1,
    centers = k_atual,
    nstart = NSTART_KMEANS,
    iter.max = 100
  )

  sil1$silhueta[i] <-
    silhueta_kmeans(
      X1,
      modelo$cluster
    )
}

exportar_csv(
  sil1,
  "estudo1_silhueta",
  digits = 4
)

plotar_selecao_k(
  sil1,
  "estudo1_silhueta",
  "Silhueta média - Estudo 1",
  "silhueta",
  "Silhueta média"
)

# Solução principal: k = 2.
set.seed(
  SEED_KMEANS
)

modelo1 <- kmeans(
  X1,
  centers = 2,
  nstart = NSTART_KMEANS,
  iter.max = 100
)

estudo1_resultado <-
  estudo1 %>%
  mutate(
    cluster =
      factor(
        modelo1$cluster
      )
  )

perfil1 <-
  estudo1_resultado %>%
  group_by(
    cluster
  ) %>%
  summarise(
    n = n(),
    percentual =
      100 *
      n /
      nrow(
        estudo1_resultado
      ),
    remuneracao_media_real =
      mean(
        remuneracao_reais,
        na.rm = TRUE
      ),
    idade_media =
      mean(
        idade,
        na.rm = TRUE
      ),
    tempo_emprego_medio =
      mean(
        tempo_emprego_m,
        na.rm = TRUE
      ),
    .groups = "drop"
  )

exportar_csv(
  perfil1,
  "estudo1_perfil_clusters",
  digits = 2
)

exportar_rds(
  modelo1,
  "modelo_estudo1_k2"
)

registrar_auditoria(
  "estudo1_silhueta_k2",
  sil1$silhueta[sil1$k == 2],
  observacao = "Resultado recalculado na nova base."
)


# =============================================================================
# 9. ESTUDO 2 - FORMAÇÃO E INSERÇÃO NO NÍVEL INDIVIDUAL
# =============================================================================

log_msg(
  "============================================================"
)

log_msg(
  "INÍCIO DO ESTUDO 2"
)

log_msg(
  "============================================================"
)

estudo2 <- ind_num %>%
  filter(
    !is.na(
      remuneracao_reais
    ),
    remuneracao_reais > 0,
    !is.na(
      idade_media
    ),
    !is.na(
      tempo_emprego_medio
    ),
    !is.na(
      nivel_formacao
    )
  ) %>%
  select(
    id_aluno,
    remuneracao_reais,
    idade_media,
    tempo_emprego_medio,
    nivel_formacao,
    n_anos_observados,
    n_registros
  )

log_msg(
  "Indivíduos do Estudo 2: ",
  nrow(estudo2)
)

registrar_auditoria(
  "estudo2_n_individuos",
  nrow(estudo2),
  observacao = "Nova base pós-conclusão."
)

vars_estudo2 <- c(
  "remuneracao_reais",
  "idade_media",
  "tempo_emprego_medio",
  "nivel_formacao"
)

X2 <- estudo2 %>%
  select(
    all_of(
      vars_estudo2
    )
  ) %>%
  padronizar(
    vars_estudo2
  )

# WSS para k = 1,...,8.
wss2 <- data.frame(
  k = 1:8,
  WSS = NA_real_
)

for (i in seq_len(
  nrow(wss2)
)) {

  k_atual <- wss2$k[i]

  set.seed(
    SEED_KMEANS +
      100 +
      k_atual
  )

  modelo <- kmeans(
    X2,
    centers = k_atual,
    nstart = NSTART_KMEANS,
    iter.max = 100
  )

  wss2$WSS[i] <-
    modelo$tot.withinss
}

exportar_csv(
  wss2,
  "estudo2_wss",
  digits = 2
)

plotar_selecao_k(
  wss2,
  "estudo2_cotovelo",
  "Método do cotovelo - Estudo 2",
  "WSS",
  "WSS"
)

# Silhueta para k = 2,...,6.
sil2 <- data.frame(
  k = K_ESTUDO2,
  silhueta = NA_real_
)

for (i in seq_len(
  nrow(sil2)
)) {

  k_atual <- sil2$k[i]

  set.seed(
    SEED_KMEANS +
      100 +
      k_atual
  )

  modelo <- kmeans(
    X2,
    centers = k_atual,
    nstart = NSTART_KMEANS,
    iter.max = 100
  )

  sil2$silhueta[i] <-
    silhueta_kmeans(
      X2,
      modelo$cluster
    )
}

exportar_csv(
  sil2,
  "estudo2_silhueta",
  digits = 4
)

plotar_selecao_k(
  sil2,
  "estudo2_silhueta",
  "Silhueta média - Estudo 2",
  "silhueta",
  "Silhueta média"
)

# Solução principal: k = 2.
set.seed(
  SEED_KMEANS +
    2000
)

modelo2 <- kmeans(
  X2,
  centers = 2,
  nstart = NSTART_KMEANS,
  iter.max = 100
)

estudo2_resultado <-
  estudo2 %>%
  mutate(
    cluster =
      factor(
        modelo2$cluster
      )
  )

perfil2 <-
  estudo2_resultado %>%
  group_by(
    cluster
  ) %>%
  summarise(
    n = n(),
    percentual =
      100 *
      n /
      nrow(
        estudo2_resultado
      ),
    remuneracao_media_real =
      mean(
        remuneracao_reais,
        na.rm = TRUE
      ),
    remuneracao_mediana_real =
      median(
        remuneracao_reais,
        na.rm = TRUE
      ),
    idade_media =
      mean(
        idade_media,
        na.rm = TRUE
      ),
    tempo_emprego_medio =
      mean(
        tempo_emprego_medio,
        na.rm = TRUE
      ),
    .groups = "drop"
  )

exportar_csv(
  perfil2,
  "estudo2_perfil_clusters",
  digits = 2
)

# Composição educacional.
composicao_formacao2 <-
  estudo2_resultado %>%
  count(
    cluster,
    nivel_formacao
  ) %>%
  group_by(
    cluster
  ) %>%
  mutate(
    percentual =
      100 *
      n /
      sum(n)
  ) %>%
  ungroup() %>%
  mutate(
    nivel_formacao =
      factor(
        nivel_formacao,
        levels = c(
          1,
          2,
          3
        ),
        labels = c(
          "Técnico",
          "Graduação",
          "Pós-graduação"
        )
      )
  )

exportar_csv(
  composicao_formacao2,
  "estudo2_composicao_formacao",
  digits = 2
)

exportar_rds(
  modelo2,
  "modelo_estudo2_k2"
)

registrar_auditoria(
  "estudo2_silhueta_k2",
  sil2$silhueta[sil2$k == 2],
  observacao = "Resultado recalculado na nova base."
)

# -----------------------------------------------------------------------------
# Sensibilidade com log da remuneração
# -----------------------------------------------------------------------------

estudo2_log <-
  estudo2 %>%
  mutate(
    log_remuneracao =
      log(
        remuneracao_reais
      )
  )

vars_estudo2_log <- c(
  "log_remuneracao",
  "idade_media",
  "tempo_emprego_medio",
  "nivel_formacao"
)

X2_log <-
  estudo2_log %>%
  select(
    all_of(
      vars_estudo2_log
    )
  ) %>%
  padronizar(
    vars_estudo2_log
  )

sens_log2 <- data.frame(
  k = 2:3,
  silhueta = NA_real_
)

for (i in seq_len(
  nrow(sens_log2)
)) {

  k_atual <- sens_log2$k[i]

  set.seed(
    SEED_KMEANS +
      3000 +
      k_atual
  )

  modelo <- kmeans(
    X2_log,
    centers = k_atual,
    nstart = NSTART_KMEANS,
    iter.max = 100
  )

  sens_log2$silhueta[i] <-
    silhueta_kmeans(
      X2_log,
      modelo$cluster
    )
}

exportar_csv(
  sens_log2,
  "estudo2_sensibilidade_log",
  digits = 4
)

registrar_auditoria(
  "estudo2_log_silhueta_k2",
  sens_log2$silhueta[sens_log2$k == 2],
  observacao = "Sensibilidade recalculada na nova base."
)

registrar_auditoria(
  "estudo2_log_silhueta_k3",
  sens_log2$silhueta[sens_log2$k == 3],
  observacao = "Sensibilidade recalculada na nova base."
)


# =============================================================================
# 10. ESTUDO 3 - FORMAÇÃO, INSERÇÃO E DESCRIÇÃO OCUPACIONAL
# =============================================================================

log_msg(
  "============================================================"
)

log_msg(
  "INÍCIO DO ESTUDO 3"
)

log_msg(
  "============================================================"
)

# Ocupação predominante:
#   1. maior frequência da ocupação selecionada para cada pessoa-ano;
#   2. em empate, ano mais recente;
#   3. em empate dentro do ano, a moda pessoa-ano já usa desempate alfabético.

ocup_contagem <-
  base_pessoa_ano %>%
  filter(
    !is.na(
      id_aluno
    ),
    !is.na(
      desc_ocupacao
    ),
    desc_ocupacao != ""
  ) %>%
  count(
    id_aluno,
    desc_ocupacao,
    name = "n_ocorrencias"
  )

ocup_ultimo <-
  base_pessoa_ano %>%
  filter(
    !is.na(
      id_aluno
    ),
    !is.na(
      desc_ocupacao
    ),
    desc_ocupacao != "",
    !is.na(
      ano_referencia
    )
  ) %>%
  group_by(
    id_aluno,
    desc_ocupacao
  ) %>%
  summarise(
    ano_ultimo =
      max(
        ano_referencia,
        na.rm = TRUE
      ),
    .groups = "drop"
  )

ocup_predominante <-
  ocup_contagem %>%
  left_join(
    ocup_ultimo,
    by = c(
      "id_aluno",
      "desc_ocupacao"
    )
  ) %>%
  arrange(
    id_aluno,
    desc(
      n_ocorrencias
    ),
    desc(
      ano_ultimo
    ),
    desc_ocupacao
  ) %>%
  distinct(
    id_aluno,
    .keep_all = TRUE
  ) %>%
  select(
    id_aluno,
    desc_ocupacao
  )

estudo3 <-
  ind_num %>%
  select(
    id_aluno,
    remuneracao_reais,
    idade_media,
    tempo_emprego_medio,
    nivel_formacao,
    n_anos_observados,
    n_registros
  ) %>%
  left_join(
    ocup_predominante,
    by = "id_aluno"
  ) %>%
  filter(
    !is.na(
      desc_ocupacao
    )
  )

log_msg(
  "Indivíduos do Estudo 3: ",
  nrow(estudo3)
)

log_msg(
  "Descrições ocupacionais distintas: ",
  n_distinct(
    estudo3$desc_ocupacao
  )
)

registrar_auditoria(
  "estudo3_n_individuos",
  nrow(estudo3),
  observacao = "Nova base pós-conclusão."
)

registrar_auditoria(
  "estudo3_n_descricoes_ocupacionais",
  n_distinct(estudo3$desc_ocupacao),
  observacao = "Nova base pós-conclusão."
)

freq_ocupacoes <-
  estudo3 %>%
  count(
    desc_ocupacao,
    sort = TRUE
  )

exportar_csv(
  freq_ocupacoes,
  "estudo3_frequencia_ocupacoes"
)

vars_num3 <- c(
  "remuneracao_reais",
  "idade_media",
  "tempo_emprego_medio",
  "nivel_formacao"
)

X3 <-
  estudo3 %>%
  mutate(
    across(
      all_of(
        vars_num3
      ),
      ~ as.numeric(
        scale(.x)
      )
    ),
    desc_ocupacao =
      factor(
        desc_ocupacao
      )
  ) %>%
  select(
    all_of(
      vars_num3
    ),
    desc_ocupacao
  )

# -----------------------------------------------------------------------------
# Custos para k = 1,...,6
# -----------------------------------------------------------------------------

custos3 <- data.frame(
  k = 1:6,
  custo = NA_real_
)

for (i in seq_len(
  nrow(custos3)
)) {

  k_atual <- custos3$k[i]

  set.seed(
    SEED_KPROTO +
      k_atual
  )

  modelo <- kproto(
    X3,
    k = k_atual,
    lambda =
      LAMBDA_PRINCIPAL,
    nstart =
      NSTART_KPROTO_EXPLORACAO
  )

  custos3$custo[i] <-
    modelo$tot.withinss
}

exportar_csv(
  custos3,
  "estudo3_custo_kproto",
  digits = 2
)

plotar_selecao_k(
  custos3,
  "estudo3_custo_kproto",
  "Custo do K-prototypes - Estudo 3",
  "custo",
  "Custo intracluster"
)

# -----------------------------------------------------------------------------
# Silhueta em amostra determinística de até 1.500 indivíduos
# -----------------------------------------------------------------------------

set.seed(
  SEED_SILHUETA
)

n_amostra3 <-
  min(
    N_AMOSTRA_SILHUETA_ESTUDO3,
    nrow(X3)
  )

amostra3 <-
  sample(
    seq_len(
      nrow(X3)
    ),
    n_amostra3
  )

write.csv(
  data.frame(
    indice_amostra =
      amostra3
  ),
  file.path(
    DIR_AUDITORIA,
    "estudo3_indices_amostra_silhueta.csv"
  ),
  row.names = FALSE
)

log_msg(
  "Amostra determinística da silhueta do Estudo 3: ",
  n_amostra3,
  " indivíduos."
)

sil3 <- data.frame(
  k = K_ESTUDO3,
  silhueta = NA_real_
)

for (i in seq_len(
  nrow(sil3)
)) {

  k_atual <- sil3$k[i]

  set.seed(
    SEED_KPROTO +
      k_atual
  )

  modelo <- kproto(
    X3,
    k = k_atual,
    lambda =
      LAMBDA_PRINCIPAL,
    nstart =
      NSTART_KPROTO_EXPLORACAO
  )

  sil3$silhueta[i] <-
    calcular_silhueta_mista(
      X =
        X3[
          amostra3,
          ,
          drop = FALSE
        ],
      clusters =
        modelo$cluster[
          amostra3
        ],
      vars_num =
        vars_num3,
      lambda =
        LAMBDA_PRINCIPAL
    )
}

exportar_csv(
  sil3,
  "estudo3_silhueta",
  digits = 4
)

plotar_selecao_k(
  sil3,
  "estudo3_silhueta",
  "Silhueta média - Estudo 3",
  "silhueta",
  "Silhueta média"
)

# -----------------------------------------------------------------------------
# Modelo principal: k = 2, lambda = 1
# -----------------------------------------------------------------------------

set.seed(
  SEED_KPROTO
)

modelo3 <- kproto(
  X3,
  k = 2,
  lambda =
    LAMBDA_PRINCIPAL,
  nstart =
    NSTART_KPROTO_FINAL
)

# -----------------------------------------------------------------------------
# Teste automático de reprodutibilidade do modelo principal
# -----------------------------------------------------------------------------
# O mesmo conjunto de dados, parâmetros e seed deve produzir a mesma solução.
# Se clusters ou custo intracluster divergirem, a execução é interrompida.
set.seed(SEED_KPROTO)

modelo3_reproducao <- kproto(
  X3,
  k = 2,
  lambda =
    LAMBDA_PRINCIPAL,
  nstart =
    NSTART_KPROTO_FINAL
)

clusters_identicos <-
  identical(
    modelo3$cluster,
    modelo3_reproducao$cluster
  )

custo_identico <-
  isTRUE(
    all.equal(
      modelo3$tot.withinss,
      modelo3_reproducao$tot.withinss,
      tolerance = 1e-12
    )
  )

if (!clusters_identicos || !custo_identico) {
  stop(
    paste0(
      "FALHA NO TESTE DE REPRODUTIBILIDADE DO ESTUDO 3: ",
      "duas execuções controladas do kproto() produziram resultados diferentes. ",
      "Verifique a versão do R, clustMixType, demais pacotes, a base de entrada ",
      "e a ordem das observações."
    )
  )
}

log_msg(
  "Teste de reprodutibilidade do Estudo 3: OK. ",
  "Clusters e custo intracluster foram reproduzidos exatamente."
)

registrar_auditoria(
  "estudo3_reprodutibilidade_kproto",
  "OK",
  "OK",
  observacao =
    "Duas execuções independentes do modelo principal, com mesma seed, dados e parâmetros, produziram a mesma solução."
)

estudo3_resultado <-
  estudo3 %>%
  mutate(
    cluster =
      factor(
        modelo3$cluster
      )
  )

perfil3 <-
  estudo3_resultado %>%
  group_by(
    cluster
  ) %>%
  summarise(
    n_individuos =
      n(),
    percentual =
      100 *
      n_individuos /
      nrow(
        estudo3_resultado
      ),
    remuneracao_media =
      mean(
        remuneracao_reais,
        na.rm = TRUE
      ),
    remuneracao_mediana =
      median(
        remuneracao_reais,
        na.rm = TRUE
      ),
    idade_media =
      mean(
        idade_media,
        na.rm = TRUE
      ),
    tempo_emprego_medio =
      mean(
        tempo_emprego_medio,
        na.rm = TRUE
      ),
    nivel_formacao_medio =
      mean(
        nivel_formacao,
        na.rm = TRUE
      ),
    anos_observados_medio =
      mean(
        n_anos_observados,
        na.rm = TRUE
      ),
    .groups = "drop"
  )

exportar_csv(
  perfil3,
  "estudo3_perfil_clusters",
  digits = 2
)

composicao3 <-
  estudo3_resultado %>%
  count(
    cluster,
    desc_ocupacao
  ) %>%
  group_by(
    cluster
  ) %>%
  mutate(
    percentual =
      100 *
      n /
      sum(n)
  ) %>%
  arrange(
    cluster,
    desc(
      percentual
    )
  ) %>%
  ungroup()

exportar_csv(
  composicao3,
  "estudo3_composicao_ocupacional",
  digits = 2
)

exportar_rds(
  modelo3,
  "modelo_estudo3_k2_lambda1"
)

registrar_auditoria(
  "estudo3_silhueta_k2",
  sil3$silhueta[sil3$k == 2],
  observacao = "Silhueta calculada na nova base; sem comparação com o TCC anterior."
)

# -----------------------------------------------------------------------------
# Sensibilidade ao lambda
# -----------------------------------------------------------------------------

sens3 <- data.frame(
  lambda =
    LAMBDA_SENSIBILIDADE,
  custo = NA_real_,
  silhueta = NA_real_
)

for (i in seq_len(
  nrow(sens3)
)) {

  lambda_atual <-
    sens3$lambda[i]

  set.seed(
    SEED_KPROTO +
      i
  )

  modelo_sens <- kproto(
    X3,
    k = 2,
    lambda =
      lambda_atual,
    nstart =
      NSTART_KPROTO_EXPLORACAO
  )

  sens3$custo[i] <-
    modelo_sens$tot.withinss

  sens3$silhueta[i] <-
    calcular_silhueta_mista(
      X =
        X3[
          amostra3,
          ,
          drop = FALSE
        ],
      clusters =
        modelo_sens$cluster[
          amostra3
        ],
      vars_num =
        vars_num3,
      lambda =
        lambda_atual
    )
}

exportar_csv(
  sens3,
  "estudo3_sensibilidade_lambda",
  digits = 4
)

# Auditoria descritiva da sensibilidade: os valores são registrados como
# resultados da nova base, sem comparação com números do TCC anterior.
for (lambda_ref in LAMBDA_SENSIBILIDADE) {
  valor_calc <- sens3$silhueta[sens3$lambda == lambda_ref]
  registrar_auditoria(
    paste0("estudo3_sensibilidade_lambda_", lambda_ref),
    valor_calc,
    observacao = paste0(
      "Sensibilidade calculada na nova base para lambda=", lambda_ref,
      "; sem comparação com o TCC anterior."
    )
  )
}


# =============================================================================
# 11. ESTUDO 4 - PERFIS INSTITUCIONAIS DOS CAMPI
# =============================================================================

log_msg(
  "============================================================"
)

log_msg(
  "INÍCIO DO ESTUDO 4"
)

log_msg(
  "============================================================"
)

# O Estudo 4 tem unidade de análise = campus.
#
# A agregação é feita sobre unidades pessoa-ano já consolidadas. Assim, cada
# vínculo individual deixa de multiplicar o peso do egresso dentro do mesmo ano.
# A média por campus continua ponderada por pessoa-ano: egressos observados em
# mais anos contribuem com mais unidades anuais. Essa definição é registrada
# para interpretação dos indicadores institucionais.
#
# Variáveis:
#   remuneração média real;
#   idade média;
#   tempo médio de emprego;
#   nível médio de formação.

estudo4 <- base_para_individual %>%
  filter(
    !is.na(
      campus
    ),
    campus != ""
  ) %>%
  group_by(
    campus
  ) %>%
  summarise(
    remuneracao_media =
      mean(
        remuneracao_reais,
        na.rm = TRUE
      ),
    idade_media =
      mean(
        idade,
        na.rm = TRUE
      ),
    tempo_emprego_medio =
      mean(
        tempo_emprego_m,
        na.rm = TRUE
      ),
    nivel_formacao_medio =
      mean(
        valor_nivel_categoria,
        na.rm = TRUE
      ),
    n_pessoa_anos =
      n(),
    n_registros_fonte =
      sum(n_registros_fonte, na.rm = TRUE),
    n_individuos =
      n_distinct(
        id_aluno
      ),
    .groups = "drop"
  )

log_msg(
  "Campi identificados no Estudo 4: ",
  nrow(estudo4)
)

registrar_auditoria(
  "estudo4_n_campi",
  nrow(estudo4),
  observacao = "Nova base pós-conclusão."
)

exportar_csv(
  estudo4,
  "estudo4_base_campi",
  digits = 2
)

vars_estudo4 <- c(
  "remuneracao_media",
  "idade_media",
  "tempo_emprego_medio",
  "nivel_formacao_medio"
)

X4 <-
  estudo4 %>%
  select(
    all_of(
      vars_estudo4
    )
  ) %>%
  padronizar(
    vars_estudo4
  )

# WSS para k = 1,...,5.
wss4 <- data.frame(
  k = 1:5,
  WSS = NA_real_
)

for (i in seq_len(
  nrow(wss4)
)) {

  k_atual <- wss4$k[i]

  set.seed(
    SEED_KMEANS +
      4000 +
      k_atual
  )

  modelo <- kmeans(
    X4,
    centers = k_atual,
    nstart =
      NSTART_KMEANS,
    iter.max = 100
  )

  wss4$WSS[i] <-
    modelo$tot.withinss
}

exportar_csv(
  wss4,
  "estudo4_wss",
  digits = 4
)

plotar_selecao_k(
  wss4,
  "estudo4_cotovelo",
  "Método do cotovelo - Estudo 4",
  "WSS",
  "WSS"
)

# Silhueta para k = 2,...,5.
sil4 <- data.frame(
  k = K_ESTUDO4,
  silhueta = NA_real_
)

for (i in seq_len(
  nrow(sil4)
)) {

  k_atual <- sil4$k[i]

  set.seed(
    SEED_KMEANS +
      5000 +
      k_atual
  )

  modelo <- kmeans(
    X4,
    centers = k_atual,
    nstart =
      NSTART_KMEANS,
    iter.max = 100
  )

  sil4$silhueta[i] <-
    silhueta_kmeans(
      X4,
      modelo$cluster
    )
}

exportar_csv(
  sil4,
  "estudo4_silhueta",
  digits = 4
)

plotar_selecao_k(
  sil4,
  "estudo4_silhueta",
  "Silhueta média - Estudo 4",
  "silhueta",
  "Silhueta média"
)

# Modelo principal: k = 2.
set.seed(
  SEED_KMEANS +
    6000
)

modelo4 <- kmeans(
  X4,
  centers = 2,
  nstart =
    NSTART_KMEANS,
  iter.max = 100
)

estudo4_resultado <-
  estudo4 %>%
  mutate(
    cluster =
      factor(
        modelo4$cluster
      )
  )

exportar_csv(
  estudo4_resultado,
  "estudo4_campi_clusters",
  digits = 2
)

perfil4 <-
  estudo4_resultado %>%
  group_by(
    cluster
  ) %>%
  summarise(
    n_campi =
      n(),
    remuneracao_media =
      mean(
        remuneracao_media,
        na.rm = TRUE
      ),
    idade_media =
      mean(
        idade_media,
        na.rm = TRUE
      ),
    tempo_emprego_medio =
      mean(
        tempo_emprego_medio,
        na.rm = TRUE
      ),
    nivel_formacao_medio =
      mean(
        nivel_formacao_medio,
        na.rm = TRUE
      ),
    .groups = "drop"
  )

exportar_csv(
  perfil4,
  "estudo4_perfil_clusters",
  digits = 2
)

exportar_rds(
  modelo4,
  "modelo_estudo4_k2"
)

registrar_auditoria(
  "estudo4_silhueta_k2",
  sil4$silhueta[sil4$k == 2],
  observacao = "Resultado recalculado na nova base."
)


# =============================================================================
# 12. AUDITORIA CONSOLIDADA
# =============================================================================

log_msg(
  "Consolidando auditoria."
)

auditoria_df <-
  bind_rows(
    auditoria
  )

write.csv(
  auditoria_df,
  file.path(
    DIR_AUDITORIA,
    "auditoria_resultados.csv"
  ),
  row.names = FALSE,
  fileEncoding = "UTF-8"
)

divergencias <-
  auditoria_df %>%
  filter(
    status ==
      "DIVERGENCIA"
  )

write.csv(
  divergencias,
  file.path(
    DIR_AUDITORIA,
    "divergencias.csv"
  ),
  row.names = FALSE,
  fileEncoding = "UTF-8"
)

if (
  nrow(divergencias) == 0
) {

  log_msg(
    "Nenhuma divergência encontrada nos itens auditados."
  )

} else {

  log_msg(
    nrow(divergencias),
    " divergência(s) encontrada(s). Consulte auditoria_resultados.csv.",
    nivel = "WARN"
  )
}


# =============================================================================
# 13. RESUMO NUMÉRICO DOS QUATRO ESTUDOS
# =============================================================================

resumo_estudos <- data.frame(
  cenario_pos_conclusao = rep(CENARIO_POS_CONCLUSAO, 4),
  definicao_cenario = rep(DESCRICAO_CENARIO_POS, 4),
  estudo = c(
    "Estudo 1",
    "Estudo 2",
    "Estudo 3",
    "Estudo 4"
  ),
  metodo = c(
    "K-means",
    "K-means",
    "K-prototypes",
    "K-means"
  ),
  unidade = c(
    "Observação anual",
    "Indivíduo",
    "Indivíduo",
    "Campus"
  ),
  k_principal = c(
    2,
    2,
    2,
    2
  ),
  silhueta = c(
    sil1$silhueta[
      sil1$k == 2
    ],
    sil2$silhueta[
      sil2$k == 2
    ],
    sil3$silhueta[
      sil3$k == 2
    ],
    sil4$silhueta[
      sil4$k == 2
    ]
  )
)

exportar_csv(
  resumo_estudos,
  "resumo_estudos",
  digits = 4
)


# =============================================================================
# 14. GENEALOGIA DAS BASES
# =============================================================================

genealogia_bases <- data.frame(
  objeto = c(
    "DADOS",
    "base_limpa",
    "base_pessoa_ano",
    "ind_num",
    "estudo1",
    "X1",
    "estudo2",
    "X2",
    "estudo3",
    "X3",
    "estudo4",
    "X4"
  ),
  descricao = c(
    "Base original lida do XLSX.",
    "Base com tipos normalizados e remuneração real em preços de 2024.",
    "Base consolidada por egresso-ano; remuneração reconstruída por mês ativo, somando apenas vínculos simultâneos em cada mês.",
    "Base individual consolidada a partir de pessoa-ano.",
    "Base anual do Estudo 1.",
    "Matriz padronizada do Estudo 1.",
    "Base individual do Estudo 2.",
    "Matriz padronizada do Estudo 2.",
    "Base individual com ocupação predominante.",
    "Matriz mista padronizada do Estudo 3.",
    "Base agregada por campus.",
    "Matriz padronizada do Estudo 4."
  ),
  n_linhas = c(
    nrow(DADOS),
    nrow(base_limpa),
    nrow(base_pessoa_ano),
    nrow(ind_num),
    nrow(estudo1),
    nrow(X1),
    nrow(estudo2),
    nrow(X2),
    nrow(estudo3),
    nrow(X3),
    nrow(estudo4),
    nrow(X4)
  ),
  n_colunas = c(
    ncol(DADOS),
    ncol(base_limpa),
    ncol(base_pessoa_ano),
    ncol(ind_num),
    ncol(estudo1),
    ncol(X1),
    ncol(estudo2),
    ncol(X2),
    ncol(estudo3),
    ncol(X3),
    ncol(estudo4),
    ncol(X4)
  )
)

write.csv(
  genealogia_bases,
  file.path(
    DIR_AUDITORIA,
    "genealogia_bases.csv"
  ),
  row.names = FALSE,
  fileEncoding = "UTF-8"
)


# =============================================================================
# 15. METADADOS E AMBIENTE COMPUTACIONAL
# =============================================================================

metadados_execucao <- data.frame(
  campo = c(
    "data_hora_fim",
    "arquivo_base",
    "cenario_pos_conclusao",
    "definicao_cenario",
    "nota_limite_rais_anual",
    "ano_base",
    "anos_analise",
    "seed_kmeans",
    "seed_kproto",
    "seed_silhueta",
    "nstart_kmeans",
    "nstart_kproto_exploracao",
    "nstart_kproto_final",
    "lambda_principal",
    "lambda_sensibilidade",
    "amostra_silhueta_estudo3",
    "R.version",
    "sistema_operacional"
  ),
  valor = c(
    as.character(
      Sys.time()
    ),
    normalizePath(
      ARQUIVO_BASE,
      mustWork = FALSE
    ),
    CENARIO_POS_CONCLUSAO,
    DESCRICAO_CENARIO_POS,
    "No cenário estrito, o ano da conclusão é excluído porque a RAIS anual não permite separar remuneração pré e pós-formatura dentro desse ano.",
    ANO_BASE,
    paste(
      ANOS_ANALISE,
      collapse = ","
    ),
    SEED_KMEANS,
    SEED_KPROTO,
    SEED_SILHUETA,
    NSTART_KMEANS,
    NSTART_KPROTO_EXPLORACAO,
    NSTART_KPROTO_FINAL,
    LAMBDA_PRINCIPAL,
    paste(
      LAMBDA_SENSIBILIDADE,
      collapse = ","
    ),
    n_amostra3,
    R.version.string,
    Sys.info()[["sysname"]]
  ),
  stringsAsFactors = FALSE
)

write.csv(
  metadados_execucao,
  file.path(
    DIR_AUDITORIA,
    "metadados_execucao.csv"
  ),
  row.names = FALSE,
  fileEncoding = "UTF-8"
)

capture.output(
  sessionInfo(),
  file = file.path(
    DIR_AUDITORIA,
    "sessionInfo.txt"
  )
)


# =============================================================================
# 16. EXPORTAÇÃO OPCIONAL DE MICRODADOS SANITIZADOS
# =============================================================================
#
# ATENÇÃO:
# Mesmo sem CPF/CNPJ, uma combinação de idade, remuneração, ocupação, ano,
# campus etc. pode permitir reidentificação.
#
# Por isso, o padrão do script é FALSE.
# Para publicação científica, recomenda-se publicar apenas:
#   * tabelas agregadas;
#   * parâmetros;
#   * scripts;
#   * logs;
#   * metadados;
#   * bases auxiliares agregadas.
#
# A decisão de publicar microdados deve ser autorizada pela instituição e
# compatível com o convênio e com a LGPD.

if (
  PUBLICAR_MICRODADOS
) {

  estudo1_publicavel <-
    estudo1 %>%
    select(
      ano_referencia,
      remuneracao_reais,
      log_remuneracao,
      idade,
      tempo_emprego_m
    )

  estudo2_publicavel <-
    estudo2 %>%
    select(
      remuneracao_reais,
      idade_media,
      tempo_emprego_medio,
      nivel_formacao,
      n_anos_observados,
      n_registros
    )

  estudo3_publicavel <-
    estudo3 %>%
    select(
      remuneracao_reais,
      idade_media,
      tempo_emprego_medio,
      nivel_formacao,
      n_anos_observados,
      n_registros,
      desc_ocupacao
    )

  write.csv(
    estudo1_publicavel,
    file.path(
      DIR_DADOS,
      "estudo1_microdados_sanitizados.csv"
    ),
    row.names = FALSE,
    fileEncoding = "UTF-8"
  )

  write.csv(
    estudo2_publicavel,
    file.path(
      DIR_DADOS,
      "estudo2_microdados_sanitizados.csv"
    ),
    row.names = FALSE,
    fileEncoding = "UTF-8"
  )

  write.csv(
    estudo3_publicavel,
    file.path(
      DIR_DADOS,
      "estudo3_microdados_sanitizados.csv"
    ),
    row.names = FALSE,
    fileEncoding = "UTF-8"
  )

  log_msg(
    "Microdados sanitizados exportados. Revisar risco de reidentificação antes de publicar.",
    nivel = "WARN"
  )
}


# =============================================================================
# 17. CHECKLIST FINAL DE EXECUÇÃO
# =============================================================================

checklist_final <- data.frame(
  item = c(
    "Base de entrada encontrada",
    "Colunas obrigatórias presentes",
    "Anos 2020-2024 disponíveis na base original",
    "Anos 2022-2024 disponíveis na base analítica estrita",
    "Fatores IPCA disponíveis",
    "Estudo 1 executado",
    "Estudo 2 executado",
    "Sensibilidade log do Estudo 2 executada",
    "Estudo 3 executado",
    "Sensibilidade lambda do Estudo 3 executada",
    "Estudo 4 executado sobre pessoa-ano consolidado",
    "Auditoria consolidada gerada",
    "sessionInfo gerado"
  ),
  status = c(
    file.exists(
      ARQUIVO_BASE
    ),
    TRUE,
    all(
      2020:2024 %in% as.integer(DADOS$ano_referencia)
    ),
    all(
      ANOS_POS_ESPERADOS %in% anos_encontrados
    ),
    n_sem_fator == 0,
    exists(
      "modelo1"
    ),
    exists(
      "modelo2"
    ),
    exists(
      "sens_log2"
    ),
    exists(
      "modelo3"
    ),
    exists(
      "sens3"
    ),
    exists(
      "modelo4"
    ),
    exists(
      "auditoria_df"
    ),
    file.exists(
      file.path(
        DIR_AUDITORIA,
        "sessionInfo.txt"
      )
    )
  )
)

write.csv(
  checklist_final,
  file.path(
    DIR_AUDITORIA,
    "checklist_final.csv"
  ),
  row.names = FALSE,
  fileEncoding = "UTF-8"
)

if (
  all(
    checklist_final$status
  )
) {

  log_msg(
    "CHECKLIST FINAL: todos os itens principais foram concluídos."
  )

} else {

  log_msg(
    "CHECKLIST FINAL: há itens não concluídos. Consulte checklist_final.csv.",
    nivel = "WARN"
  )
}


# =============================================================================
# 18. ENCERRAMENTO
# =============================================================================

log_msg(
  "============================================================"
)

log_msg(
  "FIM DA EXECUÇÃO"
)

log_msg(
  "============================================================"
)

log_msg(
  "Resultados gravados em: ",
  normalizePath(
    DIR_SAIDA,
    mustWork = FALSE
  )
)

log_msg(
  "Log principal: ",
  normalizePath(
    ARQUIVO_LOG,
    mustWork = FALSE
  )
)

# O on.exit() fecha a conexão do log.
