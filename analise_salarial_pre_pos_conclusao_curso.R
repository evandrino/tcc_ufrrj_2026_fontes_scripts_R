# Pacotes usados
suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
})

# Arquivos de entrada e saída
DIR_PROJETO <- "Z:/UFRRJ/TCC/TCC_Clusterizacao"

ARQUIVO_BASE <- file.path(
  DIR_PROJETO,
  "base_cefet_mg_rais_2021_2024_egressos_com_1_curso_ate_2024.xlsx"
)

DIR_SAIDA <- file.path(
  DIR_PROJETO,
  "output_clusterizacao_tabela4_v3",
  "trajetoria_salarial_pre_pos"
)

if (!file.exists(ARQUIVO_BASE)) {
  stop("Planilha não encontrada: ", ARQUIVO_BASE)
}

dir.create(DIR_SAIDA, recursive = TRUE, showWarnings = FALSE)

# Fatores para converter remunerações em reais médios de 2024
FATORES_IPCA <- data.frame(
  ano_referencia = 2020:2024,
  fator_2024 = c(1.2919518, 1.1929197, 1.0916165, 1.0436746, 1.0000000)
)

# Moda determinística para recuperar os atributos de cada egresso
moda_deterministica <- function(x) {
  x <- trimws(as.character(x))
  x <- x[!is.na(x) & nzchar(x)]
  
  if (length(x) == 0L) {
    return(NA_character_)
  }
  
  frequencias <- table(x)
  candidatas <- names(frequencias)[frequencias == max(frequencias)]
  sort(candidatas)[1]
}

# Formatação brasileira para as tabelas de apresentação
formatar_reais <- function(x) {
  vapply(x, function(valor) {
    if (is.na(valor)) {
      return(NA_character_)
    }
    paste0(
      "R$ ",
      formatC(
        valor,
        format = "f",
        digits = 2,
        big.mark = ".",
        decimal.mark = ","
      )
    )
  }, character(1), USE.NAMES = FALSE)
}

exportar_csv_reais <- function(dados, colunas_reais, arquivo) {
  dados_formatados <- dados
  dados_formatados[colunas_reais] <-
    lapply(dados_formatados[colunas_reais], formatar_reais)

  write.csv2(
    dados_formatados,
    arquivo,
    row.names = FALSE,
    fileEncoding = "UTF-8",
    na = ""
  )
}

# Leitura e validação da planilha
DADOS <- read_excel(
  ARQUIVO_BASE,
  sheet = 1,
  .name_repair = "unique"
)

colunas_obrigatorias <- c(
  "id_aluno", "id_curso", "ano_conclusao", "ano_referencia",
  "remun_media_nom", "campus", "nome_curso", "nivel_categoria"
)

colunas_faltantes <- setdiff(colunas_obrigatorias, names(DADOS))

if (length(colunas_faltantes) > 0L) {
  stop(
    "Colunas ausentes na planilha: ",
    paste(colunas_faltantes, collapse = ", ")
  )
}

# Preparação, filtros e correção monetária
base <- DADOS %>%
  mutate(
    ano_referencia = as.integer(ano_referencia),
    ano_conclusao = as.integer(ano_conclusao),
    remun_media_nom = as.numeric(remun_media_nom),
    campus = trimws(as.character(campus)),
    nome_curso = trimws(as.character(nome_curso)),
    nivel_categoria = trimws(as.character(nivel_categoria))
  ) %>%
  filter(
    !is.na(id_aluno),
    !is.na(ano_conclusao),
    ano_conclusao >= 2021,
    ano_conclusao <= 2024,
    !is.na(remun_media_nom),
    remun_media_nom > 0,
    ano_referencia %in% 2020:2024
  ) %>%
  group_by(id_aluno) %>%
  filter(n_distinct(id_curso, na.rm = TRUE) <= 1L) %>%
  ungroup() %>%
  mutate(remuneracao_nominal_arredondada = round(remun_media_nom)) %>%
  left_join(FATORES_IPCA, by = "ano_referencia") %>%
  mutate(
    remuneracao_reais_2024 =
      remuneracao_nominal_arredondada * fator_2024
  )

if (any(is.na(base$fator_2024))) {
  stop("Há registros sem fator IPCA. Verifique os anos de referência.")
}

# Atributos por egresso
meta <- base %>%
  group_by(id_aluno) %>%
  summarise(
    ano_conclusao = first(ano_conclusao),
    campus = moda_deterministica(campus),
    curso = moda_deterministica(nome_curso),
    nivel_formacao = moda_deterministica(nivel_categoria),
    .groups = "drop"
  )

# Consolidação dos vínculos em uma remuneração por egresso-ano
pessoa_ano <- base %>%
  group_by(id_aluno, ano_referencia) %>%
  summarise(
    ano_conclusao = first(ano_conclusao),
    remuneracao_reais_2024 = sum(remuneracao_reais_2024, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  mutate(
    tempo_relativo_conclusao = ano_referencia - ano_conclusao,
    periodo = case_when(
      tempo_relativo_conclusao < 0 ~ "pre",
      tempo_relativo_conclusao == 0 ~ "ano_conclusao_ambiguo",
      tempo_relativo_conclusao > 0 ~ "pos"
    )
  ) %>%
  left_join(
    meta %>% select(id_aluno, campus, curso, nivel_formacao),
    by = "id_aluno"
  )

# Média individual da remuneração nos anos pré e pós-conclusão
pre <- pessoa_ano %>%
  filter(periodo == "pre") %>%
  group_by(id_aluno) %>%
  summarise(
    remuneracao_media_pre = mean(remuneracao_reais_2024, na.rm = TRUE),
    .groups = "drop"
  )

pos <- pessoa_ano %>%
  filter(periodo == "pos") %>%
  group_by(id_aluno) %>%
  summarise(
    remuneracao_media_pos = mean(remuneracao_reais_2024, na.rm = TRUE),
    .groups = "drop"
  )

# Amostra pareada: cada egresso tem ao menos uma observação pré e uma pós
pareados <- inner_join(pre, pos, by = "id_aluno") %>%
  left_join(meta, by = "id_aluno") %>%
  mutate(
    diferenca_individual =
      remuneracao_media_pos - remuneracao_media_pre
  )

if (nrow(pareados) == 0L) {
  stop("Não há egressos com observações salariais nos períodos pré e pós.")
}

# Tabela 4: n, medianas e intervalo interquartil (Q1–Q3)
resumo_nivel_formacao <- pareados %>%
  group_by(nivel_formacao) %>%
  summarise(
    n_egressos = n_distinct(id_aluno),
    
    mediana_pre = median(remuneracao_media_pre, na.rm = TRUE),
    q1_pre = unname(quantile(remuneracao_media_pre, 0.25, na.rm = TRUE)),
    q3_pre = unname(quantile(remuneracao_media_pre, 0.75, na.rm = TRUE)),
    
    mediana_pos = median(remuneracao_media_pos, na.rm = TRUE),
    q1_pos = unname(quantile(remuneracao_media_pos, 0.25, na.rm = TRUE)),
    q3_pos = unname(quantile(remuneracao_media_pos, 0.75, na.rm = TRUE)),
    
    mediana_diferenca_individual =
      median(diferenca_individual, na.rm = TRUE),
    q1_diferenca_individual =
      unname(quantile(diferenca_individual, 0.25, na.rm = TRUE)),
    q3_diferenca_individual =
      unname(quantile(diferenca_individual, 0.75, na.rm = TRUE)),
    
    .groups = "drop"
  ) %>%
  arrange(factor(
    nivel_formacao,
    levels = c("Técnico", "Graduação", "Pós Graduação")
  ))

resumo_total <- pareados %>%
  summarise(
    nivel_formacao = "Total",
    n_egressos = n_distinct(id_aluno),
    mediana_pre = median(remuneracao_media_pre, na.rm = TRUE),
    q1_pre = unname(quantile(remuneracao_media_pre, 0.25, na.rm = TRUE)),
    q3_pre = unname(quantile(remuneracao_media_pre, 0.75, na.rm = TRUE)),
    mediana_pos = median(remuneracao_media_pos, na.rm = TRUE),
    q1_pos = unname(quantile(remuneracao_media_pos, 0.25, na.rm = TRUE)),
    q3_pos = unname(quantile(remuneracao_media_pos, 0.75, na.rm = TRUE)),
    mediana_diferenca_individual =
      median(diferenca_individual, na.rm = TRUE),
    q1_diferenca_individual =
      unname(quantile(diferenca_individual, 0.25, na.rm = TRUE)),
    q3_diferenca_individual =
      unname(quantile(diferenca_individual, 0.75, na.rm = TRUE))
  )

resumo_nivel_formacao <- bind_rows(resumo_nivel_formacao, resumo_total)

# Preserva os dados numéricos com precisão completa
write.csv(
  resumo_nivel_formacao,
  file.path(DIR_SAIDA, "resumo_nivel_formacao_numerico.csv"),
  row.names = FALSE,
  fileEncoding = "UTF-8"
)

# CSV de apresentação, com valores monetários em R$ e duas casas decimais
exportar_csv_reais(
  resumo_nivel_formacao,
  colunas_reais = c(
    "mediana_pre", "q1_pre", "q3_pre",
    "mediana_pos", "q1_pos", "q3_pos",
    "mediana_diferenca_individual",
    "q1_diferenca_individual", "q3_diferenca_individual"
  ),
  arquivo = file.path(DIR_SAIDA, "resumo_nivel_formacao.csv")
)

# Tabela 5: janela balanceada em t = -1, t = 0 e t = 1
janela_balanceada <- pessoa_ano %>%
  filter(
    ano_conclusao %in% 2021:2023,
    tempo_relativo_conclusao %in% c(-1L, 0L, 1L)
  ) %>%
  group_by(id_aluno) %>%
  filter(n_distinct(tempo_relativo_conclusao) == 3L) %>%
  ungroup()

resumo_janela_balanceada <- janela_balanceada %>%
  group_by(ano_conclusao, tempo_relativo_conclusao) %>%
  summarise(
    n_egressos = n_distinct(id_aluno),
    remuneracao_mediana_reais_2024 =
      median(remuneracao_reais_2024, na.rm = TRUE),
    remuneracao_media_reais_2024 =
      mean(remuneracao_reais_2024, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  arrange(ano_conclusao, tempo_relativo_conclusao)

# Preserva os dados numéricos com precisão completa
write.csv(
  resumo_janela_balanceada,
  file.path(DIR_SAIDA, "resumo_janela_balanceada_por_coorte_numerico.csv"),
  row.names = FALSE,
  fileEncoding = "UTF-8"
)

# CSV de apresentação, com valores monetários em R$ e duas casas decimais
exportar_csv_reais(
  resumo_janela_balanceada,
  colunas_reais = c(
    "remuneracao_mediana_reais_2024",
    "remuneracao_media_reais_2024"
  ),
  arquivo = file.path(
    DIR_SAIDA,
    "resumo_janela_balanceada_por_coorte.csv"
  )
)

# Compara o aumento das medianas entre os períodos da Tabela 5
medianas_tabela5 <- resumo_janela_balanceada %>%
  select(
    ano_conclusao,
    tempo_relativo_conclusao,
    remuneracao_mediana_reais_2024
  )

delta_tabela5 <- medianas_tabela5 %>%
  filter(tempo_relativo_conclusao == -1L) %>%
  transmute(
    ano_conclusao,
    mediana_t_menos1 = remuneracao_mediana_reais_2024
  ) %>%
  inner_join(
    medianas_tabela5 %>%
      filter(tempo_relativo_conclusao == 0L) %>%
      transmute(
        ano_conclusao,
        mediana_t_zero = remuneracao_mediana_reais_2024
      ),
    by = "ano_conclusao"
  ) %>%
  inner_join(
    medianas_tabela5 %>%
      filter(tempo_relativo_conclusao == 1L) %>%
      transmute(
        ano_conclusao,
        mediana_t_mais1 = remuneracao_mediana_reais_2024
      ),
    by = "ano_conclusao"
  ) %>%
  mutate(
    aumento_menos1_a_zero = mediana_t_zero - mediana_t_menos1,
    aumento_zero_a_mais1 = mediana_t_mais1 - mediana_t_zero,
    maior_aumento_menos1_a_zero =
      aumento_menos1_a_zero > aumento_zero_a_mais1
  )

n_coortes <- nrow(delta_tabela5)
n_coortes_maior_aumento <- sum(
  delta_tabela5$maior_aumento_menos1_a_zero,
  na.rm = TRUE
)

nota_tabela5 <- if (
  n_coortes > 0L &&
  n_coortes_maior_aumento > n_coortes / 2
) {
  sprintf(
    paste0(
      "Na Tabela 5, o maior aumento da mediana ocorre entre t = -1 e t = 0 ",
      "na maioria das coortes (%d de %d). Esse padrão sugere tendência de ",
      "carreira e não permite atribuir a variação ao efeito do diploma."
    ),
    n_coortes_maior_aumento,
    n_coortes
  )
} else if (n_coortes > 0L) {
  sprintf(
    paste0(
      "Na Tabela 5, o maior aumento da mediana entre t = -1 e t = 0 ocorre ",
      "em %d de %d coortes. A comparação é descritiva e não permite atribuir ",
      "a variação ao efeito do diploma."
    ),
    n_coortes_maior_aumento,
    n_coortes
  )
} else {
  paste(
    "Não foi possível comparar as medianas entre t = -1, t = 0 e t = 1",
    "nas coortes da Tabela 5."
  )
}

# Nota metodológica e limitações
nota <- c(
  "ANÁLISE DESCRITIVA DA TRAJETÓRIA SALARIAL PRÉ/PÓS-CONCLUSÃO",
  "",
  "Remuneração pessoa-ano: soma dos valores positivos de remun_media_nom dos vínculos da pessoa no mesmo ano.",
  "Valores monetários corrigidos para reais de 2024.",
  "Pré-conclusão: ano_referencia < ano_conclusao.",
  "Ano da conclusão: ano_referencia == ano_conclusao; separado pela ambiguidade temporal da RAIS anual.",
  "Pós-conclusão: ano_referencia > ano_conclusao.",
  "A comparação pareada exige ao menos um ano observado em cada período.",
  "A comparação é descritiva e não identifica efeito causal da formação.",
  "A jornada contratada não foi controlada; diferenças de remuneração também podem refletir jornadas de trabalho distintas.",
  nota_tabela5,
  "A ausência de registro salarial positivo não deve ser interpretada automaticamente como desemprego."
)

writeLines(
  nota,
  file.path(DIR_SAIDA, "nota_metodologica.txt"),
  useBytes = TRUE
)
