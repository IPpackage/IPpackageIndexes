#' @title estima_CORR_peso
#'
#' @description
#' Esta função interna calcula a correlação ponderada entre pares de variáveis especificadas,
#' aplicando o teste de independência para cada par utilizando o método definido.
#'
#' @param dados Um data frame contendo as variáveis de interesse e a coluna de pesos.
#' @param variaveis Uma lista de pares de variáveis para as quais a correlação ponderada será calculada.
#'                  Cada elemento da lista é uma sublista com dois vetores, onde o primeiro vetor contém as variáveis para x,
#'                  e o segundo vetor contém as variáveis para y.
#' @param metodo Um string indicando o método de correlação a ser usado, sendo o padrão "pearson". Outros métodos suportados
#'               pela função wdm::indep_test incluem "spearman" e "kendall".
#'
#' @details
#' Esta função realiza o teste de independência para cada par de variáveis fornecidas, ponderando pelos valores da coluna `peso`.
#' As variáveis `x` e `y` são representadas pela média das variáveis contidas nos vetores de entrada.
#' A função aplica correlação ponderada para os pares de variáveis e retorna uma tabela com as estatísticas dos testes,
#' incluindo valores-p e uma representação da significância.
#'
#' @return
#' Retorna um `tibble` com as seguintes colunas:
#' \describe{
#'   \item{var1}{Nome do primeiro grupo de variáveis.}
#'   \item{var2}{Nome do segundo grupo de variáveis.}
#'   \item{Correlacao}{Valor da correlação estimada.}
#'   \item{statistic}{Estatística do teste de correlação.}
#'   \item{p_valor}{Valor-p do teste de correlação.}
#'   \item{sig_stars}{Indicação da significância do valor-p com símbolos "*".}
#'   \item{Metodo}{Método de correlação utilizado.}
#'   \item{hip_alternativa}{Hipótese alternativa do teste.}
#'   \item{var1_vars}{Lista de variáveis agrupadas no `x`.}
#'   \item{var2_vars}{Lista de variáveis agrupadas no `y`.}
#' }
#'
#' @keywords internal

estima_CORR_peso <- function(dados, variaveis, metodo = "pearson") {

  out <- vector(mode = "list", length = length(variaveis))

  for (i in seq_along(out)) {
    out[[i]] <- wdm::indep_test(
      x = dados %>%
        dplyr::select(dplyr::all_of(variaveis[[i]][[1]])) %>%
        dplyr::mutate(!!names(variaveis[[i]][1]) := rowMeans(dplyr::pick(dplyr::all_of(variaveis[[i]][[1]])), na.rm = T)) %>%
        dplyr::pull(!!names(variaveis[[i]][1])),
      y = dados %>%
        dplyr::select(dplyr::all_of(variaveis[[i]][[2]])) %>%
        dplyr::mutate(!!names(variaveis[[i]][2]) := rowMeans(dplyr::pick(dplyr::all_of(variaveis[[i]][[2]])), na.rm = T)) %>%
        dplyr::pull(!!names(variaveis[[i]][2])),
      weights = dados %>%
        dplyr::pull(peso),
      method = metodo,
      remove_missing = TRUE,
      alternative = "two-sided"
    ) %>%
      dplyr::mutate(
        "var1" = variaveis[[i]][1] %>% names(),
        "var2" = variaveis[[i]][2] %>% names()
      ) %>%
      dplyr::mutate(
        "var1_vars" = variaveis[[i]][[1]] %>% stringr::str_flatten_comma(),
        "var2_vars" = variaveis[[i]][[2]] %>% stringr::str_flatten_comma()
      )
  }

  out_df <- out %>%
    dplyr::bind_rows() %>%
    tibble::as_tibble() %>%
    dplyr::mutate("sig_stars" = dplyr::case_when(
      p_value <= 0.001  ~ "***",
      p_value <= 0.01   ~ "**",
      p_value <= 0.05   ~ "*",
      TRUE              ~ ""
    )) %>%
    dplyr::select(
      var1, var2 ,
      "Correlacao" = estimate, statistic, "p_valor" = p_value, sig_stars, "M\u00E9todo" = method, "hip_alternativa" = alternative,
      var1_vars, var2_vars
    )

  return(out_df)
}
