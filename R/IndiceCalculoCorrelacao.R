#' @title IndiceCalculoCorrelacao
#' @name IndiceCalculoCorrelacao
#'
#' @description
#' A função `IndiceCalculoCorrelacao` calcula correlações ponderadas entre variáveis especificadas e referência
#' em um conjunto de dados dividido por grupos. A função utiliza um método de ponderação para calcular correlações específicas,
#' organizando os resultados em arquivos temporários para otimizar a memória.
#'
#' @param splits Um `data.frame` com os detalhes dos grupos para os quais as correlações serão calculadas.
#' @param vou_rodar Um `character` que indica a pasta ou o grupo a ser processado.
#' @param dados Um `data.frame` contendo os dados brutos, incluindo as variáveis e pesos necessários para o cálculo.
#' @param indice_correlacao Um `data.frame` com as especificações das variáveis e referências de índice a serem correlacionadas.
#' @param caminho_temporario Um `character` que define o diretório temporário onde os resultados intermediários serão salvos, padrão é `"99.temp cal indi"`.
#' @param caminho_resultado Um `character` que define o diretório de saída onde o resultado final será salvo, padrão é `"Output"`.
#' @param nome_salvar Um `character` que define o nome do arquivo final de saída, padrão é `"Correlacao"`.
#' @param metodo_rodar Por padrão, calcula a correlação de 'kendall' (kendall, pearson, etc).

#' @details
#' A função realiza uma série de etapas:
#' - Filtragem e preparação dos dados de acordo com os parâmetros de `splits` e `indice_correlacao`.
#' - Cálculo das correlações entre variáveis e variáveis de referência, com ponderação opcional via método de Kendall.
#' - Salvamento de resultados temporários para cada grupo em arquivos `.rds`.
#' - Geração de um arquivo final com todos os resultados de correlação para os grupos processados.
#'
#' Internamente, a função `estima_CORR_peso` é usada para realizar o cálculo da correlação ponderada para as variáveis especificadas.
#'
#' @import dplyr
#' @import stringr
#' @importFrom purrr imap
#' @importFrom tidyr pivot_wider
#' @importFrom tidyr pivot_longer
#' @importFrom tibble tibble
#' @importFrom wdm indep_test
#'
#' @examples
#'
#' base::print("Sem Exemplo")
#'
#' @export

IndiceCalculoCorrelacao <- function(
    splits,
    vou_rodar,
    dados,
    indice_correlacao,
    caminho_temporario = "99.temp cal indi",
    caminho_resultado = "Output",
    nome_salvar = "Correlacao",
    metodo_rodar = "kendall"
)
{# Start: função 'IndiceCalculoCorrelacao'

  `%nin%` = base::Negate(`%in%`)

  all_indices = indice_correlacao %>%
    dplyr::filter(
      !base::is.na(indice_variavel) &
        indice_variavel %nin% c(""," ")
    ) %>%
    dplyr::select(indice_sigla) %>%
    dplyr::pull()

  {# Start: Criando para p/salvar arquivo temporário (não armazenar no MiB do R)

    # Criar o diretório temporário
    base::dir.create(caminho_temporario) %>%
      base::suppressWarnings()

    # Remover todos os arquivos e subdiretórios dentro do diretório temporário, para começar com um diretório limpo
    base::unlink(
      x = base::paste0(caminho_temporario, "/*"),
      recursive = TRUE
    )

    }# End: Criando para p/salvar arquivo temporário (não armazenar no MiB do R)

  # calculando
  x = splits %>%
    dplyr::filter(
      dono_id %in% c(
        splits %>%
          dplyr::filter(pasta_nome %in% c(vou_rodar)) %>%
          dplyr::select(split_id) %>%
          base::unique() %>%
          dplyr::pull()
      )
    ) %>%
    # dplyr::filter(dono_id == 3) %>%
    dplyr::arrange(dono_id) %>%
    dplyr::group_by(dono_id) %>%
    dplyr::group_split()

  x %>%
    purrr::imap(
      ~ {# Start: split_i

        # split_i = 1; split_df = x[[split_i]]
        split_i = .y
        split_df = .x

        {# Start: fazer para cada dono_id

          dono_id = split_df$dono_id %>%
            base::unique()

          pasta_nome = split_df$pasta_nome %>%
            base::unique()

          split_tipo_rodando = split_df$split_tipo %>%
            base::unique()

          filtro = split_df %>%
            dplyr::select(filtro_quest) %>%
            dplyr::pull() %>%
            base::unique() %>%
            stringr::str_replace_all("=", " == ") %>%
            stringr::str_replace_all("<>", "!=") %>%
            stringr::str_replace_all("like", " == ") %>%
            stringr::str_replace_all("LIKE", " == ") %>%
            stringr::str_replace_all(" and ", " & ") %>%
            stringr::str_replace_all(" AND ", " & ") %>%
            stringr::str_replace_all(" or ", " | ") %>%
            stringr::str_replace_all(" OR ", " | ") %>%
            stringr::str_replace_all(" not in [(]", " %nin% c(") %>%
            stringr::str_replace_all(" NOT IN [(]", " %nin% c(") %>%
            stringr::str_replace_all(" in [(]", " %in% c(") %>%
            stringr::str_replace_all(" IN [(]", " %in% c(") %>%
            stringr::str_replace_all(" in[(]", " %in% c(") %>%
            stringr::str_replace_all(" IN[(]", " %in% c(") %>%
            stringr::str_replace_all("===", "==") %>%
            stringr::str_replace_all("== ==", "==") %>%
            stringr::str_replace_all("< ==", "<=") %>%
            stringr::str_replace_all("> ==", ">=")

          split_tipo = split_df %>%
            dplyr::select(split_tipo) %>%
            dplyr::pull() %>%
            base::unique()

          split_nomes = split_df %>%
            dplyr::select(split_nome) %>%
            dplyr::pull() %>%
            base::unique()

          base::cat(base::paste0("\033[1;30m--------------------------------------------------------\033[0m\n"))
          base::cat(base::paste0("\033[1;32mdono_id ", dono_id, " = ", split_nomes, " [", split_i, "/", base::length(x), ']', "\033[0m\n"))

          {# Start: Calculando

            #O cálculo da região vai seguir a mesma lógica do resultado da distribuidora,
            #ou resultado abradee. Só dentro da aba subgrupos que não vai ser preciso ponderar,
            #lembrando que MT não vai considerar agente de relacionamento no isqp e AT vai considerar.
            #Todos os índices serão referencia em algum momento

            #Calculando cada um dos índices

            # ref = base::which(all_indices == "ISG")
            # i = base::which(all_indices == "SE")
            for( ref in 1:base::length(all_indices) )
            {# Start: Calculando cada um dos índices

              variaveis_referencia = indice_correlacao %>%
                dplyr::filter(indice_sigla == all_indices[ref]) %>%
                dplyr::select(indice_variavel) %>%
                dplyr::pull() %>%
                stringr::str_remove_all(" ") %>%
                stringr::str_remove_all("[()]") %>%
                stringr::str_split(",") %>%
                base::unlist()

              indice_sigla_referencia = indice_correlacao %>%
                dplyr::filter(indice_sigla == all_indices[ref]) %>%
                dplyr::select(indice_sigla) %>%
                dplyr::pull()

              indice_id_referencia = indice_correlacao %>%
                dplyr::filter(indice_sigla == all_indices[ref]) %>%
                dplyr::select(indice_id) %>%
                dplyr::pull()

              valor_ns_nr_referencia = indice_correlacao %>%
                dplyr::filter(indice_sigla == all_indices[ref]) %>%
                dplyr::select(valor_ns_nr) %>%
                dplyr::pull()

              for( i in 1:base::length(all_indices) )
              {#star fazer para cada indice

                {# Start: Informações

                  variaveis_rodando = indice_correlacao %>%
                    dplyr::filter(indice_sigla == all_indices[i]) %>%
                    dplyr::select(indice_variavel) %>%
                    dplyr::pull() %>%
                    stringr::str_remove_all(" ") %>%
                    stringr::str_remove_all("[()]") %>%
                    stringr::str_split(",") %>%
                    base::unlist()

                  indice_sigla_rodando = indice_correlacao %>%
                    dplyr::filter(indice_sigla == all_indices[i]) %>%
                    dplyr::select(indice_sigla) %>%
                    dplyr::pull()

                  indice_id_rodando = indice_correlacao %>%
                    dplyr::filter(indice_sigla == all_indices[i]) %>%
                    dplyr::select(indice_id) %>%
                    dplyr::pull()

                  valor_ns_nr_rodando = indice_correlacao %>%
                    dplyr::filter(indice_sigla == all_indices[i]) %>%
                    dplyr::select(valor_ns_nr) %>%
                    dplyr::pull()

                  vou_ponderar = FALSE

                }# End: Informações

                {# Start: Arrumando o banco de dados

                  dados_split = dados[dados$dono_id == dono_id, ] %>%
                    dplyr::filter(
                      base::eval(
                        base::parse(
                          text = filtro
                        )
                      )
                    )

                  variaveis = variaveis_rodando

                  banco_final = dados_split %>%
                    dplyr::select(dplyr::all_of(variaveis_referencia)) %>%
                    dplyr::rename_with(
                      ~ base::paste0("referencia_", .x),
                      .cols = dplyr::everything()
                    ) %>%
                    base::cbind(
                      dados_split %>%
                        dplyr::select(dplyr::all_of(variaveis)) %>%
                        dplyr::rename_with(
                          ~ base::paste0("variaveis_", .x),
                          .cols = dplyr::everything()
                        )
                    ) %>%
                    base::cbind(
                      dados_split %>%
                        dplyr::select(peso)
                    ) %>%
                    tidyr::pivot_longer(
                      cols = dplyr::starts_with("variaveis_"),
                      names_to = "name",
                      values_to = "value",
                      values_drop_na = TRUE
                    ) %>%
                    tidyr::pivot_longer(
                      cols = dplyr::starts_with("referencia_"),
                      names_to = "n_ref",
                      values_to = "referencia",
                      values_drop_na = TRUE
                    ) %>%
                    dplyr::mutate(
                      name = stringr::str_remove_all(name, "variaveis_"),
                      n_ref = stringr::str_remove_all(n_ref, "referencia_")
                    ) %>%
                    dplyr::mutate(
                      value = dplyr::case_when(
                        stats::as.formula(
                          base::paste0("value", valor_ns_nr_rodando, " ~ NA")
                        ),
                        TRUE ~ value
                      ),
                      referencia = dplyr::case_when(
                        stats::as.formula(
                          base::paste0("referencia", valor_ns_nr_referencia, " ~ NA")
                        ),
                        TRUE ~ referencia
                      )
                    )

                  # banco_final = dados_split %>% {
                  #   a = tibble::tibble(.)
                  #   a$var1 = a[[variaveis]]
                  #   a$var2 = a[[variaveis_referencia]]
                  #   a
                  # } %>%
                  #   dplyr::select(var1, var2, "peso") %>%
                  #   tidyr::pivot_longer(cols = var1) %>%
                  #   tidyr::pivot_longer(cols = var2, names_to="n_ref", values_to = "referencia") %>%
                  #   dplyr::mutate(
                  #     value_var = dplyr::case_when(
                  #       stats::as.formula(
                  #         base::paste0("value", valor_ns_nr_rodando, " ~ NA")
                  #       ),
                  #       TRUE ~ value
                  #     ),
                  #     referencia_var = dplyr::case_when(
                  #       stats::as.formula(
                  #         base::paste0("referencia", valor_ns_nr_referencia, " ~ NA")
                  #       ),
                  #       TRUE ~ referencia
                  #     )
                  #   )
                  #
                  # {# Start: Se for v112 = NPS, var_indice virá NPS e não índice
                  #
                  #   Como é correlação, não preciso fazer transformação!
                  #
                  # }# End: Se for v112 = NPS, var_indice virá NPS e não índice
                  #
                  # banco_final = banco_final %>%
                  #   dplyr::mutate(
                  #     referencia = referencia_var,
                  #     value = value_var
                  #   ) %>%
                  #   dplyr::select(-c(referencia_var, value_var))

                }# End: Arrumando o banco de dados

                {# Start: Calculando

                  resultado_indice = base::cbind(
                    dados_split %>%
                      dplyr::summarise(
                        indice_id_referencia = indice_id_referencia,
                        indice_sigla_referencia = indice_sigla_referencia,
                        linhas = dplyr::n(),
                        linhas_peso = base::sum(peso)
                      ),
                    #Correlação é sempre sem peso (ou seja, peso = 1)!!
                    correlacao = estima_CORR_peso(
                      dados = banco_final %>%
                        dplyr::mutate(peso = 1),
                      variaveis = base::list(
                        base::list(
                          "referencia" = "referencia",
                          "value" = "value"
                        )
                      ),
                      metodo = metodo_rodar
                    )$`Correlacao` #kendall
                  ) %>%
                    dplyr::mutate(
                      indice_sigla = indice_sigla_rodando,
                      indice_id = indice_id_rodando,
                      dono_id = dono_id,
                      split_nomes = split_nomes
                    ) %>%
                    dplyr::relocate(
                      c(dono_id, split_nomes, indice_id, indice_sigla),
                      c(1, 2, 3, 4, 5, 6, 7, 5)
                    )

                }# End: Calculando

                if( i == 1 ){

                  resultado_dono_id = resultado_indice

                } else {

                  resultado_dono_id = base::rbind(resultado_dono_id, resultado_indice)

                }

                base::rm(resultado_indice) %>%
                  base::suppressWarnings()

                # base::cat(base::paste0("\033[3;97mdono_id ", dono_id, " = ", split_nomes, " [", split_i, "/", base::length(x), '] | ', all_indices[ref], " [", ref, "/", base::length(all_indices), "] x ", all_indices[i], " [", i, "/", base::length(all_indices), "]\033[0m\n"))

              }# End: fazer para cada indice

              base::cat(
                base::paste0(
                  "\033[3;97mdono_id ", dono_id, " = ", split_nomes,
                  " [", split_i, "/", base::length(x), '] | Refer\u00EAncia: ',
                  all_indices[ref], " [", ref, "/", base::length(all_indices),
                  "]\033[0m\n"
                )
              )

              {# Start: salvando no arquivo temporário

                base::saveRDS(
                  object = resultado_dono_id,
                  file = base::paste0(caminho_temporario, "/", dono_id, "_", ref, ".rds")
                )

                IPpackage::print_cor(
                  df = base::paste0("Salvo em '", base::paste0(caminho_temporario, "/", dono_id, "_", ref, ".rds'"), "\n"),
                  cor = "azul",
                  negrito = FALSE,
                  data.frame = FALSE
                )

                base::rm(resultado_dono_id) %>%
                  base::suppressWarnings()

              }# End: salvando no arquivo temporário

              base::rm(resultado_dono_id) %>%
                base::suppressWarnings()

            }# End: Calculando cada um dos índices

          }# End: Calculando

        }# End: fazer para cada dono_id

      }# End: split_i

    )

  resultado_final = base::do.call(
    base::rbind,
    base::lapply(
      base::list.files(
        path = caminho_temporario,
        pattern = "\\.rds$",
        full.names = TRUE
      ),
      readRDS
    )
  )

  # Salvando o arquivo final, como os cálculos de todos os dono_id
  base::dir.create(caminho_resultado) %>%
    base::suppressWarnings()

  base::saveRDS(
    object = resultado_final,
    file = base::paste0(caminho_resultado, "/", nome_salvar, ".rds")
  )

  IPpackage::print_cor(
    df = "\n----------------------------------------------------------------\n",
    cor = "ciano",
    data.frame = FALSE
  )

  IPpackage::print_cor(
    df = base::paste0("Arquivos tempor\u00E1rios exclu\u00EDdos e o final salvo em\n '", base::paste0(caminho_resultado, "/", nome_salvar, ".rds'"), "\n"),
    cor = "verde",
    negrito = TRUE,
    data.frame = FALSE
  )

  base::unlink(
    x = base::paste0(caminho_temporario),
    recursive = TRUE
  )

}# End: função 'IndiceCalculoCorrelacao'
