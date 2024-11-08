#' @title IndiceCalculoIndices
#' @name IndiceCalculoIndices
#'
#' @description
#' Esta função calcula índices de satisfação e importância com base em uma série de dados e parâmetros de referência.
#' Realiza uma análise detalhada de indicadores como médias ponderadas e métricas de NPS (Net Promoter Score),
#' aplicando ponderação de qualidade e preço conforme especificado.
#'
#' @param splits Um `data.frame` que contém os divisores e indicadores específicos para o cálculo dos índices.
#' @param vou_rodar Um `character` indicando o grupo ou divisão a ser processado, exemplo `c("DISTRIBUIDORA1", "DISTRIBUIDORA2")`.
#' @param referencia_indices Um `data.frame` de referência dos índices, que inclui informações sobre as variáveis
#' de cada índice, necessidade de ponderação e parâmetros adicionais.
#' @param dados Um `data.frame` com os dados a serem utilizados no cálculo dos índices.
#' @param amostra_pretendida Um `data.frame` opcional que especifica a amostra alvo para o cálculo dos índices.
#' @param all_indices Um `character vector` contendo os identificadores dos índices que serão calculados.
#' @param qualidade_preco Um `data.frame` com as ponderações relativas à qualidade e preço para os indicadores.
#' @param importancia_q5_pc Um `data.frame` com as importâncias relativas dos atributos de cada índice.
#' @param fonte_imp_entra_serie_hist Um `character` que define a fonte da importância a ser usada para a série histórica.
#' @param caminho_temporario Um `character` que especifica o diretório temporário onde os resultados intermediários
#' serão salvos, padrão `"99.temp cal indi"`.
#' @param caminho_resultado Um `character` que define o caminho onde o resultado final será salvo, padrão `"Output"`.
#' @param remover_dos_indices Um `character vector` indicando quais índices devem ser removidos da análise, padrão inclui vários identificadores como `"ISQP", "IEQP"`.
#' @param o_que_remover Um `character vector` especificando quais variáveis devem ser removidas dos índices durante a análise, padrão inclui variáveis como `"media", "n", "linhas"`.
#' @param nome_salvar Um `character` que define o nome do arquivo de saída, padrão `"Índice"`.
#'
#' @details
#' A função executa uma série de cálculos para avaliar o desempenho de índices de satisfação. As principais etapas incluem:
#' - Filtragem e pré-processamento dos dados, incluindo o ajuste de ponderações de qualidade e preço.
#' - Cálculo de médias ponderadas e índices específicos, considerando NPS e critérios de importância relativa.
#' - Geração de um arquivo temporário com os resultados intermediários para cada divisão ou indicador.
#' - Salvamento dos resultados finais no diretório de saída especificado.
#'
#' @import dplyr
#' @import stringr
#' @importFrom purrr imap
#' @importFrom tidyr pivot_wider
#' @importFrom tidyr pivot_longer
#' @importFrom tibble tibble
#'
#' @examples
#'
#' base::print("Sem Exemplo")
#'
#' @export



IndiceCalculoIndices <- function(
    splits,
    vou_rodar,
    referencia_indices,
    dados,
    amostra_pretendida = NULL,
    all_indices,
    qualidade_preco,
    importancia_q5_pc,
    fonte_imp_entra_serie_hist, # falar qual fonte inserir na série histórica
    caminho_temporario = "99.temp cal indi",
    caminho_resultado = "Output",
    remover_dos_indices = c("ISQP", "IEQP", "IIQP", "ISCP", "ISC", "IESC", "IIC", "ISCAL", "IECP", "IICP"),
    o_que_remover = c("media", "n", "n_peso", "linhas", "linhas_peso", "1", "2", "3",
                      "4", "5", "6", "7", "8", "9", "10", "11", "12",
                      "indiceescala_nota1", "indiceescala_nota2", "indiceescala_nota3",
                      "indiceescala_nota4", "indiceescala_nota5", 'indiceescala_nota6',
                      "indiceescala_nota7", "indiceescala_nota8", 'indiceescala_nota9',
                      "indiceescala_nota10", "indiceescala_nota11", 'indiceescala_nota12'
    ),
    nome_salvar = "Indice"
)
{# Start: calculo_indices_gc_abradee

  `%nin%` = base::Negate(`%in%`)

  all_indice_id = referencia_indices %>%
    dplyr::filter(IDAR %in% all_indices) %>%
    dplyr::select(indice_id) %>%
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
  splits %>%
    dplyr::filter(
      dono_id %in% c(
        splits %>%
          dplyr::filter(pasta_nome %in% c(vou_rodar)) %>%
          dplyr::select(split_id) %>%
          base::unique() %>%
          dplyr::pull()
      )
    ) %>%
    {# Start: Criando coluna com a quantidade de dono_id que estou rodando

      a = tibble::tibble(.)

      a = a %>%
        dplyr::mutate(
          n_dono_ids_rodando = a %>%
            dplyr::arrange(dono_id) %>%
            dplyr::group_by(dono_id) %>%
            dplyr::group_keys() %>%
            base::nrow()
        )

      a

    } %>%# End: Criando coluna com a quantidade de dono_id que estou rodando
    dplyr::arrange(dono_id) %>%
    dplyr::group_by(dono_id) %>%
    dplyr::group_split() %>%
    purrr::imap(

      ~ {# Start: para cada split_id

        # split_i = 2; split_df = xxxx[[split_i]]
        split_i = .y
        split_df = .x

        if( base::nrow(split_df) > 0 )
        {# Start: se tenho dados, rodar

          {# Start: fazer para cada dono_id

            {# Start: informações

              dono_id = split_df$dono_id %>%
                base::unique()

              distribuidora_id = split_df$distribuidora_id %>%
                base::unique()

              pasta_nome = split_df$pasta_nome %>%
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
                stringr::str_replace_all(" not in [(]", "%nin%c(") %>%
                stringr::str_replace_all(" NOT IN [(]", "%nin%c(") %>%
                stringr::str_replace_all(" in [(]", "%in%c(") %>%
                stringr::str_replace_all(" IN [(]", "%in%c(") %>%
                stringr::str_replace_all(" in[(]", "%in%c(") %>%
                stringr::str_replace_all(" IN[(]", "%in%c(") %>%
                stringr::str_replace_all("===", "==") %>%
                stringr::str_replace_all("==  ==", "==") %>%
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

              IPpackage::print_cor(
                df = "----------------------------------------------------------------\n",
                cor = "ciano",
                data.frame = FALSE
              )

              IPpackage::print_cor(
                df = base::paste0(
                  "dono_id ", dono_id, " = ", split_nomes, " [", split_i, "/", base::unique(split_df$n_dono_ids_rodando), "]\n"
                ),
                cor = "verde",
                negrito = TRUE,
                data.frame = FALSE
              )

            }# End: informações

            # i = base::which(all_indices == "ISQP")
            for( i in 1:base::length(all_indices) )
            {# Start: fazer para cada indice

              {# Start: informações do índice

                filtro_aplicado = filtro

                indice_rodando = referencia_indices %>%
                  dplyr::filter(IDAR == all_indices[i])

                vou_ponderar = indice_rodando$ponderar

                # Banco de dados filtrado para o dono_id
                dados_split = dados[dados$dono_id == dono_id, ] %>%
                  dplyr::filter(
                    base::eval(
                      base::parse(
                        text = filtro_aplicado
                      )
                    )
                  )

                amostra = base::nrow(dados_split)

                variaveis_variaveis_geral = indice_rodando %>%
                  dplyr::select("variaveis_geral") %>%
                  dplyr::pull()

                variaveis = base::unlist(
                  base::strsplit(
                    indice_rodando %>%
                      dplyr::select("variaveis_geral") %>%
                      dplyr::pull() %>%
                      stringr::str_remove_all("[()]") %>%
                      stringr::str_remove_all(" "),
                    ","
                  )
                )

                if( vou_ponderar == "imp_preco_qualidade" )
                {# Start: pegando as variáveis que precisarei rodar

                  indices_rodar = variaveis

                  variaveis = base::unlist(
                    base::strsplit(
                      referencia_indices %>%
                        dplyr::filter(IDAR %in% indices_rodar) %>%
                        dplyr::select("variaveis_geral") %>%
                        dplyr::pull(),
                      ","
                    )
                  )

                } else {

                  indices_rodar = NA

                }# End: pegando as variáveis que precisarei rodar

                valor_indice = indice_rodando %>%
                  dplyr::select(valor_indice) %>%
                  dplyr::pull()

                valor_ns_nr = indice_rodando %>%
                  dplyr::select(valor_ns_nr) %>%
                  dplyr::pull()

                rodar_NPS = indice_rodando %>%
                  dplyr::select(var_nps) %>%
                  dplyr::pull()

                # Falo se farei mrg tradicional ou se é média das médias
                mrg_media = base::ifelse(
                  base::is.na(indice_rodando$mrg_media),
                  FALSE,
                  indice_rodando$mrg_media
                )

                # fonte da importância
                fonte_importancia = indice_rodando$fonte_imp

              }# End: informações do índice

              {# Start: aplicando transformar_var caso necessário

                all_transformacoes = indice_rodando %>%
                  dplyr::filter(transformar_var != FALSE & !base::is.na(transformar_var) & transformar_var != TRUE) %>%
                  dplyr::select(transformar_var) %>%
                  dplyr::pull()

                if( base::length(all_transformacoes) >= 1 )
                {# Start: se preciso aplicar alguma fórmula

                  for ( w in 1:base::length(all_transformacoes) )
                    # for ( w in 1 )
                  {# Start: aplicando cada um dos transformar_var

                    df = dados_split

                    formula_text <- all_transformacoes[w] %>%
                      stringr::str_remove_all("\n")

                    # Dividir cada uma das fórmulas
                    formulas <- stringr::str_split(
                      string = formula_text %>%
                        #para manter o ")" de fechamento do dplyr
                        stringr::str_replace_all("[)],", ")),"),
                      pattern = "[)],"
                    )[[1]]

                    for( ww in 1:base::length(formulas) )
                    {# Start: aplicando todas as fórmulas

                      fml = formulas[ww]

                      IPpackage::print_cor(
                        df = base::paste0("Transforma\u00E7\u00E3o aplicada:\n", fml, "\n"),
                        cor = "vermelho",
                        data.frame = FALSE
                      )

                      # Pegando a variável a ser alterada (tudo antes do primeiro "=")
                      var_alterar = base::sub("^(.*?)=.*", "\\1", fml) %>%
                        stringr::str_remove_all(" ")

                      # Removendo tudo que está antes do primeiro "="
                      fml = base::sub(".*?=\\s*", "", fml)

                      df$var_alterar_qwerasd = df[[var_alterar]]

                      df = df %>%
                        dplyr::mutate(
                          var_alterar_qwerasd = base::eval(
                            base::parse(
                              text = fml
                            )
                          )
                        ) %>%
                        dplyr::select(-dplyr::all_of(var_alterar))

                      # Voltando para o nome original da variável
                      base::colnames(df)[base::which(base::colnames(df) == "var_alterar_qwerasd")] = var_alterar

                    }# End: aplicando todas as fórmulas

                    dados_split = df

                    base::rm(df, formula_text, formulas, ww, var_alterar, fml) %>%
                      base::suppressWarnings()

                  }# End: aplicando cada um dos transformar_var

                }# End: se preciso aplicar alguma fórmula

                base::rm(all_transformacoes, w, ww) %>%
                  base::suppressWarnings()

              }# End: aplicando transformar_var caso necessário

              {# Start: arrumando o banco de dados

                if( indice_rodando$ponderar %nin% "porte_at_mt" )
                {# Start: se não vou poderar por porte, não preciso dessa informação vinda no banco. Coloco um genérico aqui

                  dados_split$porte = "GERAL"

                }# End: se não vou poderar por porte, não preciso dessa informação vinda no banco. Coloco um genérico aqui

                banco_final = dados_split %>%
                  dplyr::select(dplyr::all_of(variaveis), "peso", porte) %>%
                  tidyr::pivot_longer(cols = dplyr::all_of(variaveis)) %>%
                  dplyr::mutate(name_orig = name) %>%
                  dplyr::relocate(name, .after = name_orig) %>%
                  {# Start: removendo NA's em 'value'

                    banco_final = tibble::tibble(.)

                    if( !all(base::is.na(banco_final$value)) )
                    {# Start: se tenho NA, remover

                      banco_final = banco_final %>%
                        dplyr::filter(!base::is.na(value))

                    }# End: se tenho NA, remover

                    banco_final

                  } %>% # End: removendo NA's em 'value'
                  dplyr::mutate(
                    #Média
                    var_media = dplyr::case_when(
                      stats::as.formula(
                        base::paste0("value", valor_ns_nr, " ~ NA")
                      ),
                      TRUE ~ value
                    ),
                    #índice
                    var_indice = dplyr::case_when(
                      stats::as.formula(
                        base::paste0("value", valor_ns_nr, " ~ NA")
                      ),
                      stats::as.formula(
                        base::paste0("value", valor_indice, " ~ 100")
                      ),
                      TRUE ~ 0
                    )
                  ) %>%
                  {# Start: se o banco tem 0 linhas, preencher com 0

                    # Isso é por causa do erro:
                    # Error in data.frame(..., check.names = FALSE) :
                    #   argumentos implicam em número de linhas distintos: 1, 0

                    a = tibble::tibble(.)

                    if( base::nrow(a) == 0 )
                    {# Start: se não tem dados, colocar como zero

                      b = tibble::tibble(
                        base::rep(0, base::ncol(a))
                      ) %>%
                        base::t() %>%
                        data.frame() %>%
                        tibble::tibble()

                      base::colnames(b) = base::colnames(a)
                      base::rownames(b) = NULL

                      a = b %>%
                        tibble::tibble() %>%
                        dplyr::mutate(
                          name = "vxnaoexiste",
                          porte = "vxnaoexiste"
                        )

                    }# End: se não tem dados, colocar como zero

                    a

                  }

                if( rodar_NPS == TRUE )
                {# Start: Se for rodar_NPS = TRUE = NPS, var_indice virá NPS e não índice

                  banco_final = banco_final %>%
                    dplyr::mutate(
                      var_indice = dplyr::case_when(
                        stats::as.formula(
                          base::paste0("value", valor_ns_nr, " ~ NA")
                        ),
                        value %in% c(0:6) ~ -100,
                        value %in% c(7:8) ~ 0,
                        value %in% c(9:10) ~ 100
                      )
                    )

                }# End: Se for rodar_NPS = TRUE = NPS, var_indice virá NPS e não índice

                if( indice_rodando$ponderar == "imp_relativo" )
                {# Start: ajustes na tabela de importância e no banco quando for "imp_relativo"

                  imp_variaveis = indice %>%
                    dplyr::filter(indice_variavel %in% variaveis) %>%
                    dplyr::select(indice_id, indice_sigla, indice_variavel) %>%
                    dplyr::left_join(
                      importancia_q5_pc %>%
                        dplyr::filter(dono_id == distribuidora_id & Fonte == fonte_importancia) %>%
                        dplyr::select(indice_id, importanciaq5_relativa),
                      by = c("indice_id" = "indice_id")
                    ) %>%
                    dplyr::filter(!base::is.na(importanciaq5_relativa)) %>%
                    {# Start: se o banco tem 0 linhas, preencher com 0

                      # Isso é por causa do erro:
                      # Error in data.frame(..., check.names = FALSE) :
                      #   argumentos implicam em número de linhas distintos: 1, 0

                      a = tibble::tibble(.)

                      if( base::nrow(a) == 0 )
                      {# Start: se não tem dados, colocar como zero

                        b = tibble::tibble(
                          base::rep(0, base::ncol(a))
                        ) %>%
                          base::t() %>%
                          data.frame() %>%
                          tibble::tibble()

                        base::colnames(b) = base::colnames(a)
                        base::rownames(b) = NULL

                        a = b %>%
                          tibble::tibble() %>%
                          dplyr::mutate(
                            indice_sigla = "vxnaoexiste",
                            indice_variavel = "vxnaoexiste"
                          )

                      }# End: se não tem dados, colocar como zero

                      a

                    }

                  {# Start: verificando se tenho MRG a ser calculado (ao invés de índice simples)

                    variaveis_original = variaveis_variaveis_geral %>%
                      stringr::str_remove_all(" ")

                    if( base::any(stringr::str_detect(variaveis_original, "[(]")) )
                    {# Start: encontrei (), indicativo de grupos de variáveis para MRG

                      # Caso de MRG padrão + ponderar = imp_relativo

                      # Quando isso acontecer, as variáveis vão virar a que tiver importância
                      # Assim, fica um mrg e consigo aplicar a importância após calcular o IDAR.
                      # Posteriormente, defini que irá utilizar a importância da primeira variável
                      # citada. ou seja, se for (v2, v1), utilizar a importância de v2

                      variaveis_original = base::unlist(
                        base::strsplit(
                          variaveis_original,
                          "),"
                        )
                      ) %>%
                        stringr::str_remove_all("[()]") %>%
                        base::as.list()

                      for( ww in 1:base::length(variaveis_original) )
                      {# Start: olhando cada grupo de variáveis

                        lista_vars = base::strsplit(
                          x = variaveis_original[ww][[1]],
                          split = ","
                        )[[1]]

                        # Uso a importância da primeira variável citada
                        variavel_vou_usar = lista_vars[1]

                        indice_var_usar = referencia_indices %>%
                          dplyr::filter(variaveis_geral == variaveis_original[ww][[1]]) %>%
                          dplyr::select(indice_id) %>%
                          dplyr::pull()

                        nome_oficial_var = imp_variaveis %>%
                          dplyr::filter(
                            indice_variavel %in% variavel_vou_usar
                          ) %>%
                          dplyr::filter(!base::is.na(importanciaq5_relativa)) %>%
                          dplyr::select(indice_variavel) %>%
                          base::unique() %>%
                          dplyr::pull()

                        if( base::all(imp_variaveis$indice_sigla == "vxnaoexiste") )
                        {# Start: se não tive dados

                          nome_oficial_var = "vxnaoexiste"

                        }# End: se não tive dados

                        {# Start: verificando se tenho mais de um 'variavel_vou_usar' na imp_variaveis

                          #  Por exemplo: se CL = v143 e CL_S = v143. Então imp_variaveis terá dois índices para a mesma var

                          a = imp_variaveis %>%
                            dplyr::filter(indice_variavel == variavel_vou_usar)

                          if( base::nrow(a) >= 2 )
                          {# Start: se tenho alguma indice_variavel que aparece mais de uma vez

                            if( base::length(base::unique( base::round(a$importanciaq5_relativa, 5) )) > 1 )
                            {# Start: se mais de 1 índice para o mesmo 'indice_variavel', erro. Se não, continuar

                              IPpackage::print_cor(
                                df = a,
                                cor = "vermelho"
                              )

                              base::stop("Mais de uma 'importanciaq5_relativa' para o mesmo 'indice_variavel' ")

                            } else {

                              imp_variaveis =  imp_variaveis %>%
                                dplyr::filter(indice_variavel %nin% variavel_vou_usar) %>%
                                dplyr::bind_rows(
                                  a %>%
                                    dplyr::filter(indice_id == indice_var_usar)
                                )

                            }# End: se mais de 1 índice para o mesmo 'indice_variavel', erro. Se não, continuar

                          }# End: se tenho alguma indice_variavel que aparece mais de uma vez

                        }# End: verificando se tenho mais de um 'variavel_vou_usar' na imp_variaveis

                        banco_final = banco_final %>%
                          dplyr::mutate(
                            name = dplyr::case_when(
                              name %in% base::unlist(
                                base::strsplit(
                                  x = variaveis_original[ww][[1]],
                                  split = ","
                                )
                              ) ~ nome_oficial_var,
                              TRUE ~ name
                            )
                          )

                      }# End: olhando cada grupo de variáveis

                      base::rm(variaveis_original, ww, nome_oficial_var) %>%
                        base::suppressWarnings()

                    }# End: encontrei (), indicativo de grupos de variáveis para MRG

                  }# End: verificando se tenho MRG a ser calculado (ao invés de índice simples)

                  {# Start: verificando se tenho mais de um 'variavel_vou_usar' na imp_variaveis

                    # Tenho uma etapa similar dentro do '# Start: olhando cada grupo de variáveis'.
                    # mas lá só faz a seleção do indice_id que irei usar no caso de ser um MRG
                    # exemplo, (v1, v2). Lá ele garante que irei usar a importância de v1 apenas
                    # (importância da primeira variável citada no MRG).
                    # Aqui é para geral, olho todo mundo. Até se tiver v1 = C1 e v1 = C1_S.
                    # Tenho que ficar com apenas um!!

                    imp_variaveis = imp_variaveis %>%
                      dplyr::group_by(indice_variavel) %>%
                      dplyr::group_split() %>%
                      purrr::imap(

                        ~{# Start: analisando quando tenho indice_variavel duplicado

                          # ai = 12; a = aa[[ai]]
                          ai = .y
                          a = .x

                          if( base::nrow(a) >= 2 )
                          {# Start: se tenho alguma indice_variavel que aparece mais de uma vez

                            # Se tenho essa var na indice_rodando, quer dizer que é o índice que quero
                            se_e_o_indice_calculado = base::any(
                              indice_rodando %>%
                                dplyr::filter(indice_id == all_indice_id[i]) %>%
                                dplyr::select(variaveis_geral) %>%
                                dplyr::pull() == a$indice_sigla
                            )


                            if( se_e_o_indice_calculado == TRUE )
                            {# Start: se for o índice que vou rodar

                              a = a %>%
                                dplyr::filter(indice_id == all_indice_id[i])

                              if( base::length(base::nrow(a) > 1) )
                              {# Start: se mais de 1 índice para o mesmo 'indice_id', erro. Se não, continuar

                                IPpackage::print_cor(
                                  df = a,
                                  cor = "vermelho"
                                )

                                base::stop("Mais de uma 'importanciaq5_relativa' para o mesmo 'indice_id' ")

                              }

                            } else {# End: se for o índice que vou rodar

                              if( base::length(base::unique( base::round(a$importanciaq5_relativa, 5) )) > 1 )
                              {# Start: se mais de 1 índice para o mesmo 'indice_variavel', erro. Se não, continuar

                                IPpackage::print_cor(
                                  df = a,
                                  cor = "vermelho"
                                )

                                base::stop("Mais de uma 'importanciaq5_relativa' para o mesmo 'indice_variavel' ")

                              } else {

                                a =  a[1, ]

                              }# End: se mais de 1 índice para o mesmo 'indice_variavel', erro. Se não, continuar

                            }# End: selecionando apenas um

                          }# End: se tenho alguma indice_variavel que aparece mais de uma vez

                          a

                        }# End: analisando quando tenho indice_variavel duplicado

                      ) %>%
                      dplyr::bind_rows()

                  }# End: verificando se tenho mais de um 'variavel_vou_usar' na imp_variaveis

                }# End: ajustes na tabela de importância e no banco quando for "imp_relativo"

                {# Start: amostra pretendida

                  if( base::is.null(amostra_pretendida) )
                  {# Start: amostra pretendida (se tiver)

                    amostra_prent = NA

                  }else{

                    amostra_prent = amostra_pretendida[amostra_pretendida$dono_id == dono_id, ] %>%
                      dplyr::filter(!base::is.na(dono_id)) %>%
                      dplyr::select(amostra_pretendida) %>%
                      base::unique() %>%
                      dplyr::pull()

                  }# End: amostra pretendida (se tiver)

                  if( base::length(amostra_prent) == 0 )
                  {# Start: se tenho nada em amostra_prent, NA

                    amostra_prent = NA

                  }# End: se tenho nada em amostra_prent, NA

                }# End: amostra pretendida

              }# End: arrumando o banco de dados

              {# Start: Calculando

                {# Start: calculando o índice

                  if( indice_rodando$ponderar == FALSE )
                  {# Start = Se não faço calculos com importância

                    if( mrg_media == TRUE )
                    {# Start: se o mrg é média das médias ou se é empilhado [é média das médias]

                      nova_formula = variaveis_variaveis_geral

                      all_vars = base::unlist(
                        base::strsplit(
                          base::gsub("[()]", "", nova_formula),
                          ","
                        )
                      )

                      # se tenho variável isolada dentro de (), remover o ()
                      # Exemplo, "(v1),(v2,v3),v4,(v5)" deve virar "v1,(v2,v3),v4,v5"
                      nova_formula = stringr::str_replace_all(nova_formula, "\\(([^,]+)\\)", "\\1")

                      if( !stringr::str_detect(nova_formula, ",") )
                      {# Start: se falou que é um MRG média das médias, preciso ter "," em variaveis_geral. Se não tiver, erro!

                        base::stop("Nenhuma ','detectada:\nSe \u00E9 mrg do tipo m\u00E9dia das m\u00E9dias, precisa ter mais de uma vari\u00E1vel em 'variaveis_geral'")

                      }# End: se falou que é um MRG média das médias, preciso ter "," em variaveis_geral. Se não tiver, erro!

                      # if( !str_detect(nova_formula, "^\\(") )
                      # {# Start: se não começa com "(", colocar no começo e no fim
                      #
                      #   nova_formula = base::paste0("(", nova_formula, ")")
                      #
                      # }# End: se não começa com "(", colocar no começo e no fim

                      nova_formula = base::paste0("(", nova_formula, ")")


                      nova_formula = nova_formula %>%
                        stringr::str_replace_all("[(]", "base::mean(c(") %>%
                        stringr::str_replace_all("[)]", "), na.rm = TRUE)")

                      # Calcula o índice e a média ponderada de cada coluna
                      medias_calculadas = dplyr::left_join(
                        banco_final %>%
                          dplyr::group_by(name_orig) %>%
                          dplyr::summarise(
                            dplyr::across(
                              .cols = var_indice,
                              .fns = ~ stats::weighted.mean(.x, w = peso, na.rm = TRUE),
                              .names = "var_indice"
                            )
                          ),
                        banco_final %>%
                          dplyr::group_by(name_orig) %>%
                          dplyr::summarise(
                            dplyr::across(
                              .cols = var_media,
                              .fns = ~ stats::weighted.mean(.x, w = peso, na.rm = TRUE),
                              .names = "var_media"
                            )
                          ),
                        by = join_by("name_orig")
                      )

                      if( base::all(base::is.na(medias_calculadas$var_indice)) )
                      {# Start: se tudo resultou em NA

                        medias_calculadas = tibble::tibble(
                          name_orig = all_vars,
                          var_indice = base::rep(NA,base::length(all_vars)),
                          var_media = base::rep(NA,base::length(all_vars))
                        )

                      }# End: se tudo resultou em NA

                      # left_join é para garantir que tenha todas as vars no banco calculado
                      medias_calculadas = dplyr::left_join(
                        tibble::tibble(
                          name_orig = all_vars
                        ),
                        medias_calculadas,
                        by = join_by("name_orig")
                      )

                      nova_formula_i = nova_formula
                      nova_formula_m = nova_formula

                      for( w in 1:base::length(all_vars))
                      {# Start: substituindo a variável pela da média

                        # índice
                        novo_valor_i = medias_calculadas %>%
                          dplyr::filter(name_orig == all_vars[w]) %>%
                          dplyr::select(var_indice) %>%
                          dplyr::pull()

                        novo_valor_i = base::ifelse(base::is.na(novo_valor_i), "NA", base::as.character(novo_valor_i))

                        nova_formula_i = nova_formula_i %>%
                          stringr::str_replace_all(
                            pattern = all_vars[w],
                            replacement = novo_valor_i
                          )

                        #Média
                        novo_valor_m = medias_calculadas %>%
                          dplyr::filter(name_orig == all_vars[w]) %>%
                          dplyr::select(var_media) %>%
                          dplyr::pull()

                        novo_valor_m = base::ifelse(base::is.na(novo_valor_m), "NA", base::as.character(novo_valor_m))

                        nova_formula_m = nova_formula_m %>%
                          stringr::str_replace_all(
                            pattern = all_vars[w],
                            replacement = novo_valor_m
                          )

                        base::rm(novo_valor_i, novo_valor_m)

                      }# End: substituindo a variável pela sa média

                      resultados_calculados =
                        tibble::tibble(
                          indice = base::eval(
                            base::parse(text = nova_formula_i)
                          ),
                          media = base::eval(
                            base::parse(text = nova_formula_m)
                          ),
                          n = 0,
                          n_peso = 0
                        )

                    } else {# End: [é média das médias] | Start: [mrg empilhado]

                      resultados_calculados =
                        banco_final %>%
                        dplyr::filter(!base::is.na(value)) %>%
                        dplyr::summarise(
                          indice = weighted.mean(var_indice, w = peso, na.rm = TRUE),
                          media = weighted.mean(var_media, w = peso, na.rm = TRUE),
                          n = n(),
                          n_peso = base::sum(peso)
                        )

                    }# End: se o mrg é média das médias ou se é empilhado

                  }# End = Se não faço calculos com importância [mrg empilhado]

                  if( indice_rodando$ponderar == "imp_relativo" )
                  {# Start = precisa calcular a soma ( Taxa de satisfação = IDAT * importância Relativa )

                    IDATs = banco_final %>%
                      dplyr::group_by(name) %>%
                      dplyr::filter(!base::is.na(value)) %>%
                      dplyr::summarise(
                        indice = weighted.mean(var_indice, w = peso, na.rm = TRUE)
                        , media = weighted.mean(var_media, w = peso, na.rm = TRUE)
                        , n = n()
                        , n_peso = sum(peso)
                      )

                    resultados_calculados =
                      IDATs %>%
                      dplyr::left_join(
                        imp_variaveis %>%
                          dplyr::select(indice_id, indice_sigla, indice_variavel, importanciaq5_relativa),
                        by = c("name" = "indice_variavel")
                      ) %>%
                      dplyr::mutate(taxa = indice * importanciaq5_relativa / 100) %>%
                      dplyr::summarise(
                        indice = base::sum(taxa, na.rm = TRUE),
                        media = NA,
                        n = NA,
                        n_peso = NA
                      )

                    base::rm(imp_variaveis, IDATs)

                  }# End = precisa calcular a soma ( Taxa de satisfação = IDAT * importância Relativa )

                  if( indice_rodando$ponderar == "imp_preco_qualidade" )
                  {# Start = precisa calcular (ISQP * Imp qualidade) + (ISCP * Imp Preço)

                    for( xi in 1:base::length(indices_rodar))
                    {# Start: calcular os índices que preciso (ISQP e ISCP)

                      variaveis_xi = base::unlist(
                        base::strsplit(
                          referencia_indices %>%
                            dplyr::filter(IDAR %in% indices_rodar[xi]) %>%
                            dplyr::select("variaveis_geral") %>%
                            dplyr::pull(),
                          ","
                        )
                      )

                      banco_final_xi = banco_final %>%
                        dplyr::filter(name %in% variaveis_xi)

                      imp_variaveis_xi = indice %>%
                        dplyr::filter(indice_variavel %in% variaveis_xi) %>%
                        dplyr::select(indice_id, indice_sigla, indice_variavel) %>%
                        dplyr::left_join(
                          importancia_q5_pc %>%
                            dplyr::filter(dono_id == distribuidora_id) %>%
                            dplyr::select(indice_id, importanciaq5_relativa),
                          by = c("indice_id" = "indice_id")
                        )

                      if ( base::all(base::is.na(imp_variaveis_xi$importanciaq5_relativa)) )
                      {# Start: se não tenho importância e se tenho

                        resultados_calculados_xi = banco_final_xi %>%
                          dplyr::filter(!base::is.na(value)) %>%
                          dplyr::summarise(
                            indice = weighted.mean(var_indice, w = peso, na.rm = TRUE)
                            , media = weighted.mean(var_media, w = peso, na.rm = TRUE)
                            , n = n()
                            , n_peso = sum(peso)
                          ) %>%
                          dplyr::mutate(rodou = indices_rodar[xi])

                      } else {

                        IDATs_xi = banco_final_xi %>%
                          dplyr::group_by(name) %>%
                          dplyr::filter(!base::is.na(value)) %>%
                          dplyr::summarise(
                            indice = weighted.mean(var_indice, w = peso, na.rm = TRUE)
                            , media = weighted.mean(var_media, w = peso, na.rm = TRUE)
                            , n = n()
                            , n_peso = sum(peso)
                          )

                        resultados_calculados_xi =
                          IDATs_xi %>%
                          dplyr::left_join(
                            imp_variaveis_xi %>%
                              dplyr::select(indice_id, indice_sigla, indice_variavel, importanciaq5_relativa),
                            by = c("name" = "indice_variavel")
                          ) %>%
                          dplyr::mutate(taxa = indice * importanciaq5_relativa / 100) %>%
                          dplyr::summarise(
                            indice = base::sum(taxa, na.rm = TRUE),
                            media = NA,
                            n = NA,
                            n_peso = NA
                          ) %>%
                          dplyr::mutate(rodou = indices_rodar[xi])

                        base::rm(imp_variaveis_xi, IDATs_xi) %>%
                          base::suppressWarnings()

                      }# End: se não tenho importância e se tenho

                      if(xi == 1)
                      {# Start: armazenando o resultado xi

                        resultado_xi = resultados_calculados_xi

                      } else {

                        resultado_xi = base::rbind(resultado_xi, resultados_calculados_xi)

                      }# End: armazenando o resultado xi

                    }# End: calcular os índices que preciso (ISQP e ISCP)

                    resultados_calculados = resultado_xi %>%
                      dplyr::left_join(
                        # Nesse caso, usar qualidade_preço da distribuidora (distribuidora_id) e não da região/subgrupo (dono_id) etc e etc
                        qualidade_preco[qualidade_preco$dono_id == distribuidora_id, ] %>%
                          pivot_longer(cols = c(2, 3), names_to = "indice", values_to = "imp") %>%
                          dplyr::select(indice, imp),
                        by = c("rodou" = "indice")
                      ) %>%
                      dplyr::mutate(
                        multiplicacao = indice * imp
                      ) %>%
                      dplyr::summarise(
                        indice = base::sum(multiplicacao),
                        media = NA,
                        n = NA,
                        n_peso = NA
                      )

                    base::rm(xi, imp_variaveis_xi, variaveis_xi) %>%
                      base::suppressWarnings()

                  }# End = precisa calcular a soma ( Taxa de satisfação = IDAT * importância Relativa )

                  if( indice_rodando$ponderar == "porte_at_mt" )
                  {# Start = ponderar pelo porte (indice_at * n_at + indice_mt * n_mt) / (n_at + n_mt)

                    resultados_calculados_porte = banco_final %>%
                      dplyr::arrange(porte) %>%
                      dplyr::group_by(porte) %>%
                      dplyr::group_split()  %>%
                      purrr::imap(
                        ~ {

                          #porte_i = 2; porte_df = dados_split_porte[[porte_i]]
                          porte_i = .y
                          porte_df = .x
                          #
                          {# Start: Informações

                            porte_rodando = base::unique(porte_df$porte)

                            variaveis_porte_rodando = base::unlist(
                              base::strsplit(
                                indice_rodando %>%
                                  dplyr::select(
                                    base::paste0(
                                      "variaveis_",
                                      porte_rodando %>%
                                        stringr::str_to_lower()
                                    )
                                  ) %>%
                                  dplyr::pull() %>%
                                  stringr::str_remove_all(" ") %>%
                                  stringr::str_remove_all("[()]"),
                                ","
                              )
                            )

                            n_linhas_porte = dados_split %>%
                              dplyr::filter(porte == porte_rodando) %>%
                              dplyr::select(dplyr::all_of(variaveis_porte_rodando), "peso") %>%
                              {# Start: removendo NA's

                                temp = tibble::tibble(.)

                                xtemp = temp %>%
                                  dplyr::select(-c(peso)) %>%
                                  dplyr::mutate(
                                    dplyr::across(
                                      .cols = dplyr::everything(),
                                      ~ base::ifelse(base::is.na(.), 1, 0)
                                    )
                                  ) %>%
                                  dplyr::mutate(x = base::rowSums(.)) %>%
                                  dplyr::select(x) %>%
                                  dplyr::pull()

                                temp$qwe = xtemp

                                temp = temp %>%
                                  dplyr::filter((qwe < (base::length(variaveis) + 1))) %>%
                                  dplyr::select(-qwe)

                                rm(xtemp)

                                temp
                              } %>% # End: removendo NA's
                              dplyr::summarise(
                                linhas = n(),
                                linhas_peso = base::sum(peso)
                              )

                          }# End: Informações

                          {# Start: calculando para esse porte

                            porte_df %>%
                              dplyr::filter(name_orig %in% c(variaveis_porte_rodando)) %>%
                              dplyr::filter(!base::is.na(value)) %>%
                              dplyr::summarise(
                                indice = stats::weighted.mean(var_indice, w = peso, na.rm = TRUE),
                                media = stats::weighted.mean(var_media, w = peso, na.rm = TRUE),
                                n = n(),
                                n_peso = base::sum(peso)
                              ) %>%
                              dplyr::mutate(
                                porte_rodando = porte_rodando,
                                linhas = n_linhas_porte$linhas,
                                linhas_peso = n_linhas_porte$linhas_peso
                              )

                          }# End: calculando para esse porte

                        }
                      ) %>%
                      dplyr::bind_rows()

                    resultados_calculados = resultados_calculados_porte %>%
                      {# Start: Ponderando

                        #Verificando se preciso ponderar
                        resultado_indice = tibble::tibble(.)

                        if( base::nrow(resultado_indice) > 1 )
                        {# Start: preciso ponderar

                          n_at = resultado_indice %>%
                            dplyr::filter(porte_rodando == 'AT') %>%
                            dplyr::select(linhas) %>%
                            dplyr::pull() %>%
                            base::as.numeric()

                          n_mt = resultado_indice %>%
                            dplyr::filter(porte_rodando == 'MT') %>%
                            dplyr::select(linhas) %>%
                            dplyr::pull() %>%
                            base::as.numeric()

                          indice_at = resultado_indice %>%
                            dplyr::filter(porte_rodando == 'AT') %>%
                            dplyr::select(indice) %>%
                            dplyr::pull() %>%
                            base::as.numeric()

                          indice_mt = resultado_indice %>%
                            dplyr::filter(porte_rodando == 'MT') %>%
                            dplyr::select(indice) %>%
                            dplyr::pull() %>%
                            base::as.numeric()

                          resultado_indice = resultado_indice[1, ] %>%
                            dplyr::mutate(
                              dplyr::across(.cols = -c(1, 2, 3), ~0)
                            ) %>%
                            dplyr::mutate(
                              indice = (indice_at * n_at + indice_mt * n_mt) / (n_at + n_mt)
                              , porte_rodando = 'ponderado (AT, MT)'
                            )

                        }# End: preciso ponderar

                        resultado_indice

                      }# End: Ponderando

                    resultados_calculados = resultados_calculados %>%
                      dplyr::select(indice, media, n, n_peso)

                  }# End = ponderar pelo porte (indice_at * n_at + indice_mt * n_mt) / (n_at + n_mt)

                }# End: calculando o índice

                resultado_indice = base::cbind(
                  #Amostra obtida
                  amostra = amostra,
                  #Amostra pretendida
                  amostra_prent = amostra_prent,
                  a_veri = base::ifelse(amostra == amostra_prent, "ok", "erro"),
                  #Soma de respostas que entraram no índice (após pivot); índice e média
                  resultados_calculados,
                  #Número de entrevistados (antes do pivot)
                  dados_split %>%
                    dplyr::select(dplyr::all_of(variaveis), "peso") %>%
                    {# Start: calculando
                      temp = tibble::tibble(.)

                      xtemp = temp %>%
                        dplyr::select(-c(peso)) %>%
                        dplyr::mutate(
                          dplyr::across(
                            .cols = dplyr::everything(),
                            ~ ifelse(base::is.na(.), 1, 0)
                          )
                        ) %>%
                        dplyr::mutate(x = base::rowSums(.)) %>%
                        dplyr::select(x) %>%
                        dplyr::pull()

                      temp$qwe = xtemp

                      temp = temp %>% dplyr::filter(
                        qwe < (base::length(variaveis) + 1)
                      ) %>%
                        dplyr::select(-qwe)

                      base::rm(xtemp)

                      temp
                    } %>% # End: calculando
                    dplyr::summarise(
                      linhas = dplyr::n(),
                      linhas_peso = base::sum(peso)
                    )
                  ,
                  #Escala aberta
                  banco_final %>%
                    dplyr::count(value, wt = peso) %>%
                    dplyr::mutate(pct_escala = n / base::sum(n) * 100) %>%
                    dplyr::select(-n) %>%
                    tidyr::pivot_wider(names_from = value, values_from = pct_escala) %>%
                    {# Start: analisando se ficou faltando algum

                      escala = tibble::tibble(.)
                      if( !base::any(base::colnames(escala) == "1") ){ escala$`1` = base::ifelse(amostra > 0, 0, NA) }
                      if( !base::any(base::colnames(escala) == "2") ){ escala$`2` = base::ifelse(amostra > 0, 0, NA) }
                      if( !base::any(base::colnames(escala) == "3") ){ escala$`3` = base::ifelse(amostra > 0, 0, NA) }
                      if( !base::any(base::colnames(escala) == "4") ){ escala$`4` = base::ifelse(amostra > 0, 0, NA) }
                      if( !base::any(base::colnames(escala) == "5") ){ escala$`5` = base::ifelse(amostra > 0, 0, NA) }
                      if( !base::any(base::colnames(escala) == "6") ){ escala$`6` = base::ifelse(amostra > 0, 0, NA) }
                      if( !base::any(base::colnames(escala) == "7") ){ escala$`7` = base::ifelse(amostra > 0, 0, NA) }
                      if( !base::any(base::colnames(escala) == "8") ){ escala$`8` = base::ifelse(amostra > 0, 0, NA) }
                      if( !base::any(base::colnames(escala) == "9") ){ escala$`9` = base::ifelse(amostra > 0, 0, NA) }
                      if( !base::any(base::colnames(escala) == "10") ){ escala$`10` = base::ifelse(amostra > 0, 0, NA) }
                      if( !base::any(base::colnames(escala) == "11") ){ escala$`11` = base::ifelse(amostra > 0, 0, NA) }
                      if( !base::any(base::colnames(escala) == "12") ){ escala$`12` = base::ifelse(amostra > 0, 0, NA) }

                      escala

                    } %>%# End: analisando se ficou faltando algum
                    dplyr::select(`1` , `2` , `3` , `4` , `5` , `6` , `7` , `8` , `9` , `10`, `11`, `12`)
                  ,
                  #Escala aberta com n
                  banco_final %>%
                    dplyr::count(value, wt = peso) %>%
                    tidyr::pivot_wider(names_from = value, values_from = n) %>%
                    {# Start: analisando se ficou faltando algum

                      escala = tibble::tibble(.)
                      if(!base::any(base::colnames(escala) == "1") ){ escala$`1` = base::ifelse(amostra > 0, 0, NA) }
                      if(!base::any(base::colnames(escala) == "2") ){ escala$`2` = base::ifelse(amostra > 0, 0, NA) }
                      if(!base::any(base::colnames(escala) == "3") ){ escala$`3` = base::ifelse(amostra > 0, 0, NA) }
                      if(!base::any(base::colnames(escala) == "4") ){ escala$`4` = base::ifelse(amostra > 0, 0, NA) }
                      if(!base::any(base::colnames(escala) == "5") ){ escala$`5` = base::ifelse(amostra > 0, 0, NA) }
                      if(!base::any(base::colnames(escala) == "6") ){ escala$`6` = base::ifelse(amostra > 0, 0, NA) }
                      if(!base::any(base::colnames(escala) == "7") ){ escala$`7` = base::ifelse(amostra > 0, 0, NA) }
                      if(!base::any(base::colnames(escala) == "8") ){ escala$`8` = base::ifelse(amostra > 0, 0, NA) }
                      if(!base::any(base::colnames(escala) == "9") ){ escala$`9` = base::ifelse(amostra > 0, 0, NA) }
                      if(!base::any(base::colnames(escala) == "10") ){ escala$`10` = base::ifelse(amostra > 0, 0, NA) }
                      if(!base::any(base::colnames(escala) == "11") ){ escala$`11` = base::ifelse(amostra > 0, 0, NA) }
                      if(!base::any(base::colnames(escala) == "12") ){ escala$`12` = base::ifelse(amostra > 0, 0, NA) }

                      escala

                    } %>% # End: analisando se ficou faltando algum
                    dplyr::rename(
                      indiceescala_nota1 = `1`,
                      indiceescala_nota2 = `2`,
                      indiceescala_nota3 = `3`,
                      indiceescala_nota4 = `4`,
                      indiceescala_nota5 = `5`,
                      indiceescala_nota6 = `6`,
                      indiceescala_nota7 = `7`,
                      indiceescala_nota8 = `8`,
                      indiceescala_nota9 = `9`,
                      indiceescala_nota10 = `10`,
                      indiceescala_nota11 = `11`,
                      indiceescala_nota12 = `12`
                    ) %>%
                    dplyr::select(
                      `indiceescala_nota1`,
                      `indiceescala_nota2`,
                      `indiceescala_nota3`,
                      `indiceescala_nota4`,
                      `indiceescala_nota5`,
                      `indiceescala_nota6`,
                      `indiceescala_nota7`,
                      `indiceescala_nota8`,
                      `indiceescala_nota9`,
                      `indiceescala_nota10`,
                      `indiceescala_nota11`,
                      `indiceescala_nota12`
                    )
                )

                resultado_indice = resultado_indice %>%
                  {
                    #Verificando se preciso excluir valores (tipo em ISQP)
                    resultado_indice = tibble::tibble(.)

                    if( indice_rodando$ponderar %in% c("imp_relativo") )
                    {

                      resultado_indice = resultado_indice %>%
                        dplyr::mutate(
                          dplyr::across(
                            .cols = -c(1, 2, 3, 4),
                            ~ 0
                          )
                        )

                    }

                    if( indice_rodando$ponderar %in% c("porte_at_mt") )
                    {

                      resultado_indice = resultado_indice %>%
                        dplyr::mutate(
                          dplyr::across(
                            .cols = -c(1, 2, 3, 4, base::ncol(resultado_indice)),
                            ~ 0
                          )
                        )

                    }

                    resultado_indice

                  } %>%
                  dplyr::mutate(
                    indice_sigla = all_indices[i],
                    indice_id = all_indice_id[i],
                    dono_id = dono_id,
                    split_nomes = split_nomes
                  ) %>%
                  dplyr::relocate(
                    c(dono_id, split_nomes, indice_id, indice_sigla),
                    c(1, 2, 3, 4, 5, 6, 7, 5)
                  ) %>%
                  dplyr::mutate(filtro = filtro_aplicado)

                IPpackage::print_cor(
                  df = base::paste0(
                    "dono_id ", dono_id, " = ", split_nomes,
                    " [", split_i, "/", base::unique(split_df$n_dono_ids_rodando), '] | ', all_indices[i],
                    " [", i, "/", base::length(all_indices), "]\n"
                  ),
                  cor = "amarelo",
                  italico = TRUE,
                  data.frame = FALSE
                )

                if( !base::is.null(o_que_remover) | !base::is.null(remover_dos_indices) )
                {# Start: Removendo valores para colunas específicas na tabela numérica

                  if( !base::is.null(o_que_remover) & !base::is.null(remover_dos_indices) )
                  {# Start: Se ambos são null, fazer

                    if( all_indices[i] %in% remover_dos_indices)
                    {# Start: se for um dos índices indicados para remoção, remover

                      # Removendo valores para colunas específicas na tabela numérica
                      resultado_indice = resultado_indice %>%
                        dplyr::mutate(
                          dplyr::across(
                            .cols = dplyr::all_of(o_que_remover),
                            ~NA
                          )
                        )

                    }# End: se for um dos índices indicados para remoção, remover

                  } else {# End: Se ambos são null, fazer
                    # Se ambos não são null's, erro

                    base::stop("'o_que_remover' e 'remover_dos_indices'\nOu ambos os par\u00E2metros s\u00E3o NULL ou ambos n\u00E3o s\u00E3o NULL")

                  }

                }# End: Removendo valores para colunas específicas na tabela numérica

                resultado_indice

              }# End: Calculando

              if( i == 1 )
              {# Start: armazenando os resultados

                resultado_dono_id = resultado_indice

              } else {

                resultado_dono_id = base::rbind(resultado_dono_id, resultado_indice)

              }# End: armazenando os resultados

              base::rm(resultado_indice)

            }# End: fazer para cada indice

          }# End: fazer para cada dono_id

          resultado_dono_id = resultado_dono_id %>%
            dplyr::mutate(pasta_nome = pasta_nome)

          # Start: Salvar no arquivo temporário (não armazenar no MiB do R)
          base::saveRDS(
            object = resultado_dono_id,
            file = base::paste0(caminho_temporario, "/", dono_id, ".rds")
          )

          IPpackage::print_cor(
            df = base::paste0("Salvo em '", base::paste0(caminho_temporario, "/", dono_id, "'.rds"), "\n"),
            cor = "azul",
            negrito = FALSE,
            data.frame = FALSE
          )

          base::rm(resultado_dono_id)

        }# End: se tenho dados, rodar

      }# End: para cada split_id

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

  if( !base::is.null(importancia_q5_pc) )
  {# Start: colocando a importância

    IPpackage::print_cor(
      df = "Colocando a import\u00E2ncia",
      cor = "amarelo",
      negrito = TRUE,
      data.frame = FALSE
    )

    # Lidando com os importanciaq5_relativa repetidos
    b = importancia_q5_pc %>%
      dplyr::group_by(dono_id, indice_id) %>%
      dplyr::group_split() %>%
      purrr::imap(
        ~{# Start: se tenho duplicata, pegar o que tiver a Fonte = 'fonte_imp_entra_serie_hist'

          # bi = 150; bdf = b[[bi]]
          bi = .y
          bdf = .x

          if( base::nrow(bdf) > 1 )
          {# Start: se mais de uma linha, filtrar

            bdf = bdf %>%
              dplyr::filter(Fonte %in% c(fonte_imp_entra_serie_hist))

            if( base::nrow(bdf) > 1 & base::length(base::unique(base::round(bdf$importanciaq5_relativa, 5))) <= 1 )
            {# Start: se continuar com mais de 1 linha e os importanciaq5_relativa forem iguais, pegar qualquer um deles

              bdf = bdf[1, ]

            }# End: se continuar com mais de 1 linha e os importanciaq5_relativa forem iguais, pegar qualquer um deles

          }# End: se mais de uma linha, filtrar

          if( base::nrow(bdf) > 1 )
          {# Start: se continuar com mais de 1 linha: erro

            base::stop("Mais de uma 'importanciaq5_relativa' e n\u00E3o consegui lidar com esse caso")

          }# End: se continuar com mais de 1 linha: erro

          bdf

        }# End: se tenho duplicata, pegar o que tiver a Fonte = 'fonte_imp_entra_serie_hist'
      ) %>%
      dplyr::bind_rows()

    a = resultado_final %>%
      dplyr::left_join(
        b %>%
          dplyr::mutate(
            seriehistorica_importancia = base::ifelse(
              importanciaq5_relativa == 0,
              importanciaq5_porcentagem / 100,
              importanciaq5_relativa / 100
            )
          ) %>%
          dplyr::select(dono_id, indice_id, seriehistorica_importancia),
        by = c(
          "dono_id" = "dono_id",
          "indice_id" = "indice_id"
        )
      ) %>%
      dplyr::mutate(
        dplyr::across(
          .cols = -c(1:4),
          ~base::ifelse(base::is.na(.), 0, .)
        )
      )

    a = a %>%
      dplyr::left_join(
        qualidade_preco %>%
          dplyr::select(dono_id, qualidadepreco_qualidade, qualidadepreco_preco),
        by = c("dono_id" = "dono_id")
      ) %>%
      dplyr::mutate(
        dplyr::across(
          .cols = c("qualidadepreco_qualidade", "qualidadepreco_preco"),
          ~ base::ifelse(base::is.na(.), 0, .)
        )
      ) %>%
      dplyr::mutate(
        seriehistorica_importancia = base::ifelse(
          indice_sigla == referencia_indices %>%
            dplyr::filter(imp_relativa_quali_preco_nome == "qualidade") %>%
            dplyr::select(IDAR) %>%
            dplyr::pull(),
          qualidadepreco_qualidade,
          seriehistorica_importancia
        ),
        seriehistorica_importancia = base::ifelse(
          indice_sigla == referencia_indices %>%
            dplyr::filter(imp_relativa_quali_preco_nome == "preco") %>%
            dplyr::select(IDAR) %>%
            dplyr::pull(),
          qualidadepreco_preco,
          seriehistorica_importancia)
      ) %>%
      dplyr::select(-c(qualidadepreco_qualidade, qualidadepreco_preco))

    resultado_final = a

    base::rm(a, b) %>%
      base::suppressWarnings()

  }# End: colocando a importância

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

}# End: calculo_indices_gc_abradee
