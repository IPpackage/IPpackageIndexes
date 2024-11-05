#' IndiceCalculoImportancia
#'
#' Realiza o cálculo da importância de índices em dados particionados, aplicando transformações, filtros e ponderações
#' para calcular importância relativa, notas médias, e correlações com atributos de qualidade e preço.
#'
#' @param dados_split Lista de data frames, onde cada elemento representa um conjunto particionado de dados
#' de uma distribuidora específica, com informações sobre índices e atributos.
#' @param all_importancia Vetor de strings indicando os nomes de índices que devem ser considerados no cálculo da importância.
#' @param referencia Data frame de referência que contém as definições e atributos de cada índice, incluindo configurações
#' de fatores de correção e critérios de inclusão.
#'
#' @details
#' A função `IndiceCalculoImportancia` divide os dados por distribuidora, aplica um filtro de condições de cálculo e calcula índices de importância, notas, e
#' atributos relativos a cada distribuição. O cálculo inclui a avaliação dos atributos com base em ponderações configuradas,
#' aplicando ajustes específicos para tipos de índice, como \code{"ISQP"} e \code{"ISCAL"}. Adicionalmente, calcula a correlação entre qualidade e preço
#' conforme especificado no parâmetro de referência. A função retorna os resultados como uma lista com quatro componentes.
#'
#' @section Etapas de Cálculo:
#' - **Configuração Inicial e Filtros**: Gera filtros de formatação e identifica as distribuidoras e atributos aplicáveis.
#' - **Ordens e Notas**: Inverte e pondera a ordem dos atributos e calcula notas médias para cada atributo avaliado.
#' - **Correção de Peso e Cálculo Final**: Aplica fatores de correção, calcula a importância ponderada e percentual, e organiza as saídas por distribuidora.
#' - **Correlação Qualidade-Preço**: Calcula métricas de correlação de qualidade e preço para os atributos especificados.
#'
#' @return Retorna uma lista com quatro data frames:
#'   - `importancia_nota`: Notas médias calculadas para cada atributo e distribuidora.
#'   - `importancia_peso`: Importância relativa calculada com base nos pesos atribuídos a cada atributo.
#'   - `importancia_q5_pc`: Cálculo ponderado e ajustado da importância por atributo.
#'   - `qualidade_preco`: Dados de correlação entre qualidade e preço, por distribuidora.
#'
#' @import dplyr
#' @import purrr
#' @import tidyr
#' @import stringr
#'
#' @examples
#'
#' base::print("Sem Exemplo")
#'
#' @export

IndiceCalculoImportancia <- function(
    dados_split,
    all_importancia,
    referencia
)
{# Start: funcao_importancia

  resultado_final = dados_split %>%
    purrr::imap(
      ~ {# Start: rodando cada um dos splits que quero

        # split_i = 1; split_df = dados_split[[split_i]]
        split_i = .y
        split_df = .x

        {# Start: informações sobre a distribuidora que estou rodando

          dono_id = split_df$dono_id %>%
            base::unique()

          distribuidora_id = split_df$distribuidora_id %>%
            base::unique()

          pasta_nome = split_df$pasta_nome %>%
            base::unique()

          filtro = split_df %>%
            dplyr::select(filtro_quest)%>%dplyr::pull()%>%unique()%>%
            stringr::str_replace_all("="," == ")%>%
            stringr::str_replace_all("<>","!=")%>%
            stringr::str_replace_all("like"," == ")%>%
            stringr::str_replace_all("LIKE"," == ")%>%
            stringr::str_replace_all(" and "," & ")%>%
            stringr::str_replace_all(" AND "," & ")%>%
            stringr::str_replace_all(" or "," | ")%>%
            stringr::str_replace_all(" OR "," | ")%>%
            stringr::str_replace_all(" not in [(]","%nin%c(")%>%
            stringr::str_replace_all(" NOT IN [(]","%nin%c(")%>%
            stringr::str_replace_all(" in [(]","%in%c(")%>%
            stringr::str_replace_all(" IN [(]","%in%c(")%>%
            stringr::str_replace_all(" in[(]","%in%c(")%>%
            stringr::str_replace_all(" IN[(]","%in%c(")%>%
            stringr::str_replace_all("===","==")%>%
            stringr::str_replace_all("==  ==","==")%>%
            stringr::str_replace_all("< ==","<=")%>%
            stringr::str_replace_all("> ==",">=")

          split_tipo = split_df %>%
            dplyr::select(split_tipo) %>%
            dplyr::pull() %>%
            base::unique()

          split_nomes = split_df %>%
            dplyr::select(split_nome) %>%
            dplyr::pull() %>%
            base::unique()

          rodo_importancia = referencia %>%
            dplyr::filter(imp_nota %>% stringr::str_detect(",")) %>%
            dplyr::select(IDAR, imp_nota) %>%
            dplyr::mutate(
              n_linhas_sem_na = NA,
              entra = NA
            ) %>%
            {# Start: verificando, via nº de linhas quais vars realmente entram

              temp = tibble::tibble(.)

              for( ntemp in 1:base::nrow(temp) )
              {# Start: verificando cada importância

                nrow_filtro = dados[dados$dono_id == dono_id, ] %>%
                  dplyr::select(
                    base::strsplit(temp$imp_nota[ntemp], ",")[[1]]
                  ) %>%
                  dplyr::filter(
                    !dplyr::if_all(
                      .cols = dplyr::everything(),
                      .fns = base::is.na
                    )
                  ) %>%
                  base::nrow()

                temp$n_linhas_sem_na[ntemp] = nrow_filtro

                # Se filtro resultou em um banco com 0 linhas = sem dados para essa importância = não calcular
                temp$entra[ntemp] = base::ifelse(nrow_filtro == 0, "N", "S")

                base::rm(nrow_filtro)

              }# End: verificando cada importância

              temp

            }# End: verificando, via nº de linhas quais vars realmente entram

          all_importancia_rodar = rodo_importancia %>%
            dplyr::filter(entra == "S") %>%
            dplyr::select(IDAR) %>%
            dplyr::pull()


          # Printando as informações
          IPpackage::print_cor(
            df = "----------------------------------------------------------------\n",
            cor = "ciano",
            data.frame = FALSE
          )

          IPpackage::print_cor(
            df = base::paste0(
              "dono_id ", dono_id, " = ", split_nomes, " [", split_i, "/", base::length(dados_split),
              "] Import\u00E2ncias: ", base::paste0(all_importancia_rodar, collapse = ", "), "\n"
            ),
            cor = "verde",
            negrito = TRUE,
            data.frame = FALSE
          )

        }# End: informações sobre a distribuidora que estou rodando

        #Calculando cada um dos índices
        i = base::which(all_importancia_rodar == "ISCAL")

        for( i in 1:length(all_importancia_rodar) )
        {# Start: fazer para cada indice (i)

          {# Start: Informações e filtros

            referenciax = referencia

            if( all_importancia_rodar[i] %in% c("ISQP", "ISCAL") )
            {# Start: se for ISQP ou ISCAL, vou calcular as áreas

              referenciax = referenciax %>%
                dplyr::mutate(
                  imp_nota = imp_area_nota,
                  imp_ordem = imp_area_ordem,
                  imp_ordem_var = imp_area_ordem_var
                )

            }# End: se for ISQP ou ISCAL, vou calcular as áreas

            indice_rodando = referenciax %>%
              dplyr::filter(IDAR == all_importancia_rodar[i])

            filtro_aplicado = filtro

            vars_nota_imp = rodo_importancia %>%
              dplyr::filter(IDAR == all_importancia_rodar[i]) %>%
              dplyr::select(imp_nota) %>%
              dplyr::pull() %>%
              base::strsplit(",") %>%
              base::unlist()

            dados_split_rodando = dados[dados$dono_id == dono_id, ]%>%
              dplyr::filter(
                base::eval(
                  base::parse(
                    text = filtro_aplicado
                  )
                )
              )

            dados_ordem = dados_split_rodando %>%
              dplyr::select(id, dplyr::all_of(vars_nota_imp)) %>%
              dplyr::filter(!dplyr::if_all(-id, is.na))

            amostra = base::nrow(dados_split_rodando)

            imp_ordem_df = referenciax %>%
              dplyr::mutate(
                IDAR2 = IDAR %>%
                  stringr::str_replace_all('[0-9]', '') %>%
                  stringr::str_replace_all(" ", "")
              ) %>%
              dplyr::filter(
                IDAR2 == all_importancia_rodar[i],
                !variaveis_geral %>%
                  stringr::str_detect(",")
              ) %>%
              dplyr::select(IDAR, imp_nota, imp_ordem_var, imp_ordem, imp_entra_ordem, importancia_tipo,imp_fator_correcao)

            if( all_importancia_rodar[i] %in% c("ISQP", "ISCAL") )
            {# Start: se for ISQP ou ISCAL, vou calcular as áreas

              imp_ordem_df = referenciax %>%
                dplyr::filter(imp_nota %in% c(vars_nota_imp)) %>%
                dplyr::select(IDAR, imp_nota, imp_ordem_var, imp_ordem, imp_entra_ordem, importancia_tipo,imp_fator_correcao)

            }# End: se for ISQP ou ISCAL, vou calcular as áreas

            {# Start: imp_ordem (ordem da importância)

              imp_ordem = imp_ordem_df %>%
                dplyr::select(imp_ordem) %>%
                dplyr::pull() %>%
                base::as.character()

              names(imp_ordem) <- imp_ordem_df %>%
                dplyr::select(imp_ordem_var) %>%
                dplyr::pull() %>%
                base::as.character()

            }# End: imp_ordem (ordem da importância)

            {# Start: imp_notas (notas da importância)

              imp_notas = imp_ordem_df %>%
                dplyr::select(IDAR) %>%
                dplyr::pull() %>%
                base::as.character()

              names(imp_notas) <- imp_ordem_df %>%
                dplyr::select(imp_nota) %>%
                dplyr::pull() %>%
                base::as.character()

            }# End: imp_notas (notas da importância)

          }# End: Informações e filtros

          {# Start: calculando para cada all_importancia_rodar

            # Para calcular a importância é preciso criar variáveis que possuam a
            # posição no ranking de cada atributo invertida.
            # Ou seja, quanto mais importante o atributo, maior o valor da posição

            {# Start: inv_ordem (inverter a ordem da importância)

              inv_ordem =
                dados_split_rodando %>%
                dplyr::select(id, dplyr::all_of(base::names(imp_ordem))) %>%
                # Excluir as linhas que estejam 100% vazias
                dplyr::filter(!dplyr::if_all(-id, is.na)) %>%
                # Excluir as colunas que estejam 100% vazias
                dplyr::select_if(~!base::all(base::is.na(.))) %>%
                tidyr::pivot_longer(
                  names_to = 'var',
                  values_to = 'val',
                  cols = -id
                ) %>%
                tidyr::pivot_wider(
                  names_from = 'val',
                  values_from = 'var',
                  names_prefix = all_importancia_rodar[i]
                ) %>%
                dplyr::mutate(
                  dplyr::across(
                    -id,
                    ~base::as.numeric(
                      dplyr::recode(
                        .x,
                        !!!imp_ordem
                      )
                    )
                  )
                ) %>%
                {# Start: se algum não entrar na ordem, remover

                  a = tibble::tibble(.)

                  if( base::any(imp_ordem_df$imp_entra_ordem == FALSE) )
                  {# Start: se algum não entra, vou remover aqui

                    remover_ordem = imp_ordem_df %>%
                      dplyr::filter(imp_entra_ordem == FALSE) %>%
                      dplyr::select(IDAR) %>%
                      dplyr::pull()

                    if( base::any(colnames(a) == remover_ordem) )
                    {# Start: se tenho alguma coluna com o nome do que deve ser excluído. Excluir

                      colnames(a)[colnames(a) == remover_ordem] = "remover_ordemx"

                      IPpackage::print_cor(
                        df = base::paste0("N\u00E3o entra no c\u00E1lculo da ordem: ", remover_ordem, " (removido)\n"),
                        cor = "vermelho",
                        italico = TRUE,
                        data.frame = FALSE
                      )

                      # Remover e refazer o rank
                      a = a %>%
                        dplyr::mutate(
                          dplyr::across(
                            -id,
                            ~base::ifelse(
                              base::is.na(remover_ordemx),
                              .x,
                              base::ifelse(
                                .x < remover_ordemx,
                                .x,
                                .x-1
                              )
                            )
                          )
                        ) %>%
                        dplyr::select(-remover_ordemx)

                    }# End: se tenho alguma coluna com o nome do que deve ser excluído. Excluir

                  }# End: se algum não entra, vou remover aqui

                  a

                } %>%# End: se algum não entrar na ordem, remover
                dplyr::mutate(n_atribulos = base::ncol(.) -1 ) %>%
                dplyr::mutate(
                  dplyr::across(
                    -id,
                    ~n_atribulos - .x + 1
                  )
                ) %>%
                dplyr::select(-n_atribulos)

            }# End: inv_ordem (inverter a ordem da importância)

            # Nas tabelas a seguir, são apresentadas as posições nos rankings
            # invertidas para todos os atributos das 5 áreas de qualidade e para o
            # conjunto das áreas.

            {# Start: soma_ordem (soma e peso)

              soma_ordem <- inv_ordem %>%
                dplyr::summarise(
                  dplyr::across(
                    .cols = -id, .fns = ~base::sum(.x)
                  )
                ) %>%
                tidyr::pivot_longer(
                  names_to = 'id_var',
                  values_to = 'importanciapeso_soma',
                  cols = dplyr::everything()
                ) %>%
                dplyr::arrange(id_var) %>%
                dplyr::mutate(
                  importanciapeso_peso = importanciapeso_soma / base::max(importanciapeso_soma)
                )

              if( all_importancia_rodar[i] %in% c("ISQP", "ISCAL") )
              {# Start: se for ISQP ou ISCAL, vou calcular as áreas (trocar os nomes para as áreas)

                soma_ordem = dplyr::left_join(
                  soma_ordem %>%
                    dplyr::rename(id = id_var),
                  imp_ordem_df %>%
                    dplyr::mutate(
                      id = base::paste0(all_importancia_rodar[i],imp_ordem)
                    ) %>%
                    dplyr::rename(id_var = IDAR) %>%
                    dplyr::select(id_var, id),
                  by = c("id" = "id")
                ) %>%
                  dplyr::select(-id) %>%
                  dplyr::relocate(id_var, .before = importanciapeso_soma)

              }# End: se for ISQP ou ISCAL, vou calcular as áreas (trocar os nomes para as áreas)

              # Adicionando informações
              soma_ordem = soma_ordem %>%
                dplyr::left_join(
                  referenciax %>%
                    dplyr::select(IDAR, indice_id),
                  by = c("id_var" = "IDAR")
                ) %>%
                dplyr::mutate(
                  importanciapeso_id = NA,
                  dono_id = dono_id,
                  importanciapeso_tipo = base::unique(imp_ordem_df$importancia_tipo)
                ) %>%
                dplyr::select(
                  id_var,
                  importanciapeso_id,
                  importanciapeso_soma,
                  importanciapeso_peso,
                  indice_id,
                  dono_id,
                  importanciapeso_tipo
                )


            }# End: soma_ordem (soma e peso)

            {# Start: dados_notas (Calcula as notas médias de importância)

              calculo_vars = dados_ordem %>%
                dplyr::select(id, dplyr::all_of(vars_nota_imp)) %>%
                tidyr::pivot_longer(names_to = 'item', values_to = 'valor', cols = -id) %>%
                # dplyr::mutate(item = dplyr::recode(item, !!!imp_notas)) %>%
                dplyr::mutate(item = item) %>%
                dplyr::group_by(item) %>%
                dplyr::group_split() %>%
                purrr::imap(
                  ~{# Start: fazendo para cada um dos imp_notas

                    # impi = 1; dfi= xdf[[impi]]
                    impi = .y
                    dfi = .x

                    dfi_notas = tibble::tibble(
                      valor = 1:12
                    ) %>%
                      dplyr::left_join(
                        dfi %>%
                          dplyr::count(valor) %>%
                          dplyr::mutate(pct = n / base::sum(n) * 100),
                        by = c("valor" = "valor")
                      ) %>%
                      dplyr::mutate(
                        nome = base::paste0("importancianota_nota", valor)
                      ) %>%
                      dplyr::select(nome, n) %>%
                      tidyr::pivot_wider(
                        names_from = nome,
                        values_from = n
                      )

                    dfi_media = dfi %>%
                      dplyr::filter(valor<=10) %>%
                      dplyr::summarise(
                        importancianota_media = base::mean(`valor`)
                      )

                    dfi_fim = base::cbind(dfi_notas, dfi_media) %>%
                      dplyr::mutate(
                        item = base::unique(dfi$item)
                      )

                    base::rm(dfi_notas, dfi_media)

                    dfi_fim

                  }# End: fazendo para cada um dos imp_notas

                ) %>%
                dplyr::bind_rows()

              dados_notas <- tibble::tibble(
                nome = imp_notas,
                item = base::names(imp_notas)
              ) %>%
                dplyr::left_join(
                  calculo_vars,
                  by = c("item" = "item")
                ) %>%
                dplyr::mutate(item = nome) %>%
                dplyr::select(- nome) %>%
                dplyr::relocate(item, .after = importancianota_media)

              rm(calculo_vars)


              # Adicionando informações
              dados_notas = dados_notas %>%
                dplyr::rename(id_var = item) %>%
                dplyr::left_join(
                  referenciax %>%
                    dplyr::select(IDAR, indice_id),
                  by = c("id_var" = "IDAR")
                ) %>%
                dplyr::mutate(
                  importancianota_id = NA,
                  dono_id = dono_id,
                  importancianota_tipo = base::unique(imp_ordem_df$importancia_tipo)
                ) %>%
                dplyr::mutate(
                  importancianota_tipo = dplyr::case_when(
                    #Por algum motivo, no servidor, o 1 é 0
                    importancianota_tipo %in% c(1) ~ 0,
                    TRUE ~ importancianota_tipo
                  )
                ) %>%
                dplyr::relocate(id_var,.before = importancianota_nota1) %>%
                dplyr::relocate(importancianota_id,.before = importancianota_nota1)

            }# End: dados_notas (Calcula as notas médias de importância)

            {# Start: tabela_atb (ponderar pelo fator de correção)

              idars_entram = imp_ordem_df %>%
                dplyr::filter(imp_entra_ordem == TRUE) %>%
                dplyr::select(IDAR) %>%
                dplyr::pull()

              tabela_atb = dados_notas %>%
                dplyr::filter(id_var %in% idars_entram) %>%
                dplyr::select(id_var, importancianota_media, indice_id, dono_id) %>%
                dplyr::rename(importanciaq5_nota = importancianota_media) %>%
                dplyr::left_join(
                  soma_ordem %>%
                    dplyr::select(id_var, importanciapeso_peso) %>%
                    dplyr::rename(importanciaq5_peso = importanciapeso_peso),
                  by = c("id_var" = "id_var")
                ) %>%
                dplyr::left_join(
                  imp_ordem_df %>%
                    dplyr::select(IDAR, imp_fator_correcao) %>%
                    dplyr::rename(importanciaq5_fator = imp_fator_correcao),
                  by = c("id_var" = "IDAR")
                ) %>%
                dplyr::mutate(
                  importanciaq5_notaponderada = importanciaq5_peso * importanciaq5_nota,
                  importanciaq5_porcentagem = importanciaq5_notaponderada / base::sum(importanciaq5_notaponderada) * 100
                  # O resto, calcular no final
                )

            }# End: tabela_atb (ponderar pelo fator de correção)

            if( i == 1 )
            {# Start: armazenando o resultado

              resultado_soma_ordem_i = soma_ordem

              dados_notas_i = dados_notas

              tabela_atb_i = tabela_atb


            } else {

              resultado_soma_ordem_i = base::rbind(
                resultado_soma_ordem_i,
                soma_ordem
              )

              dados_notas_i = base::rbind(
                dados_notas_i,
                dados_notas
              )

              tabela_atb_i = base::rbind(
                tabela_atb_i,
                tabela_atb
              )

            }# End: armazenando o resultado

            IPpackage::print_cor(
              df = base::paste0("dono_id ",dono_id," = ",split_nomes," [",split_i,"/",length(dados_split),'] | ',all_importancia_rodar[i]," [",i,"/",length(all_importancia_rodar),"]\n"),
              cor = "amarelo",
              italico = TRUE,
              data.frame = FALSE
            )

          }# End: calculando para cada all_importancia_rodar

          base::rm(referenciax, soma_ordem, dados_notas, tabela_atb)

        }# End: fazer para cada indice (i)

        {# Start: Continuar calculos de tabela_atb

          vars_area = referencia %>%
            dplyr::filter(!base::is.na(imp_area_nota)) %>%
            dplyr::select(IDAR) %>%
            dplyr::pull()

          tabela_atb = tabela_atb_i %>%
            # dplyr::filter(id_var %in% idars_entram) %>%
            dplyr::mutate(id_area = stringr::str_remove_all(id_var, '[0-9]'))

          tabela_atb = tabela_atb %>%
            dplyr::left_join(
              tabela_atb %>%
                dplyr::filter(id_var %in% vars_area) %>%
                dplyr::rename(
                  nota_area = importanciaq5_nota,
                  peso_area = importanciaq5_peso,
                  notaponderada_area = importanciaq5_notaponderada,
                  porcentagem_area = importanciaq5_porcentagem
                ) %>%
                dplyr::select(id_var, nota_area, peso_area, notaponderada_area, porcentagem_area),
              by = c("id_area" = "id_var")
            )

          #Calculos
          tabela_atb = tabela_atb %>%
            dplyr::mutate(
              importanciaq5_id = NA,
              importanciaq5_porcentagemarea = (importanciaq5_porcentagem * porcentagem_area) / 100,
              importanciaq5_corrigida = importanciaq5_porcentagemarea * importanciaq5_fator,
            ) %>%
            dplyr::mutate(
              importanciaq5_porcentagemarea = dplyr::case_when(
                id_var %in% vars_area ~ 0,
                TRUE ~ importanciaq5_porcentagemarea
              ),
              importanciaq5_fator = dplyr::case_when(
                id_var %in% vars_area ~ 0,
                TRUE ~ importanciaq5_fator
              ),
              importanciaq5_corrigida = dplyr::case_when(
                id_var %in% vars_area ~ 0,
                TRUE ~ importanciaq5_corrigida
              )
            ) %>%
            dplyr::mutate(
              importanciaq5_relativa = importanciaq5_corrigida / base::sum(importanciaq5_corrigida, na.rm = TRUE) * 100
            ) %>%
            dplyr::select(
              id_var,
              importanciaq5_id,
              importanciaq5_nota,
              importanciaq5_peso,
              importanciaq5_notaponderada,
              importanciaq5_porcentagem,
              importanciaq5_porcentagemarea,
              importanciaq5_fator,
              importanciaq5_corrigida,
              importanciaq5_relativa,
              indice_id,
              dono_id
            )

        }# End: Continuar calculos de tabela_atb

        {# Start: Calculando qualidade_preco para cada dono_id

          ref_qualidade_preco = referencia %>%
            dplyr::filter(!base::is.na(imp_relativa_quali_preco))

          vars_qualidade_preco = ref_qualidade_preco %>%
            dplyr::select(imp_relativa_quali_preco) %>%
            dplyr::pull()

          qualidade_preco_i = dados_split_rodando %>%
            dplyr::select(dplyr::all_of(vars_qualidade_preco)) %>%
            tidyr::pivot_longer(cols = dplyr::everything(), names_to = "variavel", values_to = "value") %>%
            dplyr::left_join(
              ref_qualidade_preco %>%
                dplyr::select(imp_relativa_quali_preco, imp_relativa_quali_preco_nome),
              by = c("variavel" = "imp_relativa_quali_preco")
            ) %>%
            dplyr::group_by(imp_relativa_quali_preco_nome) %>%
            dplyr::summarise(
              media = base::mean(value, na.rm = TRUE) / 100
            ) %>%
            dplyr::mutate(
              imp_relativa_quali_preco_nome = base::paste0("qualidadepreco_",imp_relativa_quali_preco_nome )
            ) %>%
            tidyr::pivot_wider( names_from = imp_relativa_quali_preco_nome, values_from = media) %>%
            dplyr::mutate(
              qualidadepreco_id = NA,
              dono_id = dono_id
            ) %>%
            dplyr::select(qualidadepreco_id, qualidadepreco_qualidade, qualidadepreco_preco, dono_id)

          base::rm(ref_qualidade_preco, vars_qualidade_preco)

        }# End: Calculando qualidade_preco para cada dono_id

        resultado_soma_ordem = resultado_soma_ordem_i %>%
          dplyr::mutate(pasta_nome = pasta_nome)

        resultado_dados_notas = dados_notas_i %>%
          dplyr::mutate(pasta_nome = pasta_nome)

        resultado_tabela_atb = tabela_atb %>%
          dplyr::mutate(pasta_nome = pasta_nome) %>%
          {# Start: se quero remover quem não tem dados em imp_area_nota

            resultado_tabela_atb = tibble::tibble(.)

            # if( remover_quem_n_tem_imp_area_nota == TRUE)
            # {
            #
            #   # para remover RE de CEPM-DO pq ele não tem imp_area_nota. Portanto, sem importância pra RE
            #   resultado_tabela_atb = resultado_tabela_atb %>%
            #     dplyr::filter(!base::is.na(importanciaq5_relativa))
            #
            # }

            resultado_tabela_atb

          }# End: se quero remover quem não tem dados em imp_area_nota

        qualidade_preco = qualidade_preco_i %>%
          dplyr::mutate(pasta_nome = pasta_nome)

        fim = base::list(
          "soma" = resultado_soma_ordem,
          "notas" = resultado_dados_notas,
          "importancia" = resultado_tabela_atb,
          "qualidade_preco" = qualidade_preco
        )

        fim

      }#  End: rodando cada um dos splits que quero
    )

  {# Start: Pegando os resultados

    importancia_nota = base::list()
    importancia_peso = base::list()
    importancia_q5_pc = base::list()
    qualidade_preco = base::list()

    for ( i in 1:base::length(resultado_final))
    {# Start: pegando todos os dono_id rodados

      importancia_nota[[i]] = resultado_final[[i]][["notas"]]
      importancia_peso[[i]] = resultado_final[[i]][["soma"]]
      importancia_q5_pc[[i]] = resultado_final[[i]][["importancia"]]
      qualidade_preco[[i]] = resultado_final[[i]][["qualidade_preco"]]


    }# End: pegando todos os dono_id rodados

    importancia_nota = importancia_nota %>%
      dplyr::bind_rows()

    importancia_peso = importancia_peso %>%
      dplyr::bind_rows()

    importancia_q5_pc = importancia_q5_pc %>%
      dplyr::bind_rows()

    qualidade_preco = qualidade_preco %>%
      dplyr::bind_rows()

    base::rm(i)

    }# End: Pegando os resultados

  return(
    base::list(
      importancia_nota = importancia_nota,
      importancia_peso = importancia_peso,
      importancia_q5_pc = importancia_q5_pc,
      qualidade_preco = qualidade_preco
    )
  )

}# End: funcao_importancia
