#' IndiceVerificandoIndices
#'
#' Esta função compara resultados de índices calculados localmente com uma série histórica do servidor, verificando discrepâncias
#' e salvando logs de erro de forma organizada em arquivos temporários e permanentes.
#'
#' @param resultado_final Data frame contendo os resultados de índices calculados localmente.
#' @param serie_historica_servidor Data frame com os valores de referência para a série histórica de índices, obtidos do servidor.
#' @param caminho_temporario (Opcional) String com o caminho onde os arquivos temporários serão armazenados durante a execução. Padrão: `"99.temp cal erro"`.
#' @param caminho_resultado String com o caminho onde os arquivos de erro finais serão armazenados. Padrão: `"Erros"`.
#'
#' @details
#' A função `verificando_indices` organiza e verifica possíveis discrepâncias entre os dados calculados localmente e a série histórica
#' do servidor em várias dimensões, incluindo número de linhas, amostras pretendidas e comparações de variáveis de importância.
#' Discrepâncias são identificadas e armazenadas como objetos RDS em subdiretórios temporários específicos.
#' Cada erro detectado (como número de linhas incorreto ou discrepâncias em importância) é salvo em um arquivo RDS temporário para posterior
#' agregação em um local de armazenamento permanente.
#'
#' @section Etapas de Verificação:
#' - **Configuração de Diretórios Temporários**: Cria e limpa subdiretórios para armazenamento temporário de logs de erro.
#' - **Verificação por Grupo**: Processa cada grupo (`dono_id`) em `resultado_final` e verifica o número de linhas e dados de amostra,
#'   comparando-os com a série histórica do servidor.
#' - **Comparação de Índices**: Compara índices calculados localmente e armazenados no servidor e identifica discrepâncias em
#'   variáveis de `media`, `indice`, e `importancia`.
#' - **Salvamento de Resultados**: Logs de erro consolidados são salvos em arquivos RDS finais e organizados por tipo no
#'   diretório especificado.
#'
#' @return A função não retorna um objeto em R. Em vez disso, salva arquivos RDS com logs de erros em `caminho_resultado` e apaga
#' o conteúdo do diretório temporário.
#'
#' @import dplyr
#' @import purrr
#' @import tibble
#'
#' @examples
#'
#' base::print("Sem Exemplo")
#'
#' @export

IndiceVerificandoIndices = function(
    resultado_final,
    serie_historica_servidor,
    caminho_temporario = "99.temp cal erro",
    caminho_resultado = "Erros"
)
{# Start: Função 'IndiceVerificandoIndices'

  {# Start: Criando para p/salvar arquivo temporário (não armazenar no MiB do R)

    # Criar o diretório temporário
    base::dir.create(caminho_temporario) %>%
      base::suppressWarnings()

    # Remover todos os arquivos e subdiretórios dentro do diretório temporário, para começar com um diretório limpo
    base::unlink(
      x = base::paste0(caminho_temporario, "/*"),
      recursive = TRUE
    )

    # Criar o diretório temporário
    caminho_temporario_erro_n_linhas = base::paste0(caminho_temporario, "/erro_n_linhas")
    base::dir.create(caminho_temporario_erro_n_linhas) %>%
      base::suppressWarnings()
    base::unlink(
      x = base::paste0(caminho_temporario_erro_n_linhas, "/*"),
      recursive = TRUE
    )

    caminho_temporario_erro_amostra_pretend = base::paste0(caminho_temporario, "/erro_amostra_pretend")
    base::dir.create(caminho_temporario_erro_amostra_pretend) %>%
      base::suppressWarnings()
    base::unlink(
      x = base::paste0(caminho_temporario_erro_amostra_pretend, "/*"),
      recursive = TRUE
    )

    caminho_temporario_erro_i = base::paste0(caminho_temporario, "/erro_i")
    base::dir.create(caminho_temporario_erro_i) %>%
      base::suppressWarnings()
    base::unlink(
      x = base::paste0(caminho_temporario_erro_i, "/*"),
      recursive = TRUE
    )

    caminho_temporario_erro_imp = base::paste0(caminho_temporario, "/erro_imp")
    base::dir.create(caminho_temporario_erro_imp) %>%
      base::suppressWarnings()
    base::unlink(
      x = base::paste0(caminho_temporario_erro_imp, "/*"),
      recursive = TRUE
    )

    caminho_temporario_erro_m = base::paste0(caminho_temporario, "/erro_m")
    base::dir.create(caminho_temporario_erro_m) %>%
      base::suppressWarnings()
    base::unlink(
      x = base::paste0(caminho_temporario_erro_m, "/*"),
      recursive = TRUE
    )

  }# End: Criando para p/salvar arquivo temporário (não armazenar no MiB do R)

  # Calculando
  resultado_final %>%
    dplyr::group_by(dono_id) %>%
    dplyr::group_split() %>%
    purrr::imap(

      ~{# Start: verificando cada um dos dono_id calculados no R

        # i=1;resultado_R = x[[i]]
        i = .y
        resultado_R = .x

        dono_id = base::unique(resultado_R$dono_id)

        {# Start: verificando o número de linhas

          # Número de linhas
          n_servidor = serie_historica_servidor %>%
            dplyr::filter(dono_id %in% base::unique(resultado_R$dono_id)) %>%
            base::nrow()

          n_R = base::nrow(resultado_R)

          n_linhas_result = n_servidor == n_R

          log_n_linhas_result = NA

          if( n_linhas_result == FALSE )
          {# Start: se erro, verificar os indice_id's com erro

            indice_id_servidor = serie_historica_servidor %>%
              dplyr::filter(dono_id %in% base::unique(resultado_R$dono_id)) %>%
              dplyr::select(indice_id) %>%
              base::unique() %>%
              dplyr::pull()

            indice_id_R = resultado_R %>%
              dplyr::select(indice_id) %>%
              dplyr::pull()

            no_servidor_n_R = indice_id_servidor[base::which(indice_id_servidor %nin% indice_id_R)]
            no_n_servidor = indice_id_R[base::which(indice_id_R %nin% indice_id_servidor)]

            log_n_linhas_result = base::paste0(
              "indice_id = (",
              base::paste0(
                no_servidor_n_R, no_n_servidor, collapse = ","
              ),
              ")"
            )

          }# End: se erro, verificar os indice_id's com erro

          erro_n_linhas = tibble::tibble(
            dono_id = dono_id,
            n_servidor = n_servidor,
            n_R = n_R,
            n_linhas_result = n_linhas_result,
            log = log_n_linhas_result
          )

        }# End: verificando o número de linhas

        {# Start: amostra vs amostra pretendida

          erro_amostra_pretend = tibble::tibble(
            dono_id = dono_id,
            amostra = resultado_R$amostra,
            amostra_prent = resultado_R$amostra_prent,
            igual = resultado_R$amostra == resultado_R$amostra_prent
          )

          resultado_R$amostra

        }# End: amostra vs amostra pretendida

        comparando = dplyr::left_join(
          resultado_R %>%
            dplyr::select(-pasta_nome)
          ,serie_historica_servidor %>%
            dplyr::select(indice_id, dono_id, seriehistorica_media, seriehistorica_indice, seriehistorica_importancia) %>%
            dplyr::rename(
              media_servidor = seriehistorica_media,
              indice_servidor = seriehistorica_indice,
              importancia_servidor = seriehistorica_importancia,
            )
          ,by = c(
            "indice_id" = "indice_id",
            "dono_id" = "dono_id"
          )
        ) %>%
          dplyr::relocate(indice_servidor,.after =  indice) %>%
          dplyr::relocate(media_servidor,.after =  media) %>%
          dplyr::relocate(importancia_servidor,.after =  seriehistorica_importancia) %>%
          dplyr::mutate(
            dplyr::across(
              .cols=c("media", 'indice', 'seriehistorica_importancia', 'indice_servidor', 'media_servidor', 'importancia_servidor'),
              ~IPpackage::round_excel(base::as.numeric(.),4)
            )
          ) %>%
          dplyr::mutate(
            i = (indice == indice_servidor),
            m = (media == media_servidor),
            imp = (seriehistorica_importancia == importancia_servidor)
          )%>%
          dplyr::left_join(
            split %>%
              dplyr::select(dono_id,split_nome,pasta_nome),
            by = c("dono_id" = "dono_id"),
            relationship = "many-to-many"
          ) %>%
          base::unique()

        erro_n_linhas = erro_n_linhas %>%
          dplyr::filter(n_linhas_result == FALSE)

        erro_amostra_pretend = erro_amostra_pretend %>%
          dplyr::filter(igual == FALSE)

        erro_i = comparando %>%
          dplyr::filter(i == FALSE) %>%
          dplyr::select(
            indice_sigla, indice_id, dono_id, split_nomes, indice, indice_servidor,
            media, media_servidor, seriehistorica_importancia, importancia_servidor,
            amostra, amostra_prent, a_veri, pasta_nome
            ) %>%
          dplyr::mutate(dif = base::abs(indice - indice_servidor))

        erro_m = comparando %>%
          dplyr::filter(m == FALSE) %>%
          dplyr::select(
            indice_sigla, indice_id, dono_id, split_nomes, indice, indice_servidor,
            media, media_servidor, seriehistorica_importancia, importancia_servidor,
            amostra, amostra_prent, a_veri, pasta_nome
            ) %>%
          dplyr::mutate(dif = base::abs(media - media_servidor))

        erro_imp = comparando %>%
          dplyr::filter(imp == FALSE) %>%
          dplyr::select(
            indice_sigla, indice_id, dono_id, split_nomes, indice, indice_servidor,
            media, media_servidor, seriehistorica_importancia, importancia_servidor,
            amostra, amostra_prent, a_veri, pasta_nome
            ) %>%
          dplyr::mutate(dif = base::abs(seriehistorica_importancia - importancia_servidor))

        # Start: Salvar no arquivo temporário (não armazenar no MiB do R)

        base::saveRDS(
          object = erro_n_linhas,
          file = base::paste0(caminho_temporario_erro_n_linhas, "/", dono_id, ".rds")
        )

        base::saveRDS(
          object = erro_amostra_pretend,
          file = base::paste0(caminho_temporario_erro_amostra_pretend, "/", dono_id, ".rds")
        )

        base::saveRDS(
          object = erro_i,
          file = base::paste0(caminho_temporario_erro_i, "/", dono_id, ".rds")
        )
        base::saveRDS(
          object = erro_m,
          file = base::paste0(caminho_temporario_erro_m, "/", dono_id, ".rds")
        )
        base::saveRDS(
          object = erro_imp,
          file = base::paste0(caminho_temporario_erro_imp, "/", dono_id, ".rds")
        )

        base::print(base::paste0("dono_id = ", dono_id))

        base::rm(erro_imp, erro_i, erro_m, comparando) %>%
          base::suppressWarnings()

      }# End: verificando cada um dos dono_id calculados no R

    )

  {# Start: salvando o resultado

    erro_n_linhas = base::do.call(
      base::rbind,
      base::lapply(
        base::list.files(
          path = caminho_temporario_erro_n_linhas,
          pattern = "\\.rds$",
          full.names = TRUE
        ),
        readRDS
      )
    )

    erro_amostra_pretend = base::do.call(
      base::rbind,
      base::lapply(
        base::list.files(
          path = caminho_temporario_erro_amostra_pretend,
          pattern = "\\.rds$",
          full.names = TRUE
        ),
        readRDS
      )
    )
    erro_i = base::do.call(
      base::rbind,
      base::lapply(
        base::list.files(
          path = caminho_temporario_erro_i,
          pattern = "\\.rds$",
          full.names = TRUE
        ),
        readRDS
      )
    )

    erro_m = base::do.call(
      base::rbind,
      base::lapply(
        base::list.files(
          path = caminho_temporario_erro_m,
          pattern = "\\.rds$",
          full.names = TRUE
        ),
        readRDS
      )
    )

    erro_imp = base::do.call(
      base::rbind,
      base::lapply(
        base::list.files(
          path = caminho_temporario_erro_imp,
          pattern = "\\.rds$",
          full.names = TRUE
        ),
        readRDS
      )
    )

    base::dir.create(caminho_resultado) %>%
      base::suppressWarnings()

    base::saveRDS(
      object = base::do.call(
        base::rbind,
        base::lapply(
          base::list.files(
            path = caminho_temporario_erro_n_linhas,
            pattern = "\\.rds$",
            full.names = TRUE
          ),
          readRDS
        )
      ) %>%
        base::unique(),
      file = base::paste0(caminho_resultado, "/erro_n_linhas.rds")
    )

    base::saveRDS(
      object = base::do.call(
        base::rbind,
        base::lapply(
          base::list.files(
            path = caminho_temporario_erro_amostra_pretend,
            pattern = "\\.rds$",
            full.names = TRUE
          ),
          readRDS
        )
      ) %>%
        base::unique(),
      file = base::paste0(caminho_resultado, "/erro_amostra_pretend.rds")
    )

    base::saveRDS(
      object = base::do.call(
        base::rbind,
        base::lapply(
          base::list.files(
            path = caminho_temporario_erro_i,
            pattern = "\\.rds$",
            full.names = TRUE
          ),
          readRDS
        )
      ),
      file = base::paste0(caminho_resultado, "/erro_i.rds")
    )

    base::saveRDS(
      object = base::do.call(
        base::rbind,
        base::lapply(
          base::list.files(
            path = caminho_temporario_erro_m,
            pattern = "\\.rds$",
            full.names = TRUE
          ),
          readRDS
        )
      ),
      file = base::paste0(caminho_resultado, "/erro_m.rds")
    )

    base::saveRDS(
      object = base::do.call(
        base::rbind,
        base::lapply(
          base::list.files(
            path = caminho_temporario_erro_imp,
            pattern = "\\.rds$",
            full.names = TRUE
          ),
          readRDS
        )
      ),
      file = base::paste0(caminho_resultado, "/erro_imp.rds")
    )

    IPpackage::print_cor(
      df = "\n----------------------------------------------------------------\n",
      cor = "ciano",
      data.frame = FALSE
    )

    IPpackage::print_cor(
      df = base::paste0("Arquivos tempor\u00E1rios exclu\u00EDdos e o final salvo em\n '", caminho_resultado, "\n"),
      cor = "verde",
      negrito = TRUE,
      data.frame = FALSE
    )

    base::unlink(
      x = base::paste0(caminho_temporario),
      recursive = TRUE
    )

    }# End: salvando o resultado

}# End: Função 'IndiceVerificandoIndices'
