#' @title IndicePlanilhaEscalaAberta
#' @description Gera uma planilha formatada com uma escala de satisfação, exibindo dados numéricos e de texto, além de um cabeçalho estilizado. A função monta uma aba nova com organização e estilo específicos de acordo com os parâmetros fornecidos.
#'
#' @param wb Workbook onde a aba será criada.
#' @param logo_innovare Caminho para a imagem do logotipo a ser inserido na planilha.
#' @param nome_aba Nome da aba a ser adicionada ao workbook.
#' @param pasta_nome Texto para indicar a pasta de dados no cabeçalho.
#' @param manter_ao_remover_dados Colunas a serem preservadas ao remover dados de índices específicos.
#' @param remover_dos_indices Vetor de índices que devem ser removidos da planilha.
#' @param indice_sigla_preco Sigla do índice de preço para identificar e estilizar linhas específicas.
#' @param nota_roda_pe Vetor de notas de rodapé opcionais.
#' @param espanhol Lógico. Se TRUE, exibe cabeçalho e títulos em espanhol.
#' @param anos Vetor contendo os anos a serem exibidos nos cabeçalhos das colunas.
#' @param dados Data frame com os dados que serão exibidos na planilha.
#' @param Tab_Fundo_cinza Data frame com informações de cores para formatar linhas específicas.
#' @param referencia_filtrado Data frame que define as cores de fundo e texto para cada índice.
#' @param print Lógico. Se TRUE, abre a planilha após criação.
#'
#' @details Esta função cria uma nova aba no workbook `wb` utilizando `openxlsx`. O layout é estruturado para exibir uma escala aberta de satisfação (5 pontos) e aplica estilos visuais avançados, incluindo coloração por indicadores e mesclagem de células.
#' Cabeçalhos e títulos são configurados com base no idioma selecionado e no vetor de anos. Estilos adicionais são aplicados a cada célula com base nas cores dos indicadores e configurações de alinhamento e formatação numérica.
#'
#' A função oferece flexibilidade para ajustar colunas específicas a serem mantidas ou removidas, personalizar o cabeçalho em diferentes idiomas e inserir logotipos. Notas de rodapé opcionais também são suportadas para exibição adicional abaixo da tabela.
#'
#' @return O workbook `wb` atualizado com a aba recém-criada e formatada.
#'
#' @import openxlsx
#' @import dplyr
#' @importFrom tibble tibble
#' @importFrom tidyr drop_na
#' @importFrom stringr str_detect
#'
#' @examples
#'
#' base::print("Sem Exemplo")
#'
#' @export


IndicePlanilhaEscalaAberta <- function(
    wb,
    logo_innovare,
    nome_aba,
    pasta_nome,
    manter_ao_remover_dados = NULL,
    remover_dos_indices = NULL,
    indice_sigla_preco = "PR1",
    nota_roda_pe = NULL, # NULL
    espanhol = FALSE,
    anos,
    dados,
    Tab_Fundo_cinza,
    referencia_filtrado,
    print = FALSE
)
{# Start: função 'IndicePlanilhaSatisfacao'

  {# Start: Configuração inicial da tabela e formatação dos dados

    # Definindo o início da tabela numérica
    inicio_tabela_numerico = 13 - 1

    # Identificando linhas onde o valor de IDAR é NA para realizar a quebra de linhas na tabela
    tabela_pulou_linha = base::which(base::is.na(dados$`IDAR`))

    # Ajustando as quebras de linha, somando o índice inicial da tabela
    tabela_pulou_linha = c(inicio_tabela_numerico, tabela_pulou_linha + inicio_tabela_numerico)

    # Definindo linhas que devem estar em negrito
    negrito = dados %>%
      dplyr::mutate(
        linha_negrito = 1:base::nrow(.) + inicio_tabela_numerico
      ) %>%
      dplyr::filter(linhas_negrito == TRUE) %>%
      dplyr::select(linha_negrito) %>%
      dplyr::pull()

    # Separando a tabela apenas com dados numéricos e convertendo para tipo numérico
    rodando_de_fato = dados %>%
      dplyr::select(-c(
        base::colnames(
          referencia_filtrado %>%
            dplyr::select(-c(titulo, titulo2, referencia_id))
        )
      )
      ) %>%
      base::colnames()

    tabela_dados_numericos = dados %>%
      dplyr::select(rodando_de_fato) %>%
      dplyr::select(-c(1)) %>%
      dplyr::mutate(
        dplyr::across(
          .cols = dplyr::everything(),
          ~ base::as.numeric(.)
        )
      )

    if( !base::is.null(manter_ao_remover_dados) | !base::is.null(remover_dos_indices) )
    {# Start: Removendo valores para colunas específicas na tabela numérica

      if( !base::is.null(manter_ao_remover_dados) & !base::is.null(remover_dos_indices) )
      {# Start: Se ambos são null, fazer

        # Removendo valores para colunas específicas na tabela numérica
        tabela_dados_numericos[
          base::which(dados$`IDAR` %in% remover_dos_indices),
          -base::which(base::colnames(tabela_dados_numericos) %in% manter_ao_remover_dados)
        ] <- NA

      } else {# End: Se ambos são null, fazer
        # Se ambos não são null's, erro

        base::stop("'manter_ao_remover_dados' e 'remover_dos_indices'\nOu ambos os par\u00E2metros s\u00E3o NULL ou ambos n\u00E3o s\u00E3o NULL")

      }

    }# End: Removendo valores para colunas específicas na tabela numérica

    # Criando tabela apenas com os dados de texto e convertendo para tipo caractere
    tabela_dados_texto = dados[, c(1, 2)] %>%
      dplyr::mutate(
        dplyr::across(
          .cols = dplyr::everything(),
          ~ base::as.character(.)
        )
      )

  } # End: Configuração inicial da tabela e formatação dos dados

  {# Start: Montando a planilha

    # Removendo a planilha anterior se já existir no workbook
    if( base::any(base::names(wb) == nome_aba) )
    {# Start: Removendo a planilha anterior se já existir no workbook

      openxlsx::removeWorksheet(
        wb = wb,
        sheet = nome_aba
      )

    }# End: Removendo a planilha anterior se já existir no workbook

    #wb <- createWorkbook(nome_aba)

    # Adicionando uma nova aba ao workbook para receber os dados
    openxlsx::addWorksheet(
      wb = wb,
      sheetName = nome_aba
    )

    # Escrevendo dados de texto na nova aba
    openxlsx::writeData(
      wb = wb,
      sheet = nome_aba,
      x = tabela_dados_texto %>%
        dplyr::mutate(
          IDAR = base::ifelse(IDAR == indice_sigla_preco, "PR", IDAR)
        ),
      startCol = 1,
      startRow = inicio_tabela_numerico + 1,
      rowNames = FALSE,
      colNames = FALSE
    )

    # Escrevendo dados numéricos na nova aba
    openxlsx::writeData(
      wb = wb,
      sheet = nome_aba,
      x = tabela_dados_numericos,
      startCol = 3,
      startRow = inicio_tabela_numerico + 1,
      rowNames = FALSE,
      colNames = FALSE
    )

    # Colocando o logo na aba
    openxlsx::insertImage(
      wb = wb,
      sheet = nome_aba,
      file = logo_innovare,
      startRow = 2,
      startCol = 2,
      height = (2.43 * 2.43) / 6.17,
      width = (4.5 * 4.5) / 11.43
    )

    {# Start: Ajustando alturas específicas das linhas

      # Ajustando altura das linhas de 2 a 5 para melhorar a apresentação
      openxlsx::setRowHeights(
        wb = wb,
        sheet = nome_aba,
        rows = c(2:5),
        heights = c(21)
      )

      # Ajustando altura das linhas 6 a 8, incluindo quebras de linha identificadas na tabela
      openxlsx::setRowHeights(
        wb = wb,
        sheet = nome_aba,
        rows = c(6:8, tabela_pulou_linha),
        heights = c(3.75)
      )

    }# End: Ajustando alturas específicas das linhas

    {# Start: Ajustando larguras específicas das colunas

      # Definindo largura da primeira coluna para exibir bem os dados
      openxlsx::setColWidths(
        wb = wb,
        sheet = nome_aba,
        cols = 1,
        widths = 6
      )

      # Definindo largura da segunda coluna para títulos mais longos
      openxlsx::setColWidths(
        wb = wb,
        sheet = nome_aba,
        cols = 2,
        widths = 64
      )

      # Definindo largura para as colunas restantes com base na quantidade de colunas em 'dados'
      openxlsx::setColWidths(
        wb = wb,
        sheet = nome_aba,
        cols = 3:ncol(dados),
        widths = 8.43
      )

    }# End: Ajustando larguras específicas das colunas

    {# Start: Escrevendo o cabeçalho do menu na aba e aplicando estilo

      # Escrevendo o título "PLANILHA DE ÍNDICES" na célula específica
      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = base::ifelse(espanhol == FALSE, "PLANILHA DE \u00CDNDICES", "PLANILLA DE \u00CDNDICES"),
        startCol = 2,
        startRow = 3,
        rowNames = FALSE,
        colNames = FALSE
      )

      # Inserindo o nome da pasta na célula abaixo do título
      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = pasta_nome,
        startCol = 2,
        startRow = 4,
        rowNames = FALSE,
        colNames = FALSE
      )

      # Aplicando estilo ao título e ao nome da pasta: fonte 16, negrito, alinhamento à direita
      openxlsx::addStyle(
        wb = wb,
        sheet = nome_aba,
        style = openxlsx::createStyle(
          fontSize = 16,
          fontColour = "black",
          textDecoration = c("BOLD"),
          halign = "right",
          valign = "center"
        ),
        rows = 3:4,
        cols = 2,
        gridExpand = TRUE
      )

    }# End: Escrevendo o cabeçalho do menu na aba e aplicando estilo

    {# Start: Configurando cabeçalhos da tabela e mesclando células para estrutura visual

      # Extraindo os anos presentes nos nomes das colunas que contêm "-"
      anos_rodando = rodando_de_fato

      # Filtrando apenas os anos identificados com o padrão "-", extraindo a parte antes do "-"
      anos_rodando = anos_rodando %>%
        stringr::str_detect("-") %>%
        base::which()

      anos_rodando = rodando_de_fato[anos_rodando] %>%
        base::sub("-.*", "", .) %>%
        base::unique()

      # Escrevendo o título "Atributos" na célula específica
      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = "Atributos",
        startCol = 2,
        startRow = 9,
        rowNames = FALSE,
        colNames = FALSE
      )

      # Removendo mesclagem anterior e aplicando mesclagem de células para o título "Atributos"
      openxlsx::removeCellMerge(
        wb = wb,
        sheet = nome_aba,
        cols = 2,
        rows = 9:11
      )

      openxlsx::mergeCells(
        wb = wb,
        sheet = nome_aba,
        cols = 2,
        rows = 9:11
      )

      # Aplicando estilo ao título "Atributos" com cor de fundo e texto em branco, negrito e centralizado
      openxlsx::addStyle(
        wb = wb,
        sheet = nome_aba,
        style = openxlsx::createStyle(
          fontSize = 10,
          fontColour = "white",
          fgFill = "#404040",
          textDecoration = c("BOLD"),
          halign = "center",
          valign = "center"
        ),
        rows = 9,
        cols = c(2),
        gridExpand = TRUE
      )

      # Iterando sobre cada ano em "anos_rodando" para configurar as colunas e cabeçalhos
      x = base::list(c(3, 9))
      i = 1
      for( i in 1:length(anos_rodando) )
      { # Start: Configuração de colunas e cabeçalhos para cada ano

        # Ajuste da posição inicial de cada ano na planilha
        if( i == 1 )
        { # Start: Definindo posição inicial para o primeiro ano

          x[[i]] = x[[i]]

        } # End: Definindo posição inicial para o primeiro ano
        else
        { # Start: Ajustando posição para cada ano subsequente

          x[[i]] = x[[i - 1]] + 7

        } # End: Ajustando posição para cada ano subsequente

        # Linha 9: Escrevendo o ano e aplicando estilo e mesclagem
        { # Start: Configuração de dados e estilo para a linha 9 (ano)

          openxlsx::writeData(
            wb = wb,
            sheet = nome_aba,
            x = as.numeric(anos_rodando[i]),
            startCol = x[[i]][[1]],
            startRow = 9,
            rowNames = FALSE,
            colNames = FALSE
          )

          openxlsx::removeCellMerge(
            wb = wb,
            sheet = nome_aba,
            cols = x[[i]][[1]]:x[[i]][[2]],
            rows = 9
          )

          openxlsx::mergeCells(
            wb = wb,
            sheet = nome_aba,
            cols = x[[i]][[1]]:x[[i]][[2]],
            rows = 9
          )

          openxlsx::addStyle(
            wb = wb,
            sheet = nome_aba,
            style = openxlsx::createStyle(
              fontSize = 10,
              fontColour = "white",
              fgFill = "#404040",
              textDecoration = c("BOLD"),
              halign = "center",
              valign = "center",
              numFmt = "0"
            ),
            rows = 9,
            cols = x[[i]][[1]],
            gridExpand = TRUE
          )

        } # End: Configuração de dados e estilo para a linha 9 (ano)

        # Linha 10: Escrevendo "Escala Agregada", "NS/NR" e "Média" com mesclagem e estilo
        { # Start: Configuração de dados e estilo para a linha 10

          openxlsx::writeData(
            wb = wb,
            sheet = nome_aba,
            x = base::ifelse(espanhol == FALSE, "Escala Agregada (5 pontos) - %", "Tabla agregada (5 puntos) - %"),
            startCol = x[[i]][[1]],
            startRow = 10,
            rowNames = FALSE,
            colNames = FALSE
          )

          openxlsx::removeCellMerge(
            wb = wb,
            sheet = nome_aba,
            cols = x[[i]][[1]]:(x[[i]][[2]] - 2),
            rows = 10
          )

          openxlsx::mergeCells(
            wb = wb,
            sheet = nome_aba,
            cols = x[[i]][[1]]:(x[[i]][[2]] - 2),
            rows = 10
          )

          openxlsx::writeData(
            wb = wb,
            sheet = nome_aba,
            x = "NS/NR",
            startCol = (x[[i]][[2]] - 1),
            startRow = 10,
            rowNames = FALSE,
            colNames = FALSE
          )

          openxlsx::writeData(
            wb = wb,
            sheet = nome_aba,
            x = base::ifelse(espanhol == FALSE, "M\u00E9dia", "Promedio"),
            startCol = x[[i]][[2]],
            startRow = 10,
            rowNames = FALSE,
            colNames = FALSE
          )

          openxlsx::removeCellMerge(
            wb = wb,
            sheet = nome_aba,
            cols = x[[i]][[2]],
            rows = 10:11
          )

          openxlsx::mergeCells(
            wb = wb,
            sheet = nome_aba,
            cols = x[[i]][[2]],
            rows = 10:11
          )

          openxlsx::addStyle(
            wb = wb,
            sheet = nome_aba,
            style = openxlsx::createStyle(
              fontSize = 10,
              fontColour = "white",
              fgFill = "#404040",
              textDecoration = c("BOLD"),
              halign = "center",
              valign = "center",
              numFmt = "0"
            ),
            rows = 10,
            cols = c(x[[i]][[1]], (x[[i]][[2]] - 1), x[[i]][[2]]),
            gridExpand = TRUE
          )

        } # End: Configuração de dados e estilo para a linha 10

        # Linha 11: Escrevendo categorias de escala "1+2", "3+4", etc., com estilo colorido
        { # Start: Configuração de dados e estilo para a linha 11

          categorias = base::list(
            "1+2" = "#A71E22",
            "3+4" = "#E41D32",
            "5+6" = "#F7CE1B",
            "7+8" = "#94CA53",
            "9+10" = "#17AE6B",
            "11+12" = "#808080"
          )

          offset = 0
          for( categoria in base::names(categorias) )
          {

            cor = categorias[[categoria]]
            openxlsx::writeData(
              wb = wb,
              sheet = nome_aba,
              x = categoria,
              startCol = (x[[i]][[1]] + offset),
              startRow = 11,
              rowNames = FALSE,
              colNames = FALSE
            )

            openxlsx::addStyle(
              wb = wb,
              sheet = nome_aba,
              style = openxlsx::createStyle(
                fontSize = 10,
                fontColour = "white",
                fgFill = cor,
                textDecoration = c("BOLD"),
                halign = "center",
                valign = "center",
                numFmt = "0"
              ),
              rows = 11,
              cols = (x[[i]][[1]] + offset),
              gridExpand = TRUE
            )
            offset = offset + 1

          }

        } # End: Configuração de dados e estilo para a linha 11

      } # End: Configuração de colunas e cabeçalhos para cada ano

    }# End: Configurando cabeçalhos da tabela e mesclando células para estrutura visual

    {# Start: Formatando a tabela (texto e números) com base nas cores dos indicadores

      # Obtendo lista única de cores de fundo dos indicadores
      all_cores = referencia_filtrado %>%
        dplyr::select(cor_fundo) %>%
        base::unique() %>%
        tidyr::drop_na() %>%
        dplyr::pull()

      for( i in 1:length(all_cores) )
      {# Start: Iterando sobre cada cor para aplicar estilos aos respectivos índices

        # Identificando quais indicadores possuem a cor atual
        indicadores_rodando = referencia_filtrado %>%
          dplyr::filter(cor_fundo == all_cores[i]) %>%
          dplyr::select(IDAR) %>%
          dplyr::pull() %>%
          base::unique()

        # Definindo 'PR'
        indicadores_rodando_procurar = base::ifelse(
          indicadores_rodando == indice_sigla_preco,
          "PR",
          indicadores_rodando
        )

        indicadores_linhas = base::which(
          tabela_dados_texto$IDAR %in% c(indicadores_rodando_procurar)
        ) + inicio_tabela_numerico

        # Definindo a cor da letra para esses indicadores
        indicador_cor = referencia_filtrado %>%
          dplyr::filter(cor_fundo == all_cores[i]) %>%
          dplyr::select(cor_letra) %>%
          dplyr::slice(1) %>%
          dplyr::pull()

        if( base::length(indicadores_linhas) > 0 )
        {# Start: Pelo menos um índice para essa cor

          {# Start: separando pelas quebras de sequência

            separado = base::list()
            nlistas = 1
            w = 1
            separado[[nlistas]] = c(indicadores_linhas[w])

            for( w in 1:c(base::length(indicadores_linhas)-1) )
            {# Start: separando pelas quebras de sequência

              if( base::length(indicadores_linhas) > 1 )
              {# Start: se tenho mais de um índice para olhar

                ww = indicadores_linhas[w + 1] == indicadores_linhas[w] + 1

                if( ww == TRUE )
                {

                  separado[[nlistas]] = base::unique(c(separado[[nlistas]], indicadores_linhas[w + 1]))

                } else {

                  nlistas = nlistas + 1

                  separado[[nlistas]] = base::unique(c(indicadores_linhas[w + 1]))

                }

              }

            }# End: separando pelas quebras de sequência

            base::rm(w, nlistas, indicadores_linhas) %>%
              base::suppressWarnings()

          }# End: separando pelas quebras de sequência

          for( w in 1:base::length(separado) )
          {# Start: rodando para cada sequência de seguidos

            indicadores_linhas = separado[[w]]

            # if( !base::any(indicadores_rodando %in% c(indice_nao_mesclar)) )
            {# Start: Mesclando células apenas para índices específicos

              openxlsx::removeCellMerge(
                wb = wb,
                sheet = nome_aba,
                cols = 1,
                rows = indicadores_linhas
              )

              openxlsx::mergeCells(
                wb = wb,
                sheet = nome_aba,
                cols = 1,
                rows = indicadores_linhas
              )

            }# End: Mesclando células apenas para índices específicos

            # Aplicando estilo para a coluna de índice com a cor de fundo e cor da letra
            openxlsx::addStyle(
              wb = wb,
              sheet = nome_aba,
              style = openxlsx::createStyle(
                fontSize = 10,
                fontColour = indicador_cor,
                fgFill = all_cores[i],
                textDecoration = c("BOLD"),
                halign = "center",
                valign = "center"
              ),
              rows = indicadores_linhas,
              cols = 1,
              gridExpand = TRUE
            )

            for( j in 1:length(indicadores_linhas) )
            {# Start: Aplicando estilo em negrito para linhas específicas

              # Determinando cor de fundo: cinza ou branco
              fund_cinza = Tab_Fundo_cinza %>%
                dplyr::filter(IDAR == dados$IDAR[indicadores_linhas[j] - inicio_tabela_numerico]) %>%
                dplyr::select(fundo_cinza) %>%
                dplyr::pull()

              fundo_cor = base::ifelse(fund_cinza == TRUE, "#F2F2F2", "white")

              # Determinando se a fonte será em negrito
              letra_negrito = NULL
              if( indicadores_linhas[j] %in% c(negrito) )
              {
                letra_negrito = "BOLD"
              }

              # Aplicando estilo para a célula de texto na segunda coluna
              openxlsx::addStyle(
                wb = wb,
                sheet = nome_aba,
                style = openxlsx::createStyle(
                  fontSize = 10,
                  fontColour = "black",
                  fgFill = fundo_cor,
                  textDecoration = letra_negrito,
                  halign = "left",
                  valign = "center",
                  wrapText = TRUE
                ),
                rows = indicadores_linhas[j],
                cols = 2,
                gridExpand = FALSE
              ) %>% base::suppressWarnings()

              for( k in 3:c(base::length(rodando_de_fato) + 1) )
              {# Start: Preenchendo células vazias com "-" e aplicando estilo numérico

                linha_rodando = indicadores_linhas[j] - inicio_tabela_numerico
                coluna_rodando = k - 2

                if( !base::is.na(dados[linha_rodando, 1]) )
                {

                  if( base::is.na(tabela_dados_numericos[linha_rodando, coluna_rodando]) )
                  {

                    openxlsx::writeData(
                      wb = wb,
                      sheet = nome_aba,
                      x = "-",
                      startCol = k,
                      startRow = indicadores_linhas[j],
                      rowNames = FALSE,
                      colNames = FALSE
                    )
                  }

                }

              }# End: Preenchendo células vazias com "-" e aplicando estilo numérico

              # Formatando linhas da tabela com valores numéricos
              openxlsx::addStyle(
                wb = wb,
                sheet = nome_aba,
                style = openxlsx::createStyle(
                  fontSize = 10,
                  fontColour = "black",
                  fgFill = fundo_cor,
                  textDecoration = letra_negrito,
                  halign = "center",
                  valign = "center",
                  wrapText = TRUE,
                  numFmt = "0.0"
                ),
                rows = indicadores_linhas[j],
                cols = c(3:c(base::length(rodando_de_fato) + 1)),
                gridExpand = FALSE
              ) %>%
                base::suppressWarnings()

              # Removendo variáveis temporárias
              base::rm(fund_cinza, fundo_cor, letra_negrito) %>%
                base::suppressWarnings()
            }# End: Aplicando estilo em negrito para linhas específicas

            # Limpando variável temporária do loop
            base::rm(j)

          }# End: rodando para cada sequência de seguidos

        }# End: Pelo menos um índice para essa cor

      }# End: Iterando sobre cada cor para aplicar estilos aos respectivos índices


      # Removendo variável temporária de índice
      base::rm(i)

    }# End: Formatando a tabela (texto e números) com base nas cores dos indicadores

    if( !base::is.null(nota_roda_pe) )
    {# Start: Nota de roda pé

      for( w in 1:base::length(nota_roda_pe) )
      {# Start: aplicando cada 'nota_roda_pe'

        openxlsx::writeData(
          wb = wb,
          sheet = nome_aba,
          x = nota_roda_pe[w],
          startCol  =  1,
          startRow  =  inicio_tabela_numerico + base::nrow(dados) + w + 1,
          rowNames  =  FALSE,
          colNames  =  FALSE
        )

        openxlsx::addStyle(
          wb,
          nome_aba,
          openxlsx::createStyle(
            fontSize  =  10,
            fontColour  =  "black",
            textDecoration  =  c("BOLD"),
            halign = "left",
            valign = "center"
          ),
          rows  =  inicio_tabela_numerico + base::nrow(dados) + w + 1,
          cols  =  1,
          gridExpand  =  TRUE
        )


      }# End: aplicando cada 'nota_roda_pe'

    }# End: Nota de roda pé

    # Removendo linhas de grade da planilha
    openxlsx::showGridLines(
      wb = wb,
      sheet = nome_aba,
      showGridLines = FALSE
    )

    # Congelando a primeira linha ativa da planilha para facilitar a visualização
    openxlsx::freezePane(
      wb = wb,
      sheet = nome_aba,
      firstActiveRow = 12,
      firstActiveCol = 3
    )

  }# End: Montando a planilha

  if( print == TRUE )
  {# Start:Abrindo a planilha se o parâmetro 'print' for TRUE

    openxlsx::openXL(wb)

  }# End: Abrindo a planilha se o parâmetro 'print' for TRUE

  # Retornando o workbook com todas as modificações
  base::return(wb)

}# End: função 'IndicePlanilhaSatisfacao'
