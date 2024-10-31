#' @title IndicePlanilhaSatisfacao
#' @description Gera uma planilha de satisfação formatada com cabeçalhos, estilos e dados, incluindo
#' logotipo, coloração, e agrupamento de células para indicadores específicos.
#'
#' @param wb Workbook aberto onde a aba será criada.
#' @param logo_innovare Caminho para o arquivo de imagem do logotipo a ser inserido na aba.
#' @param nome_aba Nome da aba a ser criada no workbook.
#' @param pasta_nome Nome a ser exibido na planilha, indicando a pasta de dados usada.
#' @param manter_ao_remover_dados Colunas que permanecerão quando dados de determinados índices forem removidos.
#' @param remover_dos_indices Vetor com as siglas dos índices que devem ser excluídos dos dados.
#' @param indice_sigla_preco Sigla do índice de preço, usada para identificar linhas específicas.
#' @param nota_roda_pe Vetor com notas de rodapé opcionais a serem exibidas abaixo da tabela.
#' @param ano_titulo_planilha Ano ou título para a planilha, exibido no cabeçalho.
#' @param espanhol Lógico. Quando TRUE, exibe o cabeçalho em espanhol.
#' @param dados Data frame contendo os dados a serem exibidos na planilha.
#' @param Tab_Fundo_cinza Data frame para definir linhas que terão fundo cinza na tabela.
#' @param referencia_filtrado Data frame que contém as cores de fundo e texto para cada índice.
#' @param print Lógico. Quando TRUE, a planilha é aberta automaticamente após a criação.
#'
#' @details Esta função organiza e estiliza dados de satisfação em um workbook `wb` usando o pacote `openxlsx`.
#' Ela cria uma nova aba com cabeçalhos customizados e aplica formatação avançada de acordo com as especificações dos índices.
#' Estilos específicos são aplicados a células de acordo com as cores de cada índice,
#' as células são mescladas em determinadas posições, e uma imagem de logotipo é adicionada à planilha.
#'
#' O código estrutura e formata a planilha de maneira a destacar informações importantes,
#' com ajuste automático das alturas das linhas, e inserção de notas de rodapé, se especificadas.
#'
#' @return Retorna o workbook `wb` com as modificações aplicadas.
#'
#' @import openxlsx
#' @import dplyr
#' @importFrom tibble tibble
#' @importFrom tidyr drop_na
#'
#' @examples
#'
#' base::print("Sem Exemplo")
#'
#' @export

IndicePlanilhaSatisfacao <- function(
    wb,
    logo_innovare,
    nome_aba,
    pasta_nome,
    manter_ao_remover_dados = "\u00CDndice", # NULL
    remover_dos_indices = c("ISQP", "IEQP", "IIQP", "ISCP", "ISC", "IESC", "IIC", "ISCAL", "IECP", "IICP"), #NULL
    indice_sigla_preco = "PR1",
    nota_roda_pe = NULL, # NULL
    ano_titulo_planilha,
    espanhol = FALSE,
    dados,
    Tab_Fundo_cinza,
    referencia_filtrado,
    print = FALSE
)
{# Start: função 'IndicePlanilhaSatisfacao'

  {# Start: Configuração inicial da tabela e formatação dos dados

    # Definindo início da tabela e ajustando linhas que devem pular
    inicio_tabela_numerico = 12 - 1

    tabela_pulou_linha = base::which(base::is.na(dados$`indice_id`))
    tabela_pulou_linha = c(inicio_tabela_numerico, tabela_pulou_linha + inicio_tabela_numerico)

    # Definindo linhas que devem estar em negrito
    negrito = dados %>%
      dplyr::mutate(
        linha_negrito = 1:base::nrow(.) + inicio_tabela_numerico
      ) %>%
      dplyr::filter(linhas_negrito == TRUE) %>%
      dplyr::select(linha_negrito) %>%
      dplyr::pull()

    # Tabela contendo apenas dados numéricos e convertendo para numérico
    tabela_dados_numericos = dados %>%
      dplyr::select(
        "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11+12",
        "Total", "Base", "\u00CDndice", "M\u00E9dia"
      ) %>%
      dplyr::mutate(
        dplyr::across(
          .cols = dplyr::everything(),
          ~ base::as.numeric(.)
        )
      ) %>%
      base::suppressWarnings()

    if( !base::is.null(manter_ao_remover_dados) | !base::is.null(remover_dos_indices) )
    {# Start: Removendo valores para colunas específicas na tabela numérica

      if( !base::is.null(manter_ao_remover_dados) & !base::is.null(remover_dos_indices) )
      {# Start: Se ambos são null, fazer

        # Removendo valores para colunas específicas na tabela numérica
        tabela_dados_numericos[
          base::which(dados$`indice_sigla` %in% remover_dos_indices),
          -base::which(base::colnames(tabela_dados_numericos) %in% manter_ao_remover_dados)
        ] <- NA

      } else {# End: Se ambos são null, fazer
        # Se ambos não são null's, erro

        base::stop("'manter_ao_remover_dados' e 'remover_dos_indices'\nOu ambos os par\u00E2metros s\u00E3o NULL ou ambos n\u00E3o s\u00E3o NULL")

      }

    }# End: Removendo valores para colunas específicas na tabela numérica

    # Tabela contendo apenas dados de texto e convertendo para caractere
    tabela_dados_texto = dados %>%
      dplyr::select(indice_sigla, titulo) %>%
      dplyr::mutate(
        dplyr::across(
          .cols = dplyr::everything(),
          ~ base::as.character(.)
        )
      )

  }# End: Configuração inicial da tabela e formatação dos dados

  {# Start: Montando a planilha

    # Removendo a planilha anterior se já existir no workbook
    if (base::any(base::names(wb) == nome_aba))
    {# Start: Removendo a planilha anterior se já existir no workbook

      openxlsx::removeWorksheet(
        wb = wb,
        sheet = nome_aba
      )

    }# End: Removendo a planilha anterior se já existir no workbook

    #wb <- createWorkbook(nome_aba)

    # Adicionando uma nova aba ao workbook para receber os dados
    openxlsx::addWorksheet(wb = wb, sheetName = nome_aba)

    # Escrevendo dados de texto na nova aba
    openxlsx::writeData(
      wb = wb,
      sheet = nome_aba,
      x = tabela_dados_texto,
      startCol = 1,
      startRow = 12,
      rowNames = FALSE,
      colNames = FALSE
    )

    # Escrevendo dados numéricos na nova aba
    openxlsx::writeData(
      wb = wb,
      sheet = nome_aba,
      x = tabela_dados_numericos,
      startCol = 3,
      startRow = 12,
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

      openxlsx::setRowHeights(
        wb = wb,
        sheet = nome_aba,
        rows = c(1, 9, 10),
        heights = c(12.75)
      )

      openxlsx::setRowHeights(
        wb = wb,
        sheet = nome_aba,
        rows = c(2:5),
        heights = c(21)
      )

      openxlsx::setRowHeights(
        wb = wb,
        sheet = nome_aba,
        rows = c(6:8, tabela_pulou_linha),
        heights = c(3.75)
      )

    }# End: Ajustando alturas específicas das linhas

    {# Start: Ajustando larguras específicas das colunas

      openxlsx::setColWidths(
        wb = wb,
        sheet = nome_aba,
        cols = 1,
        widths = 6
      )

      openxlsx::setColWidths(
        wb = wb,
        sheet = nome_aba,
        cols = 2,
        widths = 64
      )

      openxlsx::setColWidths(
        wb = wb,
        sheet = nome_aba,
        cols = 3:17,
        widths = 8.43
      )

    }# End: Ajustando larguras específicas das colunas

    {# Start: Escrevendo o cabeçalho do menu na aba e aplicando estilo

      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = base::ifelse(espanhol == FALSE, "PLANILHA DE \u00CDNDICES", "PLANILLA DE \u00CDNDICES"),
        startCol = 2,
        startRow = 3,
        rowNames = FALSE,
        colNames = FALSE
      )

      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = pasta_nome,
        startCol = 2,
        startRow = 4,
        rowNames = FALSE,
        colNames = FALSE
      )

      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = ano_titulo_planilha,
        startCol = 2,
        startRow = 5,
        rowNames = FALSE,
        colNames = FALSE
      )

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

      openxlsx::addStyle(
        wb = wb,
        sheet = nome_aba,
        style = openxlsx::createStyle(
          fontSize = 16,
          fontColour = "black",
          textDecoration = c("BOLD"),
          halign = "right",
          valign = "center",
          numFmt = "0"
        ),
        rows = 5,
        cols = 2,
        gridExpand = TRUE
      )

    }# End: Escrevendo o cabeçalho do menu na aba e aplicando estilo

    {# Start: Configurando cabeçalhos da tabela e mesclando células para estrutura visual

      # Escrevendo o título "Atributos" na aba e definindo a posição inicial
      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = "Atributos",
        startCol = 2,
        startRow = 9,
        rowNames = FALSE,
        colNames = FALSE
      )

      # Removendo mesclagem de células anterior e mesclando células na coluna 2 para o título "Atributos"
      openxlsx::removeCellMerge(
        wb = wb,
        sheet = nome_aba,
        cols = 2,
        rows = 9:10
      )

      openxlsx::mergeCells(
        wb = wb,
        sheet = nome_aba,
        cols = 2,
        rows = 9:10
      )

      # Escrevendo o título "Escala Aberta (10 pontos) - %" e ajustando a mesclagem de células para essa seção
      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = base::ifelse(espanhol == FALSE, "Escala Aberta (10 pontos) - %", "Escala Abierta (10 puntos) - %"),
        startCol = 3,
        startRow = 9,
        rowNames = FALSE,
        colNames = FALSE
      )

      # Removendo mesclagem anterior e mesclando células para o título "Escala Aberta (10 pontos) - %"
      openxlsx::removeCellMerge(
        wb = wb,
        sheet = nome_aba,
        cols = 3:12,
        rows = 9
      )

      for( w in 1:10)
      {# Start: Escrevendo cabeçalhos das colunas numéricas de 1 a 10, posicionando cada número em sua respectiva coluna na linha 10

        openxlsx::writeData(
          wb = wb,
          sheet = nome_aba,
          x = w,
          startCol = w + 2,
          startRow = 10,
          rowNames = FALSE,
          colNames = TRUE
        )

      }# End: Escrevendo cabeçalhos das colunas numéricas de 1 a 10, posicionando cada número em sua respectiva coluna na linha 10

      base::rm(w) %>%
        base::suppressWarnings()

      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = 9,
        startCol = 11,
        startRow = 10,
        rowNames = FALSE,
        colNames = TRUE
      )

      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = 10,
        startCol = 12,
        startRow = 10,
        rowNames = FALSE,
        colNames = TRUE
      )

      # Mesclando células para os cabeçalhos da escala de 1 a 10
      openxlsx::mergeCells(
        wb = wb,
        sheet = nome_aba,
        cols = 3:12,
        rows = 9
      )

      # Escrevendo o cabeçalho "NS/NR" na coluna 13
      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = "NS/NR",
        startCol = 13,
        startRow = 9,
        rowNames = FALSE,
        colNames = FALSE
      )

      # Escrevendo "11+12" como título na coluna 13, linha 10
      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = "11+12",
        startCol = 13,
        startRow = 10,
        rowNames = FALSE,
        colNames = FALSE
      )

      # Escrevendo o cabeçalho "Total" e ajustando mesclagem na coluna 14
      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = "Total",
        startCol = 14,
        startRow = 9,
        rowNames = FALSE,
        colNames = FALSE
      )

      openxlsx::removeCellMerge(
        wb = wb,
        sheet = nome_aba,
        cols = 14,
        rows = 9:10
      )

      openxlsx::mergeCells(
        wb = wb,
        sheet = nome_aba,
        cols = 14,
        rows = 9:10
      )

      # Escrevendo o cabeçalho "Base" e aplicando mesclagem de células na coluna 15
      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = "Base",
        startCol = 15,
        startRow = 9,
        rowNames = FALSE,
        colNames = FALSE
      )

      openxlsx::removeCellMerge(
        wb = wb,
        sheet = nome_aba,
        cols = 15,
        rows = 9:10
      )

      openxlsx::mergeCells(
        wb = wb,
        sheet = nome_aba,
        cols = 15,
        rows = 9:10
      )

      # Escrevendo o cabeçalho "Índice" e ajustando a mesclagem na coluna 16
      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = "\u00CDndice",
        startCol = 16,
        startRow = 9,
        rowNames = FALSE,
        colNames = FALSE
      )

      openxlsx::removeCellMerge(
        wb = wb,
        sheet = nome_aba,
        cols = 16,
        rows = 9:10
      )

      openxlsx::mergeCells(
        wb = wb,
        sheet = nome_aba,
        cols = 16,
        rows = 9:10
      )

      # Escrevendo o cabeçalho "Média" e aplicando mesclagem de células na coluna 17
      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = base::ifelse(espanhol == FALSE, "M\u00E9dia", "Promedio"),
        startCol = 17,
        startRow = 9,
        rowNames = FALSE,
        colNames = FALSE
      )

      openxlsx::removeCellMerge(
        wb = wb,
        sheet = nome_aba,
        cols = 17,
        rows = 9:10
      )

      openxlsx::mergeCells(
        wb = wb,
        sheet = nome_aba,
        cols = 17,
        rows = 9:10
      )

    }# End: Configurando cabeçalhos da tabela e mesclando células para estrutura visual

    {# Start: Aplicando estilos aos títulos da tabela com diferentes cores e formatações

      # Estilo para a linha 9, colunas 2, 3 e 13 a 17 (fundo cinza escuro)
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
        cols = c(2, 3, 13:17),
        gridExpand = TRUE
      )

      # Estilo para a linha 10, colunas 3 e 4 (fundo vermelho escuro)
      openxlsx::addStyle(
        wb = wb,
        sheet = nome_aba,
        style = openxlsx::createStyle(
          fontSize = 10,
          fontColour = "white",
          fgFill = "#A71E22",
          textDecoration = c("BOLD"),
          halign = "center",
          valign = "center",
          numFmt = "0"
        ),
        rows = 10,
        cols = c(3, 4),
        gridExpand = TRUE
      )

      # Estilo para a linha 10, colunas 5 e 6 (fundo vermelho claro)
      openxlsx::addStyle(
        wb = wb,
        sheet = nome_aba,
        style = openxlsx::createStyle(
          fontSize = 10,
          fontColour = "white",
          fgFill = "#E41D32",
          textDecoration = c("BOLD"),
          halign = "center",
          valign = "center"
        ),
        rows = 10,
        cols = c(5, 6),
        gridExpand = TRUE
      )

      # Estilo para a linha 10, colunas 7 e 8 (fundo amarelo)
      openxlsx::addStyle(
        wb = wb,
        sheet = nome_aba,
        style = openxlsx::createStyle(
          fontSize = 10,
          fontColour = "black",
          fgFill = "#F7CE1B",
          textDecoration = c("BOLD"),
          halign = "center",
          valign = "center"
        ),
        rows = 10,
        cols = c(7, 8),
        gridExpand = TRUE
      )

      # Estilo para a linha 10, colunas 9 e 10 (fundo verde claro)
      openxlsx::addStyle(
        wb = wb,
        sheet = nome_aba,
        style = openxlsx::createStyle(
          fontSize = 10,
          fontColour = "white",
          fgFill = "#94CA53",
          textDecoration = c("BOLD"),
          halign = "center",
          valign = "center"
        ),
        rows = 10,
        cols = c(9, 10),
        gridExpand = TRUE
      )

      # Estilo para a linha 10, colunas 11 e 12 (fundo verde escuro)
      openxlsx::addStyle(
        wb = wb,
        sheet = nome_aba,
        style = openxlsx::createStyle(
          fontSize = 10,
          fontColour = "white",
          fgFill = "#17AE6B",
          textDecoration = c("BOLD"),
          halign = "center",
          valign = "center"
        ),
        rows = 10,
        cols = c(11, 12),
        gridExpand = TRUE
      )

      # Estilo para a linha 10, coluna 13 (fundo cinza médio)
      openxlsx::addStyle(
        wb = wb,
        sheet = nome_aba,
        style = openxlsx::createStyle(
          fontSize = 10,
          fontColour = "white",
          fgFill = "#808080",
          textDecoration = c("BOLD"),
          halign = "center",
          valign = "center"
        ),
        rows = 10,
        cols = 13,
        gridExpand = TRUE
      )

    }# End: Aplicando estilos aos títulos da tabela com diferentes cores e formatações

    {# Start: Aplicando formatação na tabela de acordo com as cores e estilos para cada índice

      # Selecionando todas as cores de fundo disponíveis nos índices
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
          tabela_dados_texto$indice_sigla %in% c(indicadores_rodando_procurar)
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
                dplyr::filter(IDAR == dados$indice_sigla[indicadores_linhas[j] - inicio_tabela_numerico]) %>%
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

              for( k in 3:17 )
              {# Start: Preenchendo células vazias com "-" e aplicando estilo numérico
                linha_rodando = indicadores_linhas[j] - inicio_tabela_numerico
                coluna_rodando = k - 2

                if (!base::is.na(dados[linha_rodando, 1])) {
                  if (base::is.na(tabela_dados_numericos[linha_rodando, coluna_rodando])) {
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
                cols = c(3:14, 17),
                gridExpand = FALSE
              ) %>%
                base::suppressWarnings()

              # Aplicando formatação para coluna "Base"
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
                  numFmt = "0"
                ),
                rows = indicadores_linhas[j],
                cols = 15,
                gridExpand = FALSE
              ) %>%
                base::suppressWarnings()

              # Aplicando estilo específico para coluna "Índice"
              openxlsx::addStyle(
                wb = wb,
                sheet = nome_aba,
                style = openxlsx::createStyle(
                  fontSize = 10,
                  fontColour = "white",
                  fgFill = "#848484",
                  textDecoration = letra_negrito,
                  halign = "center",
                  valign = "center",
                  wrapText = TRUE,
                  numFmt = "0.0"
                ),
                rows = indicadores_linhas[j],
                cols = 16,
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

      # Limpando variável temporária do loop principal
      base::rm(i)

      openxlsx::setRowHeights(
        wb = wb,
        sheet = nome_aba,
        rows  =  inicio_tabela_numerico+nrow(dados)+1,
        heights  =  c(3.75)
      )

    }# End: Aplicando formatação na tabela de acordo com as cores e estilos para cada índice

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
      firstActiveRow = 11,
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
