#' @title IndicePlanilhaCorrelacao
#' @description Cria uma aba em um `workbook` com formatação visual específica para apresentar correlações entre atributos e satisfação geral, usando cores, mesclagem de células e outras configurações de estilo.
#'
#' @param wb Workbook para adicionar a nova aba.
#' @param logo_innovare Caminho para o logo a ser inserido.
#' @param nome_aba Nome da aba a ser criada.
#' @param ano_titulo_planilha Ano exibido no título da planilha.
#' @param pasta_nome Nome da pasta para exibição no cabeçalho.
#' @param vetor_titulo Vetor com títulos das colunas.
#' @param indice_sigla_preco Sigla a ser substituída para "PR" nos dados de texto.
#' @param nota_roda_pe Vetor de notas de rodapé (opcional).
#' @param espanhol Lógico para definir idioma do cabeçalho (padrão: FALSE).
#' @param dados Data frame com os dados da planilha.
#' @param Tab_Fundo_cinza Data frame com linhas que devem ter fundo cinza.
#' @param referencia_filtrado Data frame com cores específicas de indicadores.
#' @param cor_fundo_menu Data frame com cores e estilos para cada título de coluna.
#' @param formatacao_numero Formatação que entra nos número na 'createStyle'. Por padrão, "0.0".
#' @param print Se TRUE, abre a planilha automaticamente após a criação.
#'
#' @details A função configura uma aba completa, mesclando células e aplicando estilos específicos para formatação de correlações. Adiciona cabeçalhos customizados, cores, e estilo condicional de acordo com a natureza dos dados.
#'
#' @return Retorna o workbook `wb` com a nova aba formatada.
#'
#' @import openxlsx
#' @import dplyr
#' @importFrom tibble tibble
#' @importFrom tidyr drop_na
#' @importFrom stringr str_detect
#' @importFrom stringr str_replace_all
#' @importFrom stats cor

#' @examples
#'
#' base::print("Sem Exemplo")
#'
#' @export

IndicePlanilhaCorrelacao <- function(
    wb,
    logo_innovare,
    nome_aba,
    ano_titulo_planilha,
    pasta_nome,
    vetor_titulo,
    indice_sigla_preco = "PR1",
    nota_roda_pe = NULL, # NULL
    espanhol = FALSE,
    dados,
    Tab_Fundo_cinza,
    referencia_filtrado,
    cor_fundo_menu,
    formatacao_numero = "0.0",
    print = FALSE
)
{# Start: função 'IndicePlanilhaCorrelacao'

  {# Start: Configuração inicial da tabela e formatação dos dados

    # Definindo o início da tabela numérica
    inicio_tabela_numerico = 9 - 1

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
      dplyr::select(-c(linhas_negrito, indice_sigla)) %>%
      base::colnames()

    tabela_dados_numericos = dados %>%
      dplyr::select(dplyr::all_of(rodando_de_fato)) %>%
      dplyr::select(-c(1, 2)) %>%
      dplyr::mutate(
        dplyr::across(
          .cols = dplyr::everything(),
          ~ base::as.numeric(.)
        )
      )

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

    # Escrevendo dados de texto na nova aba, ajustando o valor de "PR1" para "PR"
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

    # Inserindo o logo na aba
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

      # Ajustando a altura da primeira linha
      openxlsx::setRowHeights(
        wb = wb,
        sheet = nome_aba,
        rows = base::c(1),
        heights = base::c(12.75)
      )

      # Ajustando a altura das linhas 2 a 5
      openxlsx::setRowHeights(
        wb = wb,
        sheet = nome_aba,
        rows = base::c(2:5),
        heights = base::c(21)
      )

      # Ajustando a altura da linha 6
      openxlsx::setRowHeights(
        wb = wb,
        sheet = nome_aba,
        rows = base::c(6),
        heights = base::c(15.75)
      )

      # Ajustando a altura das linhas de dados de texto, do início até o fim da tabela
      openxlsx::setRowHeights(
        wb = wb,
        sheet = nome_aba,
        rows = base::c(8:(base::nrow(tabela_dados_texto) + inicio_tabela_numerico + 3)),
        heights = base::c(12.75)
      )

      # Ajustando a altura das linhas com quebras de linha identificadas na tabela
      openxlsx::setRowHeights(
        wb = wb,
        sheet = nome_aba,
        rows = base::c(tabela_pulou_linha),
        heights = base::c(3.75)
      )

      # Ajustando a altura da linha 7
      openxlsx::setRowHeights(
        wb = wb,
        sheet = nome_aba,
        rows = base::c(7),
        heights = base::c(25.5)
      )

    }# End: Ajustando alturas específicas das linhas

    {# Start: Ajustando larguras específicas das colunas

      # Definindo a largura da coluna 1
      openxlsx::setColWidths(
        wb = wb,
        sheet = nome_aba,
        cols = base::c(1),
        widths = base::c(6)
      )

      # Definindo a largura da coluna 2
      openxlsx::setColWidths(
        wb = wb,
        sheet = nome_aba,
        cols = base::c(2),
        widths = base::c(64)
      )

      for( i in 3:(base::ncol(tabela_dados_numericos) + 2) )
      {# Start: Iterando sobre as colunas de `tabela_dados_numericos` para ajustar a largura com base no comprimento do texto

        # Calculando a largura da coluna com base no texto em `vetor_titulo`
        calculox <- 1.043 * base::max(stringr::str_length(base::unlist(base::strsplit(vetor_titulo[i], "\n"))))

        # Ajustando a largura da coluna `i`, com largura mínima de 8.43
        openxlsx::setColWidths(
          wb = wb,
          sheet = nome_aba,
          cols = i,
          widths = base::ifelse(calculox < 8.43, 8.43, calculox)
        )

      }# End: Iterando sobre as colunas de `tabela_dados_numericos` para ajustar a largura com base no comprimento do texto

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

      # Escrevendo o nome da pasta na célula abaixo do título
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

      # Escrevendo o subtítulo "Correlação de atributos com a Satisfação Geral - 2024"
      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = base::paste0(
          base::ifelse(espanhol == FALSE, "Correla\u00E7\u00E3o de atributos com a Satisfa\u00E7\u00E3o Geral - ", "Correlaci\u00F3n de atributos con la Satisfacci\u00F3n General - "),
          ano_titulo_planilha),
        startCol = 2,
        startRow = 6,
        rowNames = FALSE,
        colNames = FALSE
      )

      # Aplicando estilo ao subtítulo: fonte 12, negrito, alinhamento à esquerda e com wrap de texto
      openxlsx::addStyle(
        wb = wb,
        sheet = nome_aba,
        style = openxlsx::createStyle(
          fontSize = 12,
          fontColour = "black",
          # fgFill = "#404040",
          textDecoration = c("BOLD"),
          halign = "left",
          valign = "center",
          wrapText = TRUE
        ),
        rows = 6,
        cols = 2,
        gridExpand = TRUE
      )

    }# End: Escrevendo o cabeçalho do menu na aba e aplicando estilo

    {# Start: Configurando cabeçalhos da tabela e mesclando células para estrutura visual

      # Escrevendo uma célula vazia na posição especificada
      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = " ",
        startCol = 1,
        startRow = 7,
        rowNames = FALSE,
        colNames = FALSE
      )

      # Escrevendo os títulos das colunas, exceto as duas primeiras, a partir da coluna 3
      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = base::t(base::data.frame(vetor_titulo[-c(1, 2)])),
        startCol = 3,
        startRow = 7,
        rowNames = FALSE,
        colNames = FALSE
      )

      # Escrevendo o título "Atributos" na coluna 2, linha 7
      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = "Atributos",
        startCol = 2,
        startRow = 7,
        rowNames = FALSE,
        colNames = FALSE
      )

      # Aplicando estilo à célula do título "Atributos": fonte branca, fundo escuro, negrito e centralizado
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
          wrapText = TRUE
        ),
        rows = 7,
        cols = 2,
        gridExpand = TRUE
      )


      for( cori in 3:base::length(rodando_de_fato) )
      {# Start: Aplicando cores específicas a cada coluna com base em `cor_fundo_menu`

        openxlsx::addStyle(
          wb = wb,
          sheet = nome_aba,
          style = openxlsx::createStyle(
            fontSize = 10,
            fontColour = cor_fundo_menu %>%
              dplyr::filter(split_nome == base::colnames(dados)[cori]) %>%
              dplyr::select(letra) %>%
              dplyr::pull(),
            fgFill = cor_fundo_menu %>%
              dplyr::filter(split_nome == base::colnames(dados)[cori]) %>%
              dplyr::select(cor) %>%
              dplyr::pull(),
            textDecoration = c("BOLD"),
            halign = "center",
            valign = "center",
            wrapText = TRUE
          ),
          rows = 7,
          cols = cori,
          gridExpand = TRUE
        )

      }# End: Aplicando cores específicas a cada coluna com base em `cor_fundo_menu`

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

        indicadores_rodando_procurar = dados %>%
          dplyr::filter(indice_sigla %in% indicadores_rodando) %>%
          dplyr::select(IDAR) %>%
          dplyr::pull()

        # # Definindo 'PR'
        # indicadores_rodando_procurar = base::ifelse(
        #   indicadores_rodando == indice_sigla_preco,
        #   "PR",
        #   indicadores_rodando
        # )

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

            indicadores_rodando_procurar = base::unique(indicadores_rodando_procurar)
            separado = base::list()
            w = 1

            for( w in 1:base::length(indicadores_rodando_procurar) )
            {# Start: separando pelas quebras de sequência

              separado[[w]] = base::which(
                tabela_dados_texto$IDAR %in% c(indicadores_rodando_procurar[w])
              ) + inicio_tabela_numerico

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

              for( k in 3:c(base::length(rodando_de_fato)) )
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
                  numFmt = formatacao_numero
                ),
                rows = indicadores_linhas[j],
                cols = c(3:c(base::length(rodando_de_fato))),
                gridExpand = FALSE
              ) %>%
                base::suppressWarnings()

              # Aplicando estilo à coluna "IAOP", se existir
              if( base::any(stringr::str_detect(rodando_de_fato, "IAOP")) )
              {# Start: Aplicando estilo à coluna "IAOP", se existir

                openxlsx::addStyle(
                  wb = wb,
                  sheet = nome_aba,
                  style = openxlsx::createStyle(
                    fontSize = 10,
                    fontColour = 'black',
                    fgFill = fundo_cor,
                    textDecoration = letra_negrito,
                    halign = "center",
                    valign = "center",
                    wrapText = TRUE,
                    numFmt = "0.0%"
                  ),
                  rows = indicadores_linhas[j],
                  cols = base::which(stringr::str_detect(rodando_de_fato, "IAOP")),
                  gridExpand = FALSE
                ) %>%
                  base::suppressWarnings()

              }# End: Aplicando estilo à coluna "IAOP", se existir

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
      firstActiveRow = 8,
      firstActiveCol = 3
    )

  }# End: Montando a planilha

  if( print == TRUE )
  {# Start:Abrindo a planilha se o parâmetro 'print' for TRUE

    openxlsx::openXL(wb)

  }# End: Abrindo a planilha se o parâmetro 'print' for TRUE

  # Retornando o workbook com todas as modificações
  base::return(wb)

}# End: função 'IndicePlanilhaCorrelacao'
