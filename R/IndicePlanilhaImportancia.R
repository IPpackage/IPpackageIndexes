#' IndicePlanilhaImportancia
#'
#' Gera e formata uma aba de planilha para exibir índices de importância associados a dados específicos.
#' Esta função configura uma aba dentro de um workbook, aplica formatações personalizadas para texto e dados numéricos,
#' e insere estilos e imagens conforme especificado. Opções de customização permitem incluir rodapés, títulos, e adaptar o idioma.
#'
#' @param wb Workbook criado com a biblioteca `openxlsx`, onde a aba será adicionada.
#' @param logo_innovare Caminho para o logo da empresa, que será inserido na aba.
#' @param nome_aba Nome da aba dentro do workbook para a qual a planilha será gerada.
#' @param pasta_nome Nome da pasta de origem dos dados, exibido no cabeçalho da aba.
#' @param dados Data frame contendo os dados de índice para a planilha. Deve incluir colunas para índices e atributos.
#' @param df_me Data frame com dados de margem de erro associados aos índices.
#' @param indice_sigla_preco (Opcional) Sigla para identificar preços na coluna `IDAR`, padrão é `"PR1"`.
#' @param nota_roda_pe (Opcional) Vetor de caracteres para notas de rodapé a serem incluídas na aba.
#' @param espanhol (Opcional) Booleano para definir o idioma, `FALSE` exibe o cabeçalho em português, `TRUE` em espanhol.
#' @param Tab_Fundo_cinza Data frame que identifica linhas específicas para preenchimento de fundo cinza.
#' @param referencia_filtrado Data frame de referência para personalização de cores de fundo e texto na planilha.
#' @param print (Opcional) Booleano para abrir automaticamente o workbook após a execução da função; padrão é `FALSE`.
#'
#' @details
#' A função `IndicePlanilhaImportancia` organiza e formata os dados de importância dos índices dentro de uma nova aba de planilha.
#' A formatação inclui configurações específicas para cada célula e linha, permitindo um design claro e estruturado.
#' A função primeiro cria uma aba e escreve os dados de texto e numéricos, aplicando formatações de estilo com base nos
#' parâmetros de cor e formato fornecidos. Alturas e larguras de linhas e colunas são ajustadas automaticamente para acomodar
#' o conteúdo. Também são aplicados estilos diferenciados, como negrito e alinhamento, além de formatação de células baseada
#' na presença de indicadores específicos nos dados.
#'
#' @section Estrutura:
#' - **Configuração de Alturas e Larguras**: Alturas de linha e larguras de coluna são definidas para otimizar a exibição dos dados.
#' - **Cabeçalhos e Títulos**: Insere um cabeçalho com o título "PLANILHA DE ÍNDICES" ou "PLANILLA DE ÍNDICES", baseado no idioma.
#' - **Formatos de Célula e Estilos**: Células recebem cores de fundo e estilo (negrito, alinhamento, etc.) com base nos valores de referência fornecidos.
#' - **Configuração de Margens de Erro**: Valores de margem de erro são ajustados conforme dados disponíveis em `df_me`.
#' - **Opções de Rodapé**: Notas de rodapé personalizadas podem ser adicionadas, exibindo informações adicionais sobre os dados da planilha.
#'
#' @return Retorna o workbook com a aba recém-adicionada e formatada.
#'
#' @import openxlsx
#' @import dplyr
#' @importFrom tidyr drop_na
#' @importFrom stringr str_detect
#' @importFrom stringr str_length

#' @examples
#'
#' base::print("Sem Exemplo")
#'
#' @export


IndicePlanilhaImportancia <- function(
    wb,
    logo_innovare,
    nome_aba,
    pasta_nome,
    dados,
    df_me,
    indice_sigla_preco = "PR1",
    nota_roda_pe = NULL, # NULL
    espanhol = FALSE,
    Tab_Fundo_cinza,
    referencia_filtrado,
    print = FALSE
)
{# Start: função 'IndicePlanilhaImportancia'

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

    vetor_titulo = rodando_de_fato

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
        cols = 2:base::length(rodando_de_fato),
        gridExpand = TRUE
      )

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

              if( !base::is.na(indicadores_rodando[j]) )
              {

                if(
                  referencia_filtrado %>%
                  dplyr::filter(IDAR %in% c(indicadores_rodando[j])) %>%
                  dplyr::select(formatacao_imp) %>%
                  dplyr::slice(1) %>%
                  dplyr::pull() %in% c("limpa")
                )
                {# Start: Verifica se a formatação da linha é do tipo 'limpa', caso em que nenhum plano de fundo colorido será aplicado

                  # Define a cor de fundo padrão para linhas 'limpas' (sem fundo colorido)
                  fund_cinza <- FALSE
                  fundo_cor <- if (fund_cinza == TRUE) "#F2F2F2" else "white"  # Define cor do fundo com base em 'fund_cinza'

                  # Estilo em negrito aplicado ao texto da linha
                  letra_negrito <- "BOLD"

                  # Insere um espaço na coluna de índice para preservar a estrutura da planilha
                  openxlsx::writeData(
                    wb = wb,
                    sheet = nome_aba,
                    x = " ",
                    startCol = 1,
                    startRow = base::which(dados$indice_sigla %in% c(indicadores_rodando[j])) + inicio_tabela_numerico,
                    rowNames = FALSE,
                    colNames = FALSE
                  )

                  # Aplica o estilo à linha com a formatação 'limpa', mantendo texto em negrito e alinhado centralmente
                  openxlsx::addStyle(
                    wb = wb,
                    sheet = nome_aba,
                    style = openxlsx::createStyle(
                      fontSize = 10,
                      fontColour = "black",
                      fgFill = fundo_cor,
                      textDecoration = c("BOLD"),
                      halign = "left",
                      valign = "center"
                    ),
                    rows = base::which(dados$indice_sigla %in% c(indicadores_rodando[j])) + inicio_tabela_numerico,
                    cols = 1:2,
                    gridExpand = TRUE
                  )

                }# End: Verifica se a formatação da linha é do tipo 'limpa', caso em que nenhum plano de fundo colorido será aplicado

              }

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
                  numFmt = "0.0%"
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

}# End: função 'IndicePlanilhaImportancia'
