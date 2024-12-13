#' @title IndicePlanilhaGeral
#' @description Gera uma planilha detalhada e formatada em `wb` com uma estrutura de índice geral. A função cria uma aba com formatação específica e personalizada, incluindo estilos e mesclagens de células, baseados nos parâmetros fornecidos.
#'
#' @param wb Workbook onde a aba será criada.
#' @param logo_innovare Caminho do arquivo da imagem de logotipo a ser inserido na planilha.
#' @param nome_aba Nome da aba a ser adicionada ao workbook.
#' @param pasta_nome Nome da pasta exibida no cabeçalho.
#' @param manter_ao_remover_dados Colunas a serem preservadas ao remover dados de índices específicos.
#' @param remover_dos_indices Vetor de índices que devem ter seus valores removidos da planilha.
#' @param indice_sigla_preco Identifica o índice de preço e ajusta o valor "PR1" para "PR".
#' @param nota_roda_pe Vetor de notas de rodapé opcionais.
#' @param inverter_condicional Vetor de índices que invertem a condição de formatação condicional.
#' @param espanhol Se TRUE, exibe cabeçalho e títulos em espanhol.
#' @param anos Vetor de anos a ser exibido nos cabeçalhos.
#' @param dados Data frame com os dados para a planilha.
#' @param Tab_Fundo_cinza Data frame que define cores e estilos de fundo para linhas específicas.
#' @param referencia_filtrado Data frame com cores para formatação de fundo e texto para cada índice.
#' @param print Se TRUE, abre a planilha ao final.
#'
#' @details A função gera uma nova aba com layout e formatação específicos usando `openxlsx`, com inserção de dados de texto e numéricos e personalização avançada de células. O layout é otimizado para índices gerais, incluindo estilos com base nas cores dos indicadores e condições visuais para valores.
#'
#' A função permite ajustar colunas específicas, configurar o cabeçalho em espanhol, inserir um logotipo e adicionar notas de rodapé. Suporta formatação condicional com cor de célula (vermelho/verde) com base em valores diferenciais e margem de erro para os índices especificados.
#'
#' @return Workbook atualizado com a nova aba formatada.
#'
#' @import openxlsx
#' @import dplyr
#' @importFrom tibble tibble
#' @importFrom tidyr drop_na
#' @importFrom stringr str_detect
#' @importFrom stringr str_replace_all

#' @examples
#'
#' base::print("Sem Exemplo")
#'
#' @export

IndicePlanilhaGeral <- function(
    wb,
    logo_innovare,
    nome_aba,
    pasta_nome,
    manter_ao_remover_dados = NULL,
    remover_dos_indices = NULL,
    indice_sigla_preco = "PR1",
    nota_roda_pe = NULL, # NULL
    inverter_condicional = c("IIQP", "IIC", "IICP", "IICP_BR"),
    espanhol = FALSE,
    anos,
    dados,
    Tab_Fundo_cinza,
    referencia_filtrado,
    print = FALSE
)
{# Start: função 'IndicePlanilhaGeral'

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
      dplyr::select(-c(linhas_negrito)) %>%
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

      # Ajustando a altura das linhas 2 a 5 na aba especificada
      openxlsx::setRowHeights(
        wb = wb,
        sheet = nome_aba,
        rows = c(2:5),
        heights = c(21)
      )

      # Ajustando a altura das linhas com quebras de linha identificadas na tabela
      openxlsx::setRowHeights(
        wb = wb,
        sheet = nome_aba,
        rows = c(tabela_pulou_linha),
        heights = c(3.75)
      )

      # Ajustando a altura das linhas 1 e 6
      openxlsx::setRowHeights(
        wb = wb,
        sheet = nome_aba,
        rows = c(1, 6),
        heights = c(12.75)
      )

      # Ajustando a altura da linha 7
      openxlsx::setRowHeights(
        wb = wb,
        sheet = nome_aba,
        rows = c(7),
        heights = c(25.5)
      )

    }# End: Ajustando alturas específicas das linhas

    {# Start: Ajustando larguras específicas das colunas

      # Definindo a largura da coluna 1
      openxlsx::setColWidths(
        wb = wb,
        sheet = nome_aba,
        cols = 1,
        widths = 6
      )

      # Definindo a largura da coluna 2
      openxlsx::setColWidths(
        wb = wb,
        sheet = nome_aba,
        cols = 2,
        widths = 64
      )

      # Definindo a largura das colunas de 3 até o número total de colunas em 'dados'
      openxlsx::setColWidths(
        wb = wb,
        sheet = nome_aba,
        cols = 3:ncol(dados),
        widths = 8.43
      )

      if( base::any(base::colnames(dados) %>% stringr::str_detect("Diferen\u00E7a")) )
      {# Start: Ajustando largura para colunas que contêm a palavra "Diferença", caso existam

        openxlsx::setColWidths(
          wb = wb,
          sheet = nome_aba,
          cols = base::which(
            base::colnames(dados) %>%
              stringr::str_detect("Diferen\u00E7a")
          ),
          widths = 12
        )

      }# End: Ajustando largura para colunas que contêm a palavra "Diferença", caso existam

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

      # Escrevendo células vazias nas posições específicas da aba
      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = " ",
        startCol = 1,
        startRow = 6,
        rowNames = FALSE,
        colNames = FALSE
      )

      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = " ",
        startCol = 1,
        startRow = 7,
        rowNames = FALSE,
        colNames = FALSE
      )

      # Escrevendo a margem de erro com formatação específica na aba
      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = base::paste0(
          base::ifelse(espanhol == FALSE, "Margem de erro: ", "Margen de error: "),
          df_me %>%
            dplyr::filter(seriehistorica_ano == base::max(anos)) %>%
            dplyr::select(me) %>%
            dplyr::mutate(me = IPpackage::round_excel(me,0)) %>%
            dplyr::pull() %>%
            stringr::str_replace_all("[.]", ", "),
          "%"
        ),
        startCol = 2,
        startRow = 6,
        rowNames = FALSE,
        colNames = FALSE
      )

      # Escrevendo título "Atributos" e os nomes das colunas de 'dados' (excluindo as duas primeiras)
      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = "Atributos",
        startCol = 2,
        startRow = 7,
        rowNames = FALSE,
        colNames = FALSE
      )

      x = rodando_de_fato[-c(1:2)]

      if( espanhol == TRUE & base::any(stringr::str_detect(x, "Diferen\u00E7a")) )
      {# Start: trocando o texto para espanhol

        x = x %>%
          stringr::str_replace_all("Diferen\u00E7a", "Diferencia")

      }# End: trocando o texto para espanhol

      openxlsx::writeData(
        wb = wb,
        sheet = nome_aba,
        x = data.frame(t(x)),
        startCol = 3,
        startRow = 7,
        rowNames = FALSE,
        colNames = FALSE
      )

      base::rm(x)

      # Removendo mesclagem de células existente e aplicando nova mesclagem na linha 6
      openxlsx::removeCellMerge(
        wb = wb,
        sheet = nome_aba,
        cols = 2:base::length(rodando_de_fato),
        rows = 6
      )

      openxlsx::mergeCells(
        wb = wb,
        sheet = nome_aba,
        cols = 2:base::length(rodando_de_fato),
        rows = 6
      )

      # Aplicando estilo à célula na linha 6, coluna 2
      openxlsx::addStyle(
        wb = wb,
        sheet = nome_aba,
        style = openxlsx::createStyle(
          fontSize = 10,
          fontColour = "white",
          fgFill = "#404040",
          textDecoration = c("BOLD"),
          halign = "left",
          valign = "center",
          wrapText = TRUE
        ),
        rows = 6,
        cols = 2,
        gridExpand = TRUE
      )

      # Aplicando estilo nas células da linha 7, colunas 2 até o número total de colunas em 'dados'
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

        indicadores_rodando_procurar = indicadores_rodando
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
                  numFmt = "0.0"
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

    {# Start: formatação condifional da diferença

      if( base::any(base::colnames(dados) %>% stringr::str_detect("Diferen\u00E7a")) )
      {# Start: Verifica se há colunas com "Diferença" nos nomes para aplicar formatação

        # Identifica as colunas que contêm a palavra "Diferença"
        dif_colunas <- base::which(
          base::colnames(dados) %>%
            stringr::str_detect("Diferen\u00E7a")
        )

        for( d in 1:base::length(dif_colunas) )
        {# Start: Loop para aplicar formatação em cada coluna de diferença

          for (k in 1:base::nrow(dados)) {

            if( !base::is.na(dados[k, 1]) )
            {# Start: Verifica se a linha atual tem valor não nulo na coluna de IDAR

              # Define as posições da coluna e linha para aplicação de estilo
              coluna_difereca <- base::which(base::colnames(dados) %>% stringr::str_detect("Diferen\u00E7a"))
              linha_diferenca <- k + inicio_tabela_numerico

              # Arredonda o valor de diferença com 1 casa decimal
              valor_diferenca <- IPpackage::round_excel(
                dados[k, base::which(base::colnames(dados) %>% stringr::str_detect("Diferen\u00E7a"))] %>%
                  dplyr::pull() %>%
                  base::as.numeric(),
                1
              )

              # Obtém o valor de margem de erro para o ano máximo
              valor_me <- IPpackage::round_excel(
                df_me %>%
                  dplyr::filter(seriehistorica_ano == base::max(anos)) %>%
                  dplyr::select(me) %>%
                  dplyr::pull(),
                1
              )

              if( !base::is.na(valor_diferenca) )
              {# Start: Se o valor de diferença não for NA, aplica a formatação condicional

                # Define estilo de negrito se a linha estiver marcada em 'negrito'
                letra_negrito <- if (linha_diferenca %in% c(negrito)) "BOLD" else NULL


                if( dados[k, 1] %in% inverter_condicional )
                {# Start: Formatação para indicadores específicos

                  # Aplica estilo de célula verde se o valor da diferença for menor que o negativo da margem de erro
                  if (valor_diferenca < c(-valor_me)) {

                    openxlsx::addStyle(
                      wb = wb,
                      sheet = nome_aba,
                      style = openxlsx::createStyle(
                        fontSize = 10,
                        fontColour = 'black',
                        fgFill = "#00FF00",
                        textDecoration = letra_negrito,
                        halign = "center",
                        valign = "center",
                        wrapText = TRUE,
                        numFmt = "0.0"
                      ),
                      rows = linha_diferenca,
                      cols = coluna_difereca,
                      gridExpand = FALSE
                    ) %>%
                      base::suppressWarnings()

                  }

                  # Aplica estilo de célula vermelha se o valor da diferença for maior que a margem de erro
                  if (valor_diferenca > valor_me) {

                    openxlsx::addStyle(
                      wb = wb,
                      sheet = nome_aba,
                      style = openxlsx::createStyle(
                        fontSize = 10,
                        fontColour = 'white',
                        fgFill = "#FF0000",
                        textDecoration = letra_negrito,
                        halign = "center",
                        valign = "center",
                        wrapText = TRUE,
                        numFmt = "0.0"
                      ),
                      rows = linha_diferenca,
                      cols = coluna_difereca,
                      gridExpand = FALSE
                    ) %>%
                      base::suppressWarnings()

                  }

                } else {# Start: Formatação padrão para outros indicadores



                  if( valor_diferenca > valor_me )
                  {# Start: Aplica estilo de célula verde se o valor da diferença for maior que a margem de erro

                    openxlsx::addStyle(
                      wb = wb,
                      sheet = nome_aba,
                      style = openxlsx::createStyle(
                        fontSize = 10,
                        fontColour = 'black',
                        fgFill = "#00FF00",
                        textDecoration = letra_negrito,
                        halign = "center",
                        valign = "center",
                        wrapText = TRUE,
                        numFmt = "0.0"
                      ),
                      rows = linha_diferenca,
                      cols = coluna_difereca,
                      gridExpand = FALSE
                    ) %>%
                      base::suppressWarnings()

                  }# End: Aplica estilo de célula verde se o valor da diferença for maior que a margem de erro

                  if( valor_diferenca < c(-valor_me) )
                  {# Start: Aplica estilo de célula vermelha se o valor da diferença for menor que o negativo da margem de erro

                    openxlsx::addStyle(
                      wb = wb,
                      sheet = nome_aba,
                      style = openxlsx::createStyle(
                        fontSize = 10,
                        fontColour = 'white',
                        fgFill = "#FF0000",
                        textDecoration = letra_negrito,
                        halign = "center",
                        valign = "center",
                        wrapText = TRUE,
                        numFmt = "0.0"
                      ),
                      rows = linha_diferenca,
                      cols = coluna_difereca,
                      gridExpand = FALSE
                    ) %>%
                      base::suppressWarnings()

                  }# End: Aplica estilo de célula vermelha se o valor da diferença for menor que o negativo da margem de erro

                }# End: Formatação padrão para outros indicadores

              }# End: Se o valor de diferença não for NA, aplica a formatação condicional

            }# End: Verifica se a linha atual tem valor não nulo na coluna de IDAR

          }

        }# End: Loop para aplicar formatação em cada coluna de diferença

      }# End: Verifica se há colunas com "Diferença" nos nomes para aplicar formatação

      # Remove as variáveis temporárias usadas no loop
      base::rm(coluna_difereca, k, linha_diferenca, valor_diferenca, valor_me) %>%
        base::suppressWarnings()

    }# End: formatação condifional da diferença

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

}# End: função 'IndicePlanilhaGeral'
