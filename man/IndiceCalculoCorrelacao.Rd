% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/IndiceCalculoCorrelacao.R
\name{IndiceCalculoCorrelacao}
\alias{IndiceCalculoCorrelacao}
\title{IndiceCalculoCorrelacao}
\usage{
IndiceCalculoCorrelacao(
  splits,
  vou_rodar,
  dados,
  indice_correlacao,
  caminho_temporario = "99.temp cal indi",
  caminho_resultado = "Output",
  nome_salvar = "Correlacao",
  metodo_rodar = "kendall"
)
}
\arguments{
\item{splits}{Um `data.frame` com os detalhes dos grupos para os quais as correlações serão calculadas.}

\item{vou_rodar}{Um `character` que indica a pasta ou o grupo a ser processado.}

\item{dados}{Um `data.frame` contendo os dados brutos, incluindo as variáveis e pesos necessários para o cálculo.}

\item{indice_correlacao}{Um `data.frame` com as especificações das variáveis e referências de índice a serem correlacionadas.}

\item{caminho_temporario}{Um `character` que define o diretório temporário onde os resultados intermediários serão salvos, padrão é `"99.temp cal indi"`.}

\item{caminho_resultado}{Um `character` que define o diretório de saída onde o resultado final será salvo, padrão é `"Output"`.}

\item{nome_salvar}{Um `character` que define o nome do arquivo final de saída, padrão é `"Correlacao"`.}

\item{metodo_rodar}{Por padrão, calcula a correlação de 'kendall' (kendall, pearson, etc).}
}
\description{
A função `IndiceCalculoCorrelacao` calcula correlações ponderadas entre variáveis especificadas e referência
em um conjunto de dados dividido por grupos. A função utiliza um método de ponderação para calcular correlações específicas,
organizando os resultados em arquivos temporários para otimizar a memória.
}
\details{
A função realiza uma série de etapas:
- Filtragem e preparação dos dados de acordo com os parâmetros de `splits` e `indice_correlacao`.
- Cálculo das correlações entre variáveis e variáveis de referência, com ponderação opcional via método de Kendall.
- Salvamento de resultados temporários para cada grupo em arquivos `.rds`.
- Geração de um arquivo final com todos os resultados de correlação para os grupos processados.

Internamente, a função `estima_CORR_peso` é usada para realizar o cálculo da correlação ponderada para as variáveis especificadas.
}
\examples{

base::print("Sem Exemplo")

}
