# Innovare Pesquisa - Pacote R para Índices (IPpackageIndexes)

![Static Badge](https://img.shields.io/badge/STATUS-DESENVOLVIMENTO-critical)

![RStudio](https://img.shields.io/badge/RStudio-R-blue.svg)


Este pacotece é uma coleção abrangente de funcionalidades desenvolvidas para a aplicação de diversas funções do R, que são amplamente utilizadas pelos colaboradores.

   
## Instalação

- Obs.: Obrigatório já ter instalado o pacote 'IPpackage' antes

- Primeira instalação


```bash
  # Verifica se o pacote devtools está instalado
  if (!requireNamespace("devtools", quietly = TRUE)) 
    {

    # Se não estiver instalado, instala o pacote devtools
    install.packages("devtools")

    }
  #Carregando devtools
  library(devtools)

  #Baixando IPpackageIndexes
  devtools::install_github("IPpackage/IPpackageIndexes",force = TRUE)

  #Carregando IPpackageIndexes
  library(IPpackageIndexes)
```

- Se o pacote _IPpackage_ já estiver instalado em sua máquina e você desejar atualizá-lo.
```bash
  # Verifica se o pacote devtools está instalado
  if (!requireNamespace("devtools", quietly = TRUE)) 
    {

    # Se não estiver instalado, instala o pacote devtools
    install.packages("devtools")

    }
  #Carregando devtools
  library(devtools)

  #Descarregar IPpackageIndexes
  detach("package:IPpackageIndexes", unload = TRUE)

  #Desinstalar IPpackageIndexes
  remove.packages("IPpackageIndexes")

  #Baixando IPpackageIndexes
  devtools::install_github("IPpackage/IPpackageIndexes",force = TRUE)

  #Carregando IPpackageIndexes
  library(IPpackageIndexes)
```
