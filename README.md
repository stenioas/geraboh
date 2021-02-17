<h1 align="center">
  gera-boh
</h1>

<h3 align="center">
	BOH - <a href="https://hospedin.com/">Hospedin</a>
</h3>
<p align="center">Um simples script python para gerar o boletim de ocupação hoteleira de hotéis e pousadas que utilizam o PMS da Hospedin. Você pode usá-lo e modificá-lo como quiser.</p>

<p align="center">
  <img src="https://img.shields.io/badge/Maintained%3F-No-red?style=for-the-badge">
  <img src="https://img.shields.io/github/license/stenioas/malpi?style=for-the-badge">
  <img src="https://img.shields.io/github/issues/stenioas/malpi?color=violet&style=for-the-badge">
  <img src="https://img.shields.io/github/stars/stenioas/malpi?style=for-the-badge">
</p>

## Notas
* Atualmente o script só funciona em sistemas Windows.

## Pré-requisitos

- Uma conexão com a internet funcionando.
- Ter a última versão do [Python](https://www.python.org/) instalada.
- Todas as libs necessárias instaladas.
- Ter um conta Hospedin válida.

##### Libs necessárias:
- beautifulsoup4
- mechanize
- pywin32
- openpyxl

##### Execute o comando abaixo para instalar todas as libs:

	pip install beautifulsoup4 mechanize pywin32 openpyxl

## Obtendo o script

### git
	git clone https://github.com/stenioas/geraboh

## Como usar

### Primeira etapa

Editar o arquivo `gera-boh.py` e preencher a sessão abaixo com os dados corretos do hotel/pousada:

    # DADOS DA EMPRESA

    company_url_name = "0"
    company_name = "0"
    num_embratur = "0"
    num_of_rooms = "0"
    num_of_beds = "0"
Exemplo:

	# DADOS DA EMPRESA
	
    company_url_name = "https://pms.hospedin.com/nome-do-seu-hotel-ou-pousada-aqui"
    company_name = "nome-do-seu-hotel-ou-pousada-aqui"
    num_embratur = "numero-embratur-aqui"
    num_of_rooms = "total-de-quartos-aqui"
    num_of_beds = "numero-de-leitos-aqui"

### Segunda etapa

Após todos os requisitos atendidos e a primeira etapa realizada de forma correta você deve iniciar o prompt de comando do Windows, acessar o diretório do projeto e executar o comando abaixo:

	python gera-boh.py

Após a execução do comando sem nenhum retorno de erro você deve digitar seu email e senha utilizados para efetuar login na plataforma do Hospedin, e em seguida digitar o mês desejado(2 dígitos) e ano desejado(4 dígitos) para gerar o BOH.

<h2 align="center">Obrigado por dedicar seu tempo a conhecer o meu projeto!</h2>

