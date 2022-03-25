<h1>geraboh</h1>

<p>
  <img src="https://img.shields.io/badge/maintained%3F-Yes-339933?style=flat-square">&nbsp;<img src="https://img.shields.io/github/license/stenioas/malpi?style=flat-square">&nbsp;<img src="https://img.shields.io/github/issues/stenioas/malpi?color=violet&style=flat-square">&nbsp;<img src="https://img.shields.io/github/stars/stenioas/malpi?style=flat-square">
</p>

<samp>Um script python que gera o Boletim de Ocupação Hoteleira(BOH) para hotéis e pousadas que utilizam o sistema <a href="https://www.hospedin.com">Hospedin</a>.</samp>

> **ATENÇÃO!!!**
>
> Atualmente o script é compatível apenas com plataformas Windows.

### PRÉ-REQUISITOS

- Conexão com a internet.
- Última versão do Python, você pode baixá-lo [aqui](https://www.python.org/).
- Todas as dependências abaixo satisfeitas.

### DEPENDÊNCIAS

- beautifulsoup4
- mechanize
- pywin32
- openpyxl

### INSTALANDO

**1.** Com o python já instalado, abra um prompt de comando do windows(cmd).

**2.** Atualize o gerenciador de pacotes com o comando abaixo:

    python -m pip install --upgrade pip

**3.** Instale as dependências com o comando abaixo:

    pip install beautifulsoup4 mechanize pywin32 openpyxl

**4.** Baixe o arquivo zip do projeto [aqui](https://github.com/stenioas/geraboh/archive/refs/heads/master.zip).

**5.** Extraia o arquivo baixado onde achar melhor.

### CONFIGURANDO

#### **Dados do Estabelecimento**

Utilizando o editor de texto de sua preferência, edite o arquivo `settings.ini`, que está dentro do diretório conf, e preencha a sessão **"ESTABELECIMENTO"** seguindo o padrão do exemplo abaixo:

Exemplo para configuração do estabelecimento:

```ini
[ESTABELECIMENTO]
nome=Meu Hotel
nome_na_url=meu-hotel
distrito=MEIRELES
municipio=FORTALEZA
uf=CE
uhs=40
leitos=118
cadastro_mtur=123456789012345
```

##### **ATRIBUTOS DO ESTABELECIMENTO**

| Atributo      | Descrição                                                                                                                                |
| ------------- | ---------------------------------------------------------------------------------------------------------------------------------------- |
| nome          | O nome do estabelecimento, ex: `Meu Hotel`.                                                                                              |
| nome_na_url   | O valor desse atributo está descrito na URL do Hospedin, após o login ser efetuado, ex: https://pms.hospedin.com/nome-do-hotel-aqui.     |
| distrito      | Distrito/Bairro da localização do estabelecimento.                                                                                       |
| municipio     | Município/Cidade da localização do estabelecimento.                                                                                      |
| uf            | Estado da localização do estabelecimento.                                                                                                |
| uhs           | O total de unidades hoteleiras que o estabelecimento possui.                                                                             |
| leitos        | O total de leitos que o estabelecimento suporta.                                                                                         |
| cadastro_mtur | Número de cadastro junto ao Ministério do Turismo. Esse atributo deve ser preenchido com apenas números, sem outros tipos de caracteres. |

#### **Login Automático**

> **!!! ATENÇÃO !!!** O usuário utilizado para realizar o login automático deve ter as suas pemissões configuradas apenas para este propósito, visto que as credencias ficarão expostas neste arquivo, evitando acesso indesejado ao sistema!!!

Mais uma vez, utilizando o editor de texto de sua preferência, edite o arquivo `settings.ini`, que está dentro do diretório conf, e preencha a sessão **"LOGIN"** seguindo o padrão do exemplo abaixo:

Exemplo para configuração do login automático:

```ini
[LOGIN]
auto=on
usuario=usuario@email.com.br
senha=123456
```

##### **ATRIBUTOS DE CONFIGURAÇÃO**

| Atributo | Descrição                                                                                                   |
| -------- | ----------------------------------------------------------------------------------------------------------- |
| auto     | Define se será solicitado usuário e senha sempre que executar a aplicação. Valores permitidos `on` e `off`. |
| usuario  | E-mail de login do usuário que será utilizado                                                               |
| senha    | Senha de login do usuário que será utilizado                                                                |
