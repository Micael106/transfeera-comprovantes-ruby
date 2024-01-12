````markdown
# Processador de Arquivos XLSX

Este script processa arquivos XLSX em um diretório fornecido, extrai dados com base em colunas específicas e faz o download de arquivos PDF correspondentes a URLs fornecidas. Também gera um novo arquivo Excel com linhas contendo o status 'DEVOLVIDO'.

## Requisitos

- Ruby (Instale a partir do [site oficial](https://www.ruby-lang.org/pt/documentation/installation/))
- Gema `rubyXL` (Instale com `gem install rubyXL`)
- Gema `open-uri-cached` (Instale com `gem install open-uri-cached`)

## Uso

1. Clone ou baixe o repositório.

2. Abra um terminal e navegue até o diretório que contém o script:

   ```bash
   cd caminho/do/diretorio/do/script
   ```
````

3. Execute o script com o seguinte comando:

   ```bash
   ruby app.rb <caminho_do_diretorio>
   ```

   Substitua `<caminho_do_diretorio>` pelo caminho para o diretório que contém seus arquivos XLSX.

## Visão Geral do Script

- O script lê arquivos XLSX no diretório especificado.
- Ele extrai dados das colunas especificadas e faz o download de arquivos PDF a partir das URLs correspondentes.
- Um novo arquivo Excel é gerado com linhas contendo o status 'DEVOLVIDO'.

## Notas

- Certifique-se de instalar as gemas necessárias usando os comandos fornecidos.
- Garanta conectividade à internet adequada, pois o script faz o download de arquivos PDF a partir de URLs.

Sinta-se à vontade para personalizar o script ou contribuir para sua melhoria!

```

```
