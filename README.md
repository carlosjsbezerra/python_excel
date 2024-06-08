# Projeto de Manipulação de Planilhas Excel com Python

Este projeto é um exemplo de manipulação de planilhas Excel utilizando a biblioteca `pandas` e `openpyxl` em Python. O objetivo principal é demonstrar como carregar, modificar, e salvar planilhas Excel, além de inserir novas colunas e aplicar funções específicas dentro das planilhas.

## Estrutura do Projeto

- `Pasta001/planilha001.xlsx`: Planilha Excel original que será carregada e modificada.
- `Pasta002/nova_planilha.xlsx`: Nova planilha Excel que será salva após as modificações.

## Funcionalidades

1. **Carregar a Planilha Excel**: Carrega a planilha `planilha001.xlsx` localizada na pasta `Pasta001` utilizando `pandas` com o engine `openpyxl`.
2. **Exibir Colunas Originais**: Exibe as colunas originais da planilha carregada.
3. **Apagar Colunas Específicas**: Remove as colunas especificadas da planilha.
4. **Exibir Colunas Restantes**: Exibe as colunas restantes após a remoção das colunas especificadas.
5. **Manter e Reordenar Colunas**: Mantém apenas as colunas desejadas e as reordena se necessário.
6. **Inserir Nova Coluna**: Insere uma nova coluna chamada "Regional" entre as colunas "Nome da Escola" e "Município".
7. **Salvar Nova Planilha**: Salva a nova planilha modificada na pasta `Pasta002` com o nome `nova_planilha.xlsx`.
8. **Criar e Preencher Nova Aba**: Cria uma nova aba chamada `referencia_escolas` na nova planilha e preenche com dados específicos de municípios e regionais.
9. **Adicionar Função VLOOKUP**: Adiciona a função VLOOKUP nas células da nova planilha para realizar consultas entre as abas.
10. **Adicionar Filtros**: Adiciona filtros nas colunas especificadas na nova planilha.

## Pré-requisitos

- Python 3.6 ou superior
- Bibliotecas: `pandas`, `openpyxl`

## Instalação

1. Clone o repositório:
   ```sh
   git clone https://github.com/seu-usuario/nome-do-repositorio.git

2. Navegue até o diretório do projeto:
cd nome-do-repositorio

3. Instale as dependências:
pip install pandas openpyxl

## Uso
1. Certifique-se de que o arquivo `planilha001.xlsx` está na pasta Pasta001.
2. python seu_script.py

`python seu_script.py`

3. O script irá carregar a planilha, realizar as modificações e salvar a nova planilha em `Pasta002/nova_planilha.xlsx`.

Contribuição
Contribuições são bem-vindas! Sinta-se à vontade para abrir uma issue ou enviar um pull request.

Licença
Este projeto está licenciado sob a licença MIT. Veja o arquivo LICENSE para mais detalhes.

Contato
- Nome: CARLOS J S BEZERRA
- Email: karlos.juliano123@gmail.com
- GitHub: https://github.com/carlosjsbezerra
