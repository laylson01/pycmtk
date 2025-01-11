# ERP Empresarial

Este projeto é uma aplicação de gestão empresarial com funcionalidades para cadastro de clientes, pesquisa de informações, e geração de relatórios a partir de dados armazenados em arquivos Excel. A interface foi desenvolvida utilizando `customtkinter`, com temas e widgets modernos, e interage com um banco de dados SQLite para armazenar e gerenciar as informações dos clientes.

## Funcionalidades

- **Cadastro de Clientes**: Adicione clientes ao sistema, preenchendo informações como nome, CPF, endereço, email, entre outros.
- **Pesquisa de Cliente**: Permite a pesquisa de clientes por nome no banco de dados ou arquivo Excel.
- **Relatórios**: Carrega e exibe dados de clientes armazenados em um arquivo Excel com a opção de visualizar em formato de tabela.
- **Interface de Usuário**: Interface gráfica moderna com tabs para navegar entre as diferentes funcionalidades (Home, Clientes, Produtos, Vendas, Relatórios).

## Requisitos

- Python 3.x
- Bibliotecas:
  - `customtkinter`
  - `tkcalendar`
  - `pandas`
  - `openpyxl`
  - `sqlite3`
  - `ttk` (para criação de tabelas)
  
Você pode instalar as dependências usando o `pip`:

```bash
pip install customtkinter tkcalendar pandas openpyxl
```

## Estrutura do Projeto

- **dashboard_empresarial.py**: Contém a lógica da aplicação e a interface gráfica.
- **database.py**: Define a classe `Database`, que gerencia o banco de dados SQLite para armazenar as informações dos clientes.
- **excel_handler.py**: Define a classe `ExcelHandler`, que lida com a exportação e importação de dados no arquivo Excel `clientes.xlsx`.
- **clientes.db**: Banco de dados SQLite que armazena as informações dos clientes.
- **clientes.xlsx**: Arquivo Excel utilizado para armazenar e carregar dados dos clientes.

## Como Usar

1. **Iniciar o programa**: Execute o script `dashboard_empresarial.py` para iniciar a aplicação.
2. **Adicionar Cliente**: Vá para a aba "Clientes", preencha os campos do formulário e clique no botão "Adicionar Cliente" para salvar os dados no banco de dados e no arquivo Excel.
3. **Pesquisar Cliente**: Na aba "Home", digite o nome de um cliente no campo de pesquisa e clique em "Pesquisar" para ver as informações detalhadas do cliente.
4. **Relatórios**: Na aba "Relatórios", clique em "Carregar Dados do Excel" para visualizar os dados dos clientes em formato de tabela.

## Arquivos Excel

- A aplicação utiliza um arquivo Excel chamado `clientes.xlsx` para carregar e armazenar dados dos clientes. Caso o arquivo não esteja presente, será exibida uma mensagem de erro.
  
## Contribuições

Sinta-se à vontade para contribuir para o projeto. Para isso, siga as etapas abaixo:

1. Faça o fork do projeto.
2. Crie uma nova branch (`git checkout -b feature/nova-feature`).
3. Faça suas alterações.
4. Envie um pull request com uma descrição detalhada das mudanças.
## Imagem da Interface

Aqui está uma captura de tela da interface do ERP Empresarial:

![Interface do ERP](https://raw.githubusercontent.com/laylson01/pycmtk/refs/heads/main/Cliente_tab.png)

## Licença

Este projeto está licenciado sob a Licença MIT - consulte o arquivo [LICENSE](LICENSE) para mais detalhes.