# Sistema de Cadastros de Cedentes, Captadores e Franqueados

Este é um projeto Python que implementa um sistema de cadastros para cedentes, captadores e franqueados. O sistema permite a criação, alteração, exclusão e filtragem de registros, bem como a exportação dos dados para uma planilha Excel no formato XLSX. Para executar o sistema, siga as instruções abaixo.

## Funcionalidades

- Cadastro de empresas com as seguintes informações:
  - CNPJ
  - Nome Fantasia
  - Razão Social
  - Captador
  - Gerente
  - Franqueado
  - Comissões
  - Desenho Operacional
  - Tipo de Operação
  - Status
  - Email
  - Telefone
  - E outras informações relevantes.

- Operações CRUD (Criar, Ler, Atualizar, Deletar) para os registros.

- Filtragem dos registros com base em critérios específicos.

- Exportação dos dados para uma planilha Excel no formato XLSX.

## Configuração

Antes de executar o sistema, siga estas etapas:

1. Crie uma tabela no banco de dados SQL Server para armazenar os registros.

2. Altere as credenciais de conexão com o banco de dados no arquivo `.env`. Certifique-se de fornecer o nome do servidor, nome do banco de dados, nome de usuário e senha corretos.

## Execução

Para executar o sistema, siga estas etapas:

1. Certifique-se de ter o Python instalado em seu sistema.

2. Instale as dependências do projeto executando o seguinte comando:

   ```
   pip install -r requirements.txt
   ```

3. Execute o aplicativo Python:

   ```
   python main.py
   ```

4. O aplicativo será iniciado, permitindo que você comece a cadastrar, atualizar e gerenciar os registros.

## Tecnologias Utilizadas

- Python
- SQL Server 2019 para o armazenamento de dados
- Versão do tkinter: 8.6

## Licença

Este projeto é distribuído sob a licença [MIT](LICENSE).

---