# Hexa Stock - Sistema de Gerenciamento de Estoque de TI

![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)

Hexa Stock é uma aplicação desktop para gerenciamento de inventário e rastreamento de ativos de TI, desenvolvida em Python com uma interface gráfica criada em Tkinter. A aplicação utiliza o Google Sheets como um backend de banco de dados, permitindo o uso colaborativo em tempo real por múltiplos usuários em uma rede local, sem a necessidade de instalar um servidor de banco de dados ou realizar configurações complexas de rede.

<img width="1363" height="717" alt="image" src="https://github.com/user-attachments/assets/39e20b86-9e67-42a0-8527-270635e8447f" />
<img width="1365" height="718" alt="image" src="https://github.com/user-attachments/assets/26cff9f2-699f-4d94-9289-03efb0ac015e" />



## Funcionalidades Principais

O sistema é dividido em dois módulos principais, acessíveis por abas:

### Estoque da Informática
- **Dashboard Interativo:** Visualização rápida de métricas importantes como total de itens, tipos de itens únicos, itens com estoque baixo e movimentações no mês.
- **Cadastro de Itens:** Adição de novos equipamentos com campos para nome, categoria, descrição, quantidade inicial, estoque mínimo e número de série/SKU.
- **Gestão de Estoque:** Edição e exclusão de itens do inventário.
- **Ciclo de Movimentação Completo:**
  - **Saída:** Registro de saída de itens do estoque para um setor.
  - **Entrada/Devolução:** Registro do retorno de itens ao estoque.
  - **Descarte:** Processo formal para dar baixa em equipamentos defeituosos, com campo obrigatório para motivo ou laudo.
- **Pesquisa e Filtro:** Barra de pesquisa que filtra a lista de equipamentos em tempo real.
- **Relatórios em HTML:** Geração de relatórios customizáveis do estado do estoque, com opção de incluir o histórico de movimentações e filtrar por datas.

### Movimentação Entre Setores
- **Rastreamento de Ativos:** Módulo dedicado para registrar a movimentação de equipamentos que já estão em uso entre diferentes setores, sem afetar o estoque principal.
- **Formulário Inteligente:** O formulário se adapta ao tipo de equipamento, exigindo ServiceTag apenas para itens como Desktops e Monitores.
- **Suporte a "Kits":** Funcionalidade para registrar a movimentação de um kit (ex: 2 monitores + 1 desktop) inserindo os patrimônios e ServiceTags de cada componente.
- **Histórico Dedicado:** Tabela de visualização com todo o histórico de movimentações entre setores.
- **Fluxo de Regularização:** Permite marcar uma movimentação como "Regularizada", indicando que o processo foi validado em um sistema externo.
- **Relatórios com Filtros:** Geração de relatórios específicos para esta aba, com filtros por status (Pendente, Regularizado) e por intervalo de datas.

## Tecnologias Utilizadas
- **Linguagem:** Python 3
- **Interface Gráfica:** Tkinter (ttk)
- **Backend/Banco de Dados:** Google Sheets
- **Bibliotecas Principais:** `gspread`, `pandas`, `oauth2client`

## Configuração e Instalação

Siga os passos abaixo para executar o projeto.

### Pré-requisitos
- Python 3.9 ou superior.
- Uma conta Google.

### Passo 1: Clonar o Repositório
```bash
git clone [https://github.com/seu-usuario/seu-repositorio.git](https://github.com/seu-usuario/seu-repositorio.git)
cd seu-repositorio
```

### Passo 2: Instalar as Dependências
É altamente recomendado criar um ambiente virtual.
```bash
python -m venv venv
source venv/bin/activate  # No Windows: venv\Scripts\activate
```
Em seguida, instale as bibliotecas necessárias:
```bash
pip install gspread pandas oauth2client
```

### Passo 3: Configuração da API do Google Cloud
Para que o programa possa acessar sua Planilha Google, você precisa de uma chave de acesso.

1.  **Acesse o [Google Cloud Console](https://console.cloud.google.com/)**.
2.  **Crie um Novo Projeto** (ex: "Hexa Stock API").
3.  No menu de busca, procure e **ative as seguintes APIs**:
    - `Google Drive API`
    - `Google Sheets API`
4.  No menu lateral (☰), vá para **"APIs e Serviços" > "Credenciais"**.
5.  Clique em **"+ CRIAR CREDENCIAIS"** e escolha **"Conta de serviço"**.
6.  Dê um nome à conta (ex: `editor-planilha-stock`), clique em "CRIAR E CONTINUAR" e depois em "CONCLUÍDO".
7.  Na tela de credenciais, clique na conta de serviço que você acabou de criar.
8.  Vá para a aba **"CHAVES"**, clique em **"ADICIONAR CHAVE" > "Criar nova chave"**.
9.  Escolha o formato **JSON** e clique em "CRIAR".
10. **Um arquivo `.json` será baixado. Renomeie-o para `credentials.json` e coloque-o na pasta raiz do projeto.**

### Passo 4: Configuração da Planilha Google
1.  **Crie uma nova Planilha Google**. Dê a ela o nome exato definido no código: **`ControleDeEstoqueTI`**.
2.  **Crie 4 abas** com os seguintes nomes exatos:
    - `equipamentos`
    - `movimentacoes`
    - `movimentacoes_setores`
    - `config`
3.  **Adicione os Cabeçalhos** na primeira linha de cada aba, conforme abaixo:
    - **`equipamentos`**: `id`, `nome`, `numero_serie`, `descricao`, `quantidade`, `status`, `data_cadastro`, `estoque_minimo`, `categoria`
    - **`movimentacoes`**: `id_movimentacao`, `id_equipamento_fk`, `tipo_movimentacao`, `quantidade_movida`, `destino_origem`, `solicitante`, `chamado`, `responsavel_movimentacao`, `data_movimentacao`, `motivo_laudo`
    - **`movimentacoes_setores`**: `id`, `data_movimentacao`, `responsavel`, `tipo_equipamento`, `patrimonio`, `servicetag`, `setor_origem`, `setor_destino`, `observacao`, `chamado`, `solicitante`, `status_regularizacao`
    - **`config`**: `parametro`, `valor` (e preencha com seus destinos, categorias, etc.).
4.  **Compartilhe a Planilha:**
    - Abra o arquivo `credentials.json` e copie o valor do campo `"client_email"`.
    - Na Planilha Google, clique em **"Compartilhar"**.
    - Cole o e-mail no campo de compartilhamento, dê a ele a permissão de **"Editor"** e clique em "Enviar".
  
## Sobre os Custos (Custo de Uso)

Este projeto utiliza as APIs do Google Cloud (Google Sheets e Google Drive) para funcionar como um banco de dados colaborativo.

Para a escala de uso pretendida por esta aplicação (equipes pequenas a médias em uma rede local), o uso destas APIs se enquadra confortavelmente na **cota de uso gratuito** oferecida pelo Google.

- **Custo Efetivo:** **Gratuito**.
- **Modelo:** O Google oferece um volume muito alto de requisições (leituras e escritas) por dia, que não deve ser atingido pelo uso normal do programa.
- **Segurança contra Cobranças:** Caso o limite da cota gratuita seja excedido, a API simplesmente para de responder temporariamente até que a cota seja renovada. **Não há cobrança automática.** A cobrança só é habilitada se o usuário do projeto no Google Cloud explicitamente adicionar informações de faturamento e solicitar um aumento de cota, um processo que não é necessário para o funcionamento deste aplicativo.

### Passo 5: Executando a Aplicação
Com tudo configurado, basta executar o script:
```bash
python gestor_estoque.py
```

## Como Compilar para `.exe` (Windows)
Para criar um executável que funcione em outros computadores sem Python, instale o PyInstaller:
```bash
pip install pyinstaller
```
Em seguida, execute o comando de compilação a partir da pasta raiz do projeto:
```bash
pyinstaller --windowed --onefile --icon="icon.ico" --add-data "credentials.json;." --add-data "relatorio_template.html;." --add-data "relatorio_setores_template.html;." gestor_estoque.py
```
O arquivo `.exe` final estará na pasta `dist`.
