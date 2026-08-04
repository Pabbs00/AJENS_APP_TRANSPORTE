# 🚌 GERADOR DE ROTEIROS - AJENS

Aplicativo em [Streamlit](https://streamlit.io/) que gera automaticamente os roteiros de transporte dos estudantes associados a partir das respostas do formulário Google **"AJENS 💙 🚌+📚=🎓"**. O sistema aloca os estudantes nos veículos disponíveis (ônibus, micro-ônibus e vans), respeitando:

- horário de retorno (21:30 ou 22:00);
- prioridade de preenchimento (ônibus e micro-ônibus primeiro);
- regras de mistura entre polos e instituições para otimizar rotas próximas;
- remanejamento automático de quem não coube no primeiro horário para o horário seguinte.

Ao final, o sistema exibe um resumo da lotação, indica se algum veículo pode ser dispensado ou se há falta de capacidade, e gera uma planilha única em `.xlsx` pronta para download.

## Funcionalidades

O aplicativo executa as seguintes funções:

- lê e processa planilhas em `.xlsx` ou `.csv` exportadas do formulário Google;
- detecta automaticamente as colunas principais ou permite remapeá-las manualmente na barra lateral;
- limpa dados incompletos, remove cancelamentos e padroniza informações como instituição, horário e uso do transporte;
- classifica os estudantes por instituição/polo para otimizar a alocação por veículo;
- aloca os estudantes nos ônibus, micro-ônibus e vans disponíveis, respeitando vagas, horários e prioridade de distribuição;
- faz o remanejamento automático dos estudantes que não couberam no horário das 21:30 para os carros das 22:00;
- marca visualmente os monitores, os estudantes remanejados e os casos de ida/retorno em destaque;
- exibe a lista de espera para os alunos remanejados do carro de 21:30 em uma aba separada da planilha;
- gera o arquivo final em Excel com nome no formato de data e ano, pronto para download.

## Instalação local

```bash
# 1. Clone o repositório
git clone <URL_DO_SEU_REPOSITORIO>
cd <NOME_DA_PASTA>

# 2. Crie e ative um ambiente virtual (recomendado)
python3 -m venv venv
source venv/bin/activate        # Linux/Mac
venv\Scripts\activate           # Windows

# 3. Instale as dependências
pip install -r requirements.txt
```

### ▶️ Como rodar

```bash
streamlit run app_transporte.py
```

O navegador abrirá automaticamente em `http://localhost:8501`.

## 📥 De onde vem a planilha

O aplicativo espera o arquivo com as **respostas do formulário Google** (planilha do Google Sheets vinculada a ele), exportado como:

- **Excel**: no Google Sheets, `Arquivo → Fazer download → Microsoft Excel (.xlsx)`; ou
- **CSV**: `Arquivo → Fazer download → Valores separados por vírgula (.csv)`.

### Estrutura de colunas esperada

O aplicativo detecta as colunas automaticamente pelo texto do cabeçalho, mas foi calibrado para a ordem atual do formulário **"AJENS 💙 🚌+📚=🎓"**:

| # | Coluna | Uso no aplicativo |
|---|--------|-------------------|
| 0 | Carimbo de data/hora | Ordenação de chegada |
| 1 | Endereço de e-mail | Não utilizada |
| 2 | Texto de aviso/ciência | Não utilizada |
| 3 | Seu nome e sobrenome? | Nome do estudante |
| 4 | Qual seu curso? | Não utilizada |
| 5 | Qual a instituição? | Identificação do polo |
| 6 | Como irá utilizar o transporte? | Ida/volta |
| 7 | Horário de retorno? | 21:30 ou 22:00 |
| 8 | Desembarque | Bairro de desembarque |

Se a ordem das perguntas do formulário mudar, use o painel **"🛠️ Ajustes de Colunas"** na barra lateral do aplicativo para remapear manualmente cada coluna; não é necessário editar o código.

## 📁 Estrutura do projeto

```text
.
├── app_transporte.py       # App principal (Streamlit)
├── requirements.txt        # Dependências Python
├── .gitignore
└── README.md
```