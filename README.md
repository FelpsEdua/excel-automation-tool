# Excel Automation Tool

Automatizador de planilhas desenvolvido em Python para limpeza, organização e análise de dados.

## 🚀 Funcionalidades

- Leitura de arquivos CSV e Excel
- Remoção de linhas duplicadas
- Limpeza de espaços em campos de texto
- Preenchimento de valores vazios
- Cálculo automático de faturamento
- Geração de relatório Excel automatizado
- Análise por produto
- Análise por cliente
- Dashboard com indicadores
- Formatação profissional do arquivo Excel
- Tratamento de erros

## 🛠️ Tecnologias

- Python
- Pandas
- OpenPyXL

## 📊 Dashboard

![Dashboard do sistema](imagens/dashboard.png)

O dashboard apresenta indicadores como:

- Registros originais
- Registros finais
- Duplicados removidos
- Faturamento total

## 📑 Relatório Excel

![Relatório Excel](imagens/relatorio.png)

O relatório gerado contém:

- Dados Limpos
- Resumo
- Análise por Produto
- Análise por Cliente
- Faturamento

## 📁 Estrutura do projeto

```text
excel-automation-tool/
│
├── main.py
├── requirements.txt
├── README.md
│
├── entrada/
├── saida/
│
├── imagens/
│   ├── dashboard.png
│   └── relatorio.png
│
└── exemplos/
    └── vendas.csv
```

## ⚙️ Instalação

### 1. Clone o repositório

```bash
git clone https://github.com/FelpsEdua/excel-automation-tool.git
cd excel-automation-tool
```

### 2. Instale as dependências

```bash
pip install -r requirements.txt
```

## ▶️ Como executar

Execute o programa:

```bash
python main.py
```

O sistema solicitará o caminho do arquivo CSV ou Excel que será processado.

### Exemplo

```text
exemplos/vendas.csv
```

Após o processamento, o relatório será gerado automaticamente em:

```text
saida/relatorio.xlsx
```

## 🧪 Exemplo completo

```bash
git clone https://github.com/FelpsEdua/excel-automation-tool.git
cd excel-automation-tool
pip install -r requirements.txt
python main.py
```

Quando solicitado, informe:

```text
exemplos/vendas.csv
```

## 📈 Resultado

O sistema processa os dados, remove duplicidades, aplica as regras de limpeza e gera um relatório Excel estruturado para análise.

## 💼 Aplicações

A ferramenta pode ser adaptada para:

- Organização de planilhas
- Limpeza de bases de dados
- Relatórios de vendas
- Consolidação de arquivos
- Tratamento de dados administrativos
- Processos repetitivos em Excel
- Geração automática de relatórios

## 📌 Próximas melhorias

- Interface gráfica
- Seleção de arquivos por botão
- Exportação de diferentes relatórios
- Regras de limpeza personalizadas
- Suporte a múltiplos arquivos
- Geração de executável

## 👨‍💻 Autor

Projeto desenvolvido como ferramenta de automação de processos utilizando Python.

**GitHub:** https://github.com/FelpsEdua