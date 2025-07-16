# LeadAnalyzerOffline

## Descrição

O **LeadAnalyzerOffline** é uma aplicação desktop desenvolvida em Python para análise de leads, vendas e eficiência de canais de captação. Com uma interface moderna e intuitiva, permite importar planilhas do Excel, filtrar dados por cidade, período e origem, visualizar dashboards interativos e exportar relatórios completos em PDF ou Excel.

## Funcionalidades

- **Importação de Planilhas Excel**: Carregue arquivos `.xlsx` ou `.xls` com múltiplas abas (cada aba representa uma cidade).
- **Filtros Avançados**: Filtre por cidade, intervalo de períodos e origens de leads.
- **Dashboard Interativo**: Visualize métricas principais, gráficos de conversão, evolução mensal, eficiência de vendas, correlação e dispersão entre leads e vendas.
- **Exportação de Relatórios**: Gere relatórios completos em PDF ou exporte tabelas para Excel.
- **Pesquisa e Seleção de Origens**: Pesquise e selecione múltiplas origens de leads de forma prática.
- **Interface Moderna**: Desenvolvido com CustomTkinter para uma experiência visual agradável.

## Como Usar

1. **Pré-requisitos**  
   - Python 3.8 ou superior
   - Instale as dependências com:
     ```bash
     pip install -r requirements.txt
     ```

2. **Executando o Sistema**
   - Navegue até a pasta `LeadAnalyzerOffline` e execute:
     ```bash
     python main.py
     ```

3. **Passos no Sistema**
   - Clique em **Carregar Planilha** e selecione seu arquivo Excel.
   - Utilize os filtros à esquerda para escolher cidade, período e origens.
   - Navegue entre as visualizações no menu "Tipo de Visualização".
   - Exporte relatórios em PDF ou tabelas em Excel conforme necessário.

## Estrutura Esperada da Planilha

- Cada aba representa uma cidade (exceto abas ignoradas como "Salvador", "Planilha1", etc.).
- As colunas principais devem incluir:  
  - `periodo` (ou variações como "período")
  - `contatos`
  - `aproveitados`
  - `vendas`
  - `origem`

## Principais Visualizações

- **Visão Geral**: Métricas totais e gráficos de distribuição.
- **Desempenho por Origem**: Tabela detalhada por canal.
- **Conversão por Canal**: Comparativo de taxas de conversão.
- **Evolução Mensal**: Tendências ao longo do tempo.
- **Top Canais**: Canais mais eficientes.
- **Eficiência de Vendas**: Vendas por lead.
- **Correlação Leads-Vendas**: Relação estatística entre leads e vendas.
- **Dispersão Leads x Vendas**: Gráficos de dispersão para análise visual.

## Dependências

Veja o arquivo `requirements.txt` para a lista completa. Principais bibliotecas:
- pandas
- numpy
- customtkinter
- matplotlib
- seaborn
- reportlab

## Observações

- O sistema é totalmente offline, não requer conexão com a internet.
- Relatórios gerados são salvos localmente no formato escolhido pelo usuário.

## Licença

Este projeto é de uso interno/confidencial. Consulte o responsável pelo sistema para mais informações.