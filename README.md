# 🤖 Automação de Leitura e Organização de Notas Fiscais XML

Sistema desenvolvido em Python que automatiza todo o processo de extração, organização e formatação de dados de Notas Fiscais Eletrônicas (NF-e) no formato XML.

💡 **Problema resolvido**: Eliminei o trabalho manual, repetitivo e sujeito a erros de digitação, que antes levava horas, e transformei em um processo de segundos.

---

## 🚀 Funcionalidades
✅ Leitura automática de todos os arquivos XML de uma pasta
✅ Compatibilidade com QUALQUER estrutura ou versão de NF-e
✅ Extração inteligente dos dados principais:
   - Número da Nota
   - Data de Emissão
   - Nome do Fornecedor
   - Valor Total
   - Dados de parcelamento, vencimento e cálculo de dias para pagamento
✅ Geração de planilha Excel profissional, já formatada
✅ Formatação de valores em Real (R$)
✅ Cores, bordas e alinhamento automático
✅ Tratamento de erros: não trava, avisa apenas o necessário e lê tudo o que for válido
✅ Interface simples: basta clicar 1 vez para rodar

---

## 🛠️ Tecnologias Utilizadas
- **Python 3.x**
- `xmltodict` → Para leitura e interpretação de arquivos XML
- `openpyxl` → Para criação e formatação avançada de planilhas Excel
- Manipulação de arquivos e sistema de pastas
- Tratamento de exceções e erros

---

## 📁 Estrutura do Projeto
📂 automacao-notas-fiscais-xml
┣ 📄 robo_nf.py → Código principal do sistema
┣ 📄 RODAR.bat → Arquivo de execução (1 clique)
┗ 📂 notas/ → Pasta onde você coloca os arquivos XML


---

## ▶️ Como usar
1. Instale o Python no computador
2. Coloque todos os arquivos XML dentro da pasta `notas`
3. Dê 2 cliques no arquivo `RODAR.bat`
4. Pronto! A planilha `Planilha_Notas_Fiscais.xlsx` será criada na mesma pasta

---

## 📊 Resultados Obtidos
- Redução de **95% do tempo** gasto no processo
- Eliminação total de erros de digitação
- Padronização dos dados para uso contábil, financeiro ou de estoque
- Organização rápida e confiável

---

Desenvolvido por **Clayton Andrade**
