# 🏥 Gerador de Relatórios FUSEX

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://python.org)
[![Tkinter](https://img.shields.io/badge/GUI-Tkinter-green.svg)](https://docs.python.org/3/library/tkinter.html)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

Um aplicativo desktop para automação da geração de relatórios médicos padronizados para o convênio FUSEX, desenvolvido em Python com interface gráfica intuitiva.

## ✨ Funcionalidades

- 🧩 **Relatórios PNE** - Para Portadores de Necessidades Especiais
- 👤 **Relatórios Típicos** - Para desenvolvimento típico
- 📊 **Importação Excel** - Carrega dados de planilhas `.xlsx` e `.xls`
- 👁️ **Preview dos Dados** - Visualização prévia antes da geração
- 🎯 **Múltiplas Especialidades** - Suporte para:
  - Terapia ABA
  - Psicoterapia
  - Terapia Ocupacional
  - Fonoaudiologia
  - Psicomotricidade
  - Psicopedagogia
- 📄 **Papel Timbrado** - Documentos com formatação profissional
- 🚀 **Geração em Lote** - Múltiplos relatórios de uma vez

## 🛠️ Tecnologias Utilizadas

- **Python 3.8+**
- **Tkinter** - Interface gráfica
- **pandas** - Manipulação de dados Excel
- **python-docx** - Geração de documentos Word
- **PyInstaller** - Empacotamento do executável

## 📋 Pré-requisitos

### Para usar o executável:

- Windows 7 ou superior
- Nenhuma instalação adicional necessária

### Para desenvolvimento:

```bash
Python 3.8+
pip install pandas python-docx openpyxl
```

## 🚀 Instalação e Uso

### Opção 1: Executável (Recomendado)

1. Baixe o arquivo `GERADOR_FUSEX_v1.0.zip` da seção [Releases](../../releases)
2. Extraia o arquivo em qualquer pasta
3. Execute `Gerador_Relatorios_FUSEX_v1.0.exe`

### Opção 2: Código Fonte

```bash
# Clone o repositório
git clone https://github.com/ViniciusG03/gerador-de-relatorios.git
cd gerador-de-relatorios

# Instale as dependências
pip install pandas python-docx openpyxl

# Execute o aplicativo
python main.py
```

## 📊 Formato da Planilha Excel

A planilha deve conter as seguintes colunas obrigatórias:

| Coluna                 | Descrição                  | Exemplo            |
| ---------------------- | -------------------------- | ------------------ |
| **NOME**               | Nome completo do paciente  | João Silva Santos  |
| **DATA DE NASCIMENTO** | Data no formato DD/MM/AAAA | 15/03/2010         |
| **RESPONSÁVEL**        | Nome do responsável        | Maria Silva Santos |
| **ESPECIALIDADE**      | Área de atendimento        | Psicoterapia       |
| **MÊS DE REFERÊNCIA**  | Período do relatório       | Janeiro/2025       |

### Exemplo de Planilha:

```
NOME                | DATA DE NASCIMENTO | RESPONSÁVEL        | ESPECIALIDADE      | MÊS DE REFERÊNCIA
João Silva Santos  | 15/03/2010        | Maria Silva Santos | Psicoterapia       | Janeiro/2025
João Silva Santos  | 15/03/2010        | Maria Silva Santos | Terapia ABA        | Janeiro/2025
Ana Costa Lima     | 22/08/2015        | Carlos Costa Lima  | Fonoaudiologia     | Janeiro/2025
```

## 🎯 Como Usar

### Passo a Passo:

1. **Inicie o aplicativo**

   - Execute o arquivo `.exe` ou `python main.py`

2. **Selecione o tipo de relatório**

   - 🧩 **PNE**: Para pacientes com necessidades especiais
   - 👤 **Típico**: Para desenvolvimento típico

3. **Carregue os dados**

   - 📁 **Selecionar Arquivo**: Escolha sua planilha Excel
   - 📄 **Usar Exemplo**: Teste com dados de exemplo

4. **Visualize o preview**

   - Confira os dados carregados na tabela

5. **Gere os relatórios**
   - Clique em "🚀 Gerar Relatórios"
   - Escolha a pasta de destino
   - Aguarde o processamento

## 📁 Estrutura dos Arquivos Gerados

```
📂 Pasta_Escolhida/
├── 📄 Relatório_PNE_João_Silva_Santos.docx
├── 📄 Relatório_PNE_Ana_Costa_Lima.docx
└── 📄 Relatório_Típico_Carlos_Mendes.docx
```

## 🔧 Estrutura do Projeto

```
gerador-de-relatorios/
├── 📄 main.py                    # Código principal
├── 📄 fusex_tipico.xlsx          # Planilha de exemplo
├── 📄 papel timbrado.docx        # Template do papel timbrado
├── 📄 Gerador_Relatorios_FUSEX_v1.0.spec  # Config PyInstaller
├── 📂 build/                     # Arquivos de build
├── 📂 dist/                      # Executável gerado
└── 📂 GERADOR_FUSEX_v1.0/        # Versão distribuição
    ├── 📄 Gerador_Relatorios_FUSEX_v1.0.exe
    ├── 📄 fusex_exemplo.xlsx
    └── 📄 LEIA-ME.txt
```

## ⚙️ Personalização

### Modelos de Relatório

Os modelos podem ser personalizados editando as funções:

- `generate_pne_report()` - Relatórios PNE
- `generate_tipico_report()` - Relatórios Típicos

### Papel Timbrado

Substitua o arquivo `papel timbrado.docx` por seu modelo personalizado.

### Especialidades

Adicione novas especialidades modificando as verificações de especialidade no código.

## 🐛 Solução de Problemas

### Problemas Comuns:

**❌ "Colunas obrigatórias não encontradas"**

- Verifique se sua planilha possui todas as colunas necessárias
- Certifique-se de que os nomes das colunas estão exatos

**❌ "Erro ao carregar arquivo"**

- Verifique se o arquivo Excel não está aberto em outro programa
- Confirme se o formato é `.xlsx` ou `.xls`

**❌ "Erro ao gerar relatórios"**

- Verifique se há espaço suficiente na pasta de destino
- Confirme se você tem permissões de escrita na pasta

### Logs de Erro:

Os erros são exibidos em janelas de diálogo. Para mais detalhes, execute pelo terminal:

```bash
python main.py
```

## 🤝 Contribuindo

1. Faça um Fork do projeto
2. Crie uma branch para sua feature (`git checkout -b feature/nova-funcionalidade`)
3. Commit suas mudanças (`git commit -am 'Adiciona nova funcionalidade'`)
4. Push para a branch (`git push origin feature/nova-funcionalidade`)
5. Abra um Pull Request

## 📝 Changelog

### v1.0 (Atual)

- ✅ Interface gráfica intuitiva
- ✅ Geração de relatórios PNE e Típicos
- ✅ Suporte a múltiplas especialidades
- ✅ Importação de planilhas Excel
- ✅ Preview de dados
- ✅ Geração em lote
- ✅ Papel timbrado personalizado

## 📄 Licença

Este projeto está licenciado sob a Licença MIT - veja o arquivo [LICENSE](LICENSE) para detalhes.

## 👨‍💻 Autor

**Vinícius G03**

- GitHub: [@ViniciusG03](https://github.com/ViniciusG03)

## 📞 Suporte

Para dúvidas, problemas ou sugestões:

- 🐛 Abra uma [Issue](../../issues)
- 💬 Entre em contato via GitHub

---

<div align="center">
  <p>⭐ Se este projeto foi útil para você, considere dar uma estrela!</p>
  <p>Feito com ❤️ para automatizar relatórios médicos</p>
</div>
