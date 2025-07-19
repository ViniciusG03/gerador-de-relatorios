# 🏥 Gerador de Relatórios FUSEX v2.0

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://python.org)
[![Tkinter](https://img.shields.io/badge/GUI-Tkinter-green.svg)](https://docs.python.org/3/library/tkinter.html)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Release](https://img.shields.io/github/v/release/ViniciusG03/gerador-de-relatorios)](https://github.com/ViniciusG03/gerador-de-relatorios/releases)

Um aplicativo desktop para automação da geração de relatórios médicos padronizados para o convênio FUSEX, desenvolvido em Python com interface gráfica intuitiva. Agora na versão 2.0 com suporte expandido para mais especialidades.

## ✨ Funcionalidades

- 🧩 **Relatórios PNE** - Para Portadores de Necessidades Especiais
- 👤 **Relatórios Típicos** - Para desenvolvimento típico
- 📊 **Importação Excel** - Carrega dados de planilhas `.xlsx` e `.xls`
- 👁️ **Preview dos Dados** - Visualização prévia antes da geração
- 🎯 **Múltiplas Especialidades** - Suporte expandido para:
  - Terapia ABA
  - Psicoterapia
  - Terapia Ocupacional
  - Fonoaudiologia
  - Psicomotricidade
  - Psicopedagogia
  - **🆕 Nutrição** - Relatórios especializados para terapia alimentar
  - **🆕 Fisioterapia** - Evolução e programação fisioterapêutica
- 📄 **Papel Timbrado** - Documentos com formatação profissional
- 🚀 **Geração em Lote** - Múltiplos relatórios de uma vez
- **🆕 Sistema de Assinaturas** - Assinaturas digitais específicas por especialidade
- **🆕 Listas com Formatação** - Bullets e formatação aprimorada nos relatórios

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
pip install -r requirements.txt
```

**Dependências principais:**

- `pandas` - Manipulação de dados Excel
- `python-docx` - Geração de documentos Word
- `openpyxl` - Leitura de arquivos Excel
- `pytest` - Testes unitários (opcional)

## 🚀 Instalação e Uso

### Opção 1: Executável (Recomendado)

1. Acesse a página de [Releases](https://github.com/ViniciusG03/gerador-de-relatorios/releases)
2. Baixe a versão mais recente (`GERADOR_FUSEX_v2.0.zip`)
3. Extraia o arquivo em qualquer pasta
4. Execute `Gerador_Relatorios_FUSEX_v2.0.exe`

### Opção 2: Código Fonte

```bash
# Clone o repositório
git clone https://github.com/ViniciusG03/gerador-de-relatorios.git
cd gerador-de-relatorios

# Instale as dependências
pip install -r requirements.txt

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
Pedro Oliveira     | 10/12/2018        | Lucia Oliveira     | Fisioterapia       | Janeiro/2025
Sofia Mendes       | 05/07/2020        | Rafael Mendes      | Nutrição           | Janeiro/2025
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
├── 📄 Relatório_PNE_Pedro_Oliveira.docx
├── 📄 Relatório_PNE_Sofia_Mendes.docx
└── 📄 Relatório_Típico_Carlos_Mendes.docx
```

## 🔧 Estrutura do Projeto

```
gerador-de-relatorios/
├── 📄 main.py                    # Código principal
├── 📄 requirements.txt           # Dependências do projeto
├── 📄 fusex_tipico.xlsx          # Planilha de exemplo
├── 📄 papel timbrado.docx        # Template do papel timbrado
├── � gerador_relatorios/        # Módulo principal
│   ├── 📄 __init__.py
│   ├── 📄 data_loader.py         # Carregamento de dados Excel
│   ├── 📄 gui.py                 # Interface gráfica
│   ├── 📄 reports.py             # Geração de relatórios
│   └── 📄 utils.py               # Utilitários
├── 📂 tests/                     # Testes unitários
│   ├── 📄 test_data_loader.py
│   └── 📄 test_reports.py
├── 📂 build/                     # Arquivos de build
├── 📂 dist/                      # Executável gerado
└── 📂 GERADOR_FUSEX_v2.0/        # Versão distribuição
    ├── 📄 Gerador_Relatorios_FUSEX_v2.0.exe
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

Adicione novas especialidades modificando o dicionário `SIGNATURES` no arquivo `reports.py`:

```python
SIGNATURES = {
    "NOVA_ESPECIALIDADE": [
        ("assinado eletronicamente", {"italic": True}),
        ("NOME DO PROFISSIONAL", {"bold": True}),
        ("Titulo Profissional", {}),
        ("Numero Registro", {})
    ]
}
```

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

**❌ "Fusex Típico não contempla NUTRIÇÃO/FISIOTERAPIA"**

- Estas especialidades estão disponíveis apenas em relatórios PNE
- Para Fusex Típico, use outras especialidades disponíveis

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

### v2.0 (Atual) - Novembro 2024

**🆕 Novas Funcionalidades:**

- ✅ Suporte completo para **Fisioterapia** (apenas PNE)
- ✅ Suporte completo para **Nutrição** (apenas PNE)
- ✅ Sistema de assinaturas digitais por especialidade
- ✅ Listas com bullets e formatação aprimorada
- ✅ Evolução específica para cada especialidade
- ✅ Melhoria na estrutura de código com módulos separados
- ✅ Sistema de testes automatizados

**🔧 Melhorias:**

- ✅ Refinamento na geração de relatórios PNE
- ✅ Validação de especialidades por tipo de relatório
- ✅ Interface mais robusta e informativa
- ✅ Documentação expandida

### v1.0 - Lançamento Inicial

- ✅ Interface gráfica intuitiva
- ✅ Geração de relatórios PNE e Típicos
- ✅ Suporte a 6 especialidades básicas
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

- 🐛 Abra uma [Issue](https://github.com/ViniciusG03/gerador-de-relatorios/issues)
- � Baixe a versão mais recente nas [Releases](https://github.com/ViniciusG03/gerador-de-relatorios/releases)
- �💬 Entre em contato via GitHub

### 📊 Estatísticas do Projeto

- 🎯 **8 Especialidades** suportadas
- 📋 **2 Tipos** de relatórios (PNE e Típico)
- 🧪 **Testado** com pytest
- 📦 **Empacotado** com PyInstaller

---

## 📦 Instalação Rápida

Para começar a usar o projeto rapidamente:

```bash
# Clonar o repositório
git clone https://github.com/ViniciusG03/gerador-de-relatorios.git
cd gerador-de-relatorios

# Instalar dependências
pip install -r requirements.txt

# Executar aplicativo
python main.py
```

<div align="center">
  <p>⭐ Se este projeto foi útil para você, considere dar uma estrela!</p>
  <p>Feito com ❤️ para automatizar relatórios médicos</p>
  <p><strong>Versão 2.0 - Julho 2025</strong></p>
</div>
