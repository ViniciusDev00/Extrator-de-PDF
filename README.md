# Extrator de ORFA - NSF (Processamento Inteligente)

![Logo NSF](Logo_NSF_C.png)

O **Extrator de ORFA** é uma ferramenta robusta e intuitiva desenvolvida para automatizar a extração de dados técnicos (como especificações de laterais e materiais) de arquivos PDF. Focada no fluxo de trabalho da **NSF**, a aplicação transforma documentos complexos em relatórios organizados de forma rápida e eficiente.

## 🚀 Funcionalidades

* **Extração Inteligente**: Identifica automaticamente números de ORFA, nomes de expositores e detalhes de laterais (Direita/Esquerda) através de processamento de texto e mapeamento de coordenadas.
* **Múltiplos Formatos de Exportação**:
    * **Excel (.xlsx)**: Ideal para manipulação de dados e integração com outros sistemas.
    * **PDF Profissional**: Gera relatórios formatados com cabeçalho personalizado da NSF, paginação e design limpo.
* **Processamento em Lote**: Permite a seleção e o processamento de múltiplos arquivos PDF simultaneamente.
* **Interface Moderna**: Design responsivo com área de upload intuitiva e feedback de status em tempo real.

## 🛠️ Tecnologias Utilizadas

Este projeto foi construído utilizando tecnologias web modernas para garantir performance sem a necessidade de um backend:

* **HTML5 & CSS3**: Estrutura e estilização avançada com variáveis CSS.
* **JavaScript (ES6+)**: Lógica principal de processamento e manipulação de arquivos.
* **PDF.js**: Biblioteca para leitura e extração de conteúdo de arquivos PDF diretamente no navegador.
* **SheetJS (XLSX)**: Utilizada para a geração de planilhas Excel.
* **jsPDF & jsPDF-AutoTable**: Bibliotecas para a criação dinâmica de relatórios em PDF com tabelas profissionais.

## 📋 Como Usar

1.  **Instalação**: Clone o repositório:
    ```bash
    git clone [https://github.com/ViniciusDev00/Extrator-de-PDF.git](https://github.com/ViniciusDev00/Extrator-de-PDF.git)
    ```
2.  **Acessar**: Abra o arquivo `index.html` em seu navegador ou acesse a [Demonstração Online](https://viniciusdev00.github.io/Extrator-de-PDF/).
3.  **Upload**: Clique na área de seleção ou arraste seus arquivos PDF.
4.  **Exportar**: Escolha entre **"Exportar Excel"** ou **"Exportar PDF"** após o processamento dos arquivos.

## 🏗️ Estrutura do Projeto

* `index.html`: Interface principal do usuário.
* `style.css`: Estilização e identidade visual da NSF.
* `script.js`: Motor de extração e lógica de exportação.
* `Logo_NSF_C.png`: Logotipo utilizado nos documentos gerados.

---
**Desenvolvido por Vinicius Biancolini**
© 2026 NSF - Extrator ORFA v1.0.0
