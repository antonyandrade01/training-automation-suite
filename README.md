# Training Automation Suite

![Python](https://img.shields.io/badge/Python-3.7%2B-blue?style=for-the-badge&logo=python)
![MySQL](https://img.shields.io/badge/MySQL-4479A1?style=for-the-badge&logo=mysql&logoColor=white)
![PowerPoint](https://img.shields.io/badge/PowerPoint-B7472A?style=for-the-badge&logo=microsoftpowerpoint&logoColor=white)
![LibreOffice](https://img.shields.io/badge/LibreOffice-18A303?style=for-the-badge&logo=libreofficet&logoColor=white)
![License](https://img.shields.io/github/license/antonyandrade01/training-automation-suite?style=for-the-badge)

<!-- English Version (Default) -->
<div align="center">

### 🇬🇧 English Version

A powerful Python-based automation suite designed to streamline and accelerate the creation of technical training materials and the updating of business support tickets. This tool transforms a manual, multi-hour process into a fast, consistent, and error-free workflow.
</div>

#### The Problem: The Manual Bottleneck

In many companies, creating release training presentations and updating associated support tickets is a significant operational bottleneck. The process often involves:
*   Manually querying databases to get task lists.
*   Manually searching for screenshots and assets in network folders.
*   Painstakingly creating dozens of PowerPoint slides, copying and pasting information.
*   Manually updating each corresponding support ticket with the release version.

This process is not only slow but also highly susceptible to human error.

#### ✨ The Solution: An Automated Pipeline

This suite provides a command-line interface (CLI) to automate the entire workflow:

1.  **Project Verification:** Connects to a MySQL database to verify project tasks, identifying missing folders or assets and generating a clear discrepancy report.
2.  **CSV Generation:** Automatically queries the project database to generate a structured CSV file containing all necessary ticket information.
3.  **Automated Presentation Generation:**
    *   Reads the generated CSV and a base PowerPoint template (`.pptx`).
    *   Dynamically creates a new slide for each task.
    *   Automatically finds, resizes, and embeds the corresponding screenshots.
    *   Applies complex, multi-step animations to the images on each slide, ready for presentation.
4.  **Bulk Ticket Updates:** Uses a support desk API to batch-update all tickets listed in the CSV, posting a standardized action with the new release version.

#### 🚀 How to Use

*(Note: This suite is tailored to a specific business environment but can be adapted by changing the configurations in the `.env` file.)*

1.  **Setup:**
    *   Clone the repository: `git clone https://github.com/antonyandrade01/training-automation-suite.git`
    *   Install dependencies: `pip install -r requirements.txt`
    *   Copy `.env.example` to a new file named `.env` and fill it with your real credentials (database, API token, paths, etc.).
    *   Place your PowerPoint template file as `Layout-Base.pptx` in the root directory.

2.  **Run the application:**
    *   Execute the main script from your terminal: `python main.py`
    *   Follow the interactive menu to choose the desired action.

---
<!-- Collapsible Portuguese Version -->
<details align="center">
  <summary><b>🇧🇷 Clique aqui para ver a versão em Português</b></summary>
  
  ### 🇧🇷 Versão em Português

  <p>Uma poderosa suíte de automação em Python, projetada para otimizar e acelerar a criação de materiais de treinamento técnico e a atualização de tickets de suporte. Esta ferramenta transforma um processo manual de várias horas em um fluxo de trabalho rápido, consistente e livre de erros.</p>

  <h4>O Problema: O Gargalo Manual</h4>
  <p>Em muitas empresas, a criação de apresentações de treinamento e a atualização dos tickets de suporte associados é um gargalo operacional. O processo frequentemente envolve:</p>
  <ul>
    <li>Consultar manualmente bancos de dados.</li>
    <li>Procurar manualmente por screenshots em pastas de rede.</li>
    <li>Criar dezenas de slides no PowerPoint, copiando e colando informações.</li>
    <li>Atualizar manualmente cada ticket de suporte com a versão do lançamento.</li>
  </ul>
  <p>Este processo é lento e altamente suscetível a erros humanos.</p>
  
  <h4>✨ A Solução: Um Pipeline Automatizado</h4>
  <p>Esta suíte oferece uma interface de linha de comando (CLI) para automatizar todo o fluxo de trabalho:</p>
  <ol>
    <li><strong>Verificação de Projeto:</strong> Conecta-se a um banco de dados MySQL para verificar tarefas, identificando pendências e gerando um relatório de discrepâncias.</li>
    <li><strong>Geração de CSV:</strong> Consulta o banco de dados para gerar um arquivo CSV estruturado com as informações dos tickets.</li>
    <li><strong>Geração Automatizada de Apresentação:</strong>
      <ul>
        <li>Lê o CSV e um template de PowerPoint (<code>.pptx</code>).</li>
        <li>Cria dinamicamente um slide para cada tarefa.</li>
        <li>Encontra, redimensiona e insere automaticamente os screenshots.</li>
        <li>Aplica animações complexas e sequenciais às imagens em cada slide.</li>
      </ul>
    </li>
    <li><strong>Atualizações de Ticket em Massa:</strong> Usa a API de um sistema de suporte para atualizar em lote todos os tickets listados no CSV com uma ação padronizada.</li>
  </ol>
</details>

---

### License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.