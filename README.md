# Bot SEI ARTESP 🤖⚖️

Este projeto é uma automação desenvolvida em **Python** para otimizar a extração de especificações de lotes dentro da plataforma **SEI (Sistema Eletrônico de Informações)** utilizada pela **ARTESP**.

O bot resolve o problema de consulta manual de centenas de processos, acessando a tela de edição de cada um e "pescando" a informação do Lote via Web Scraping.

## 🚀 Funcionalidades

* **Automação de Busca**: Realiza a pesquisa rápida de processos a partir de uma lista em `.txt`.
* **Web Scraping Robusto**: Utiliza JavaScript injetado via Selenium para navegar em estruturas complexas de *iframes* do SEI.
* **Tratamento de Exceções**: Lida automaticamente com alertas do navegador (ex: "Link sem assinatura") e quedas de sessão.
* **Relatórios Automatizados**: Gera um arquivo Excel (`.xlsx`) com as especificações extraídas, formatando as colunas automaticamente para melhor leitura.

## 🛠️ Tecnologias Utilizadas

* **Python**: Linguagem principal.
* **Selenium WebDriver**: Automação da navegação web.
* **Pandas**: Manipulação de dados e criação de relatórios.
* **XlsxWriter**: Formatação avançada de planilhas Excel.
* **Regex**: Extração precisa de padrões de texto (Lotes).

## 📋 Como utilizar

1.  Certifique-se de ter o Chrome instalado.
2.  Inicie o Chrome no modo de depuração (porta 9222).
3.  Crie um arquivo `processos.txt` na raiz do projeto com os números dos processos (um por linha).
4.  Execute o script principal:
    ```bash
    python robo_artesp.py
    ```

## 🧠 Desafios Superados

Durante o desenvolvimento, foram aplicadas técnicas avançadas para superar as limitações da plataforma SEI:
* **Mergulho em Frames**: Navegação dinâmica entre múltiplos iframes para localizar botões de ação.
* **Sincronismo de Sessão**: Implementação de resets de contexto para evitar que o servidor derrubasse o login durante consultas em massa.
* **Injeção de JS**: Uso de scripts manuais para extração de dados onde o Selenium puro encontrava barreiras de renderização.

---
Desenvolvido por **Luiz Gama** como parte de estudos de automação e cibersegurança.# bot-sei-artesp1
