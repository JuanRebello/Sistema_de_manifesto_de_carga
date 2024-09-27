
# Sistema de Manifesto de carga

Este projeto é uma aplicação para gerar Manifesto de carga, onde também é feito uma verredura em Site interno da Força aérea Brasileira.
## Funcionalidades

- **O app.py realiza o Scrapy:** Realiza um scrapy em um sistema web da FAB e guarda os dados pré-definidos em um arquivo ".xlsx". Porém esse site não é acessível, pois se trata de um sistema interno da FAB
- **Gerador de Manifesto de Carga:** Com uma interface gráfica, o script utiliza os dados gerado do scrapy do app.py e gera o arquivo Manifesto Carga


## Tecnologias Utilizadas

- **Python:** Linguagem de programação utilizada para desenvolver o gerador de senhas.
- **Scrapy:** Biblioteca para realizar varredura em um site
- **Selenium:** Biblioteca para automações de sites
- **Tkinter:** Usado para interface gráfica
## Como Executar

1. **Clone o Repositório:**

   ```bash
   git clone https://github.com/seuusuario/seurepositorio.git
   
2. **Navegue para o Diretório do Projeto:**

    ```bash
    cd seurepositorio

3. **Instale as bibliotecas:**

   ```bash
   pip install selenium scrapy pandas pywin32

    
5. **Execute o Script:**

    ```bash
    python SilomsOff.py
    

# Exemplo de Uso

  1. Selecione o PCAN.
  2. Em seguida, clique em adicionar volume.
  3. Escolha os volumes que desejar inserir no manifesto e clique em salvar selecionados.
  4. Escolha a Aeronave.
  5. Selecione o número de cópias que deseja imprimir
  6. clique em imprimir

# Contribuição

  Sinta-se à vontade para contribuir para este projeto! Você pode abrir uma issue ou enviar um pull request com melhorias e sugestões.


# Contato
  Se você tiver alguma dúvida ou sugestão, sinta-se à vontade para entrar em contato:

  Email: juan.pm.rebello@outlook.com
  GitHub: JuanRebello
