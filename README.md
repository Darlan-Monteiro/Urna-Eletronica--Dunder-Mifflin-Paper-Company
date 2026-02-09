# 📄 Sistema de Votação - Dunder Mifflin Paper Company Inc.

> *"Eu quero que as pessoas tenham medo do quanto elas me amam."* - Michael Scott

![Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![VBA](https://img.shields.io/badge/VBA-Visual%20Basic-blue?style=for-the-badge)
![Schrute Bucks](https://img.shields.io/badge/Currency-Schrute%20Bucks-brown?style=for-the-badge)

## 📁 Sobre o Projeto

A Corporate (David Wallace) solicitou uma auditoria imediata no processo de escolha do novo **Gerente Regional** da filial de Scranton. O método antigo (papéis jogados num chapéu) foi considerado inseguro devido a "incidentes" envolvendo o Creed Bratton.

Eu desenvolvi esta **Urna Eletrônica em Excel/VBA** para garantir uma eleição justa, transparente e livre de fraudes (exceto se o Ryan tentar invadir o sistema).

O sistema inclui uma interface de votação segura (UserForm), reconhecimento de candidatos com fotos e um Dashboard Administrativo em tempo real que deixaria o Oscar Martinez orgulhoso.

## 📸 Screenshots (Provas)

### A Urna (Interface do Usuário)
*Desenvolvida para ser tão simples que até o Kevin consegue usar.*
<img width="1875" height="550" alt="image" src="https://github.com/user-attachments/assets/2daa478f-0a8e-4f3d-a4e8-e097450beedc" />
<img width="1105" height="980" alt="image" src="https://github.com/user-attachments/assets/921646d5-3c3e-4007-aa42-db29c29ee899" />



### O Dashboard (Modo "Dark Mode" Executivo)
*Análise de dados em tempo real. Inclui auditoria de votos nulos para detectar sabotagens do Toby.*
<img width="1698" height="691" alt="image" src="https://github.com/user-attachments/assets/d94938ca-db44-4e7a-b1c3-667719f4986b" />


## 🚀 Funcionalidades

- **Segurança Anti-Toby:** O sistema roda em uma janela modal (UserForm), impedindo que usuários alterem as células manualmente durante a votação.
- **Reconhecimento Facial (Via Arquivo):** Digite **31** e o sistema carrega a foto do Michael Scott. Digite **35** e ele carrega o Dwight.
- **Tratamento de Erros:** O código é robusto o suficiente para aceitar arquivos `.jpg`, `.jpeg`, `.bmp` e `.gif`. Se a foto não existir, o sistema não trava (porque travamentos diminuem a produtividade).
- **Log de Votação:** Cada voto é registrado com um carimbo de data/hora (Timestamp) preciso. Sabemos exatamente quem votou logo depois do almoço.
- **Dashboard Automático:**
  - 📊 **Ranking em Tempo Real:** Atualizado via Macro.
  - 📈 **Análise de Produtividade:** Gráfico de linha mostrando horários de pico.
  - 🍩 **Auditoria:** Gráfico de Rosca para monitorar votos brancos e nulos.

