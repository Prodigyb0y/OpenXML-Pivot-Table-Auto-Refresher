# 📈 OpenXML Pivot Table Auto-Refresher

Uma ferramenta de engenharia de dados para configurar programaticamente o comportamento de atualização de Tabelas Dinâmicas em arquivos Excel (`.xlsx`), sem a necessidade de ter o Microsoft Excel instalado no servidor.

![Python](https://img.shields.io/badge/python-3.9%2B-blue)
![Library](https://img.shields.io/badge/lib-openpyxl-green)
![Type](https://img.shields.io/badge/type-metadata_manipulation-purple)

## 🎯 O Desafio Técnico

Diferente de bibliotecas de automação de interface (como `pywin32`), bibliotecas de manipulação de arquivos como `openpyxl` não possuem um motor de cálculo. Elas não podem recalcular os valores de uma Tabela Dinâmica porque isso exigiria replicar toda a lógica proprietária da Microsoft.

## 💡 A Solução

Este projeto atua na **Camada de Definição** do arquivo Excel. Em vez de tentar calcular os dados, ele modifica as propriedades XML da Tabela Dinâmica para ativar a flag `refreshOnLoad`.

**Resultado:** O arquivo é modificado de forma segura e, na próxima vez que qualquer usuário (ou robô) abrir este arquivo no Excel, os dados serão recalculados automaticamente.

### Principais Funcionalidades

* **Zero Dependência de Office:** Roda em servidores Linux/Docker, pois não exige instalação do Excel.
* **Safety-First:** Sistema automático de backup antes de qualquer modificação no binário do arquivo.
* **Logging Granular:** Rastreabilidade total das ações para auditoria.

## 🛠️ Instalação

```bash
pip install openpyxl
