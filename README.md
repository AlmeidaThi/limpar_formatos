# Limpeza de Formatos em Planilhas Excel

Utilitario em Python para processar arquivos `.xlsx` de uma pasta, remover formatacoes existentes das colunas e reaplicar formatos padronizados para campos de data e hora. Os arquivos tratados sao salvos em uma subpasta chamada `Limpeza_Completa`, junto com logs da execucao.

## O que o script faz

- Localiza todos os arquivos `.xlsx` na mesma pasta do script.
- Cria automaticamente a pasta `Limpeza_Completa`.
- Abre cada planilha usando o Microsoft Excel via automacao COM.
- Detecta a linha de cabecalho analisando as primeiras linhas da aba.
- Desfaz mesclagens no intervalo usado da planilha.
- Limpa os formatos das colunas.
- Identifica colunas de data, hora, identificacao, texto e numericas.
- Aplica formato `dd/mm/aaaa` ou `dd/mm/aaaa hh:mm:ss` para datas.
- Aplica formato `hh:mm:ss` para colunas de hora.
- Salva uma copia processada de cada arquivo na pasta de destino.
- Gera log em texto e log detalhado em Excel.

## Estrutura esperada

```text
.
+-- _limparformatos.py
+-- _Executa_limparformatos.bat
+-- arquivo_1.xlsx
+-- arquivo_2.xlsx
+-- Limpeza_Completa/
    +-- arquivos_processados.xlsx
    +-- _log_execucao.txt
    +-- _log_execucao_detalhado.xlsx
```

## Requisitos

- Windows.
- Microsoft Excel instalado.
- Python 3.
- Bibliotecas Python:
  - `pandas`
  - `tqdm`
  - `pywin32`
  - `openpyxl`

## Como usar

1. Coloque o script `_limparformatos.py` na mesma pasta dos arquivos `.xlsx` que serao processados.
2. Execute pelo terminal:

```bash
python _limparformatos.py
```

Ao final, os arquivos processados estarao na pasta `Limpeza_Completa`.

## Logs

O processamento gera dois arquivos de log:

- `_log_execucao.txt`: resumo da execucao, arquivos processados e mensagens de erro.
- `_log_execucao_detalhado.xlsx`: detalhes por arquivo, aba, coluna, tipo identificado, acao aplicada e motivo.

## Observacoes

- O script processa apenas arquivos `.xlsx`.
- Os arquivos originais nao sao sobrescritos; as copias tratadas sao salvas em `Limpeza_Completa`.
- O Excel e aberto em segundo plano durante a execucao.
- Se algum arquivo estiver aberto ou protegido, o processamento pode falhar para esse arquivo e o erro sera registrado no log.

