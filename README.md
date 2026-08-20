# PyBank

Lê faturas de cartão de crédito em **PDF** de bancos brasileiros, extrai as
transações e gera uma planilha **Excel formatada** (`Saidas/Faturas.xlsx`), com
uma aba por banco. Detecta o banco automaticamente, concilia o valor extraído
contra o total real da fatura e registra tudo em log.

---

## Funcionalidades

- **Autodetecção do banco** — identifica de qual banco é cada PDF pelo conteúdo,
  sem precisar configurar nada.
- **Conciliação** — compara a soma extraída com o "total a pagar" real do PDF e
  escreve o resultado na planilha (`Tudo certo`, `Falta`, `Sobra` ou
  `Impossível conferir`).
- **Datas corrigidas** — o ano vem da data de emissão do PDF (não do relógio),
  com correção automática da virada de ano (compra de dezembro numa fatura de
  janeiro sai com o ano certo).
- **Excel pronto pra ler** — fórmulas, formato monetário `R$`, datas
  `DD/MM/YYYY`, fontes e largura de coluna ajustada automaticamente.
- **Tolerante a falhas** — um PDF que não é fatura é ignorado com aviso; uma
  fatura sem total é processada mesmo assim e sinalizada. Nada interrompe o lote.
- **Log completo** — histórico e tempo de execução em `Logs/Faturas.log`.

## Bancos suportados

Nubank · Banco Inter · Sicredi · Banco do Brasil

---

## Demonstração

Rodando nas faturas de exemplo, o script gera a planilha formatada e registra
cada etapa no log.

**Planilha gerada** — uma aba por banco, com fórmulas, formato `R$`, datas
padronizadas, largura de coluna automática e a conferência na coluna lateral:

![Planilha de saída gerada pelo PyBank](assets/Exemplo_saida.png)

**Log da execução** — identificação do banco, extração, conciliação e tempo total. Repare no `WARNING`: um PDF **intruso** que não é fatura (`documento_aleatorio.pdf`) é reconhecido e **ignorado sem interromper** o processamento dos outros — nada derruba o lote:

![Log de execução do PyBank](assets/Exemplo_log.png)

---

## Instalação

Requer **Python 3.8+**. Instale as dependências:

```bash
pip install pdfplumber pandas openpyxl
```

## Como usar

1. Coloque os PDFs das faturas na pasta `Faturas/`.
2. Rode:

   ```bash
   python main.py
   ```

3. A planilha sai em `Saidas/Faturas.xlsx` (uma aba por banco) e o histórico em
   `Logs/Faturas.log`.

As pastas `Saidas/` e `Logs/` são criadas automaticamente se não existirem.

## Estrutura de pastas

```
PyBank/
├── main.py            # todo o pipeline
├── Faturas/           # entrada: PDFs das faturas (você coloca aqui)
├── Faturas_teste/     # faturas fictícias de exemplo (uma por banco)
├── assets/            # prints usados no README
├── Saidas/            # saída: Faturas.xlsx (gerado)
└── Logs/              # saída: Faturas.log (gerado)
```

---

## Dados de exemplo

A pasta `Faturas_teste/` traz **4 faturas fictícias** (uma por banco), com
lojas, valores e datas totalmente inventados — nenhum dado real. Servem pra
testar ou demonstrar o projeto sem precisar de uma fatura de verdade.

Pra rodar nelas, aponte a variável `pasta_faturas` (no `main.py`) para
`Faturas_teste/`, ou copie os PDFs para `Faturas/`. As quatro conciliam
certinho (`Tudo certo` no log), incluindo os encargos do Nubank e do Sicredi.

---

## Como funciona

Pipeline em uma passada por PDF:

```
Faturas/*.pdf
     │
     ▼
identificador   →  qual banco? (âncora única no texto da 1ª página)
     │
     ▼
parser do banco →  extrai transações, encargos e o total real
     │              (ano vindo do metadado do PDF + correção dez/jan)
     ▼
preparar_df     →  DataFrame padronizado, valores convertidos pra número
     │
     ▼
conferir        →  soma extraída × total real  →  aviso na célula K1 da aba
     │
     ▼
exportar_xlsx   →  grava cada banco numa aba de Saidas/Faturas.xlsx
     │
     ▼
manipular_xlsx  →  fórmulas, fontes, formato R$/data, autofit das colunas
```

### Colunas da planilha

| Coluna | Origem |
|---|---|
| Data | extraída da fatura |
| Titular | preenchimento manual |
| Descrição | extraída da fatura |
| Parcela | extraída da fatura (`1/1` quando não há) |
| Valor | extraído da fatura |
| Valor pago | preenchimento manual |
| Valor restante | fórmula: `Valor − Valor pago` |

---

## Observações e limitações

- **Cada banco tem seu próprio formato de PDF.** Os padrões de extração são
  específicos do layout de cada um — mexer no parser de um banco não afeta os
  outros.
- **A planilha é um ponto de partida.** As colunas *Titular* e *Valor pago* vêm
  vazias, pra você preencher à mão; *Valor restante* se calcula sozinha.
- **A conciliação reporta, não decide.** Uma diferença entre o extraído e o
  total pode ser esperada (créditos, estornos, saldo anterior). O aviso serve
  pra você conferir, não é necessariamente um erro.
- **Faturas contêm dados pessoais** (CPF, transações). Não versione a pasta
  `Faturas/` nem a saída ao subir para um repositório público.

## Diagnóstico

Se algo não bater, o primeiro lugar a olhar é o `Logs/Faturas.log` — todo o
fluxo (leituras, extrações, conciliações, avisos e erros) é registrado lá.
