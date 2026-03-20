# Cadastro Sisaeduc

Pipeline para:

1. extrair informações das fichas `.docx`
2. gerar `output/csv/students.csv`
3. preencher o formulário do Sisaeduc com `pyautogui`

## Scripts

- `scripts/extract_students.py`: lê as fichas de `MATRICULAS CLODOMIRO LEAL/` e gera o CSV.
- `scripts/fill_sisaeduc.py`: usa o CSV e a ordem definida em `config/sisaeduc_tab_order.json`.

## Uso

Gerar o CSV:

```bash
python scripts/extract_students.py
```

Testar a automação sem tocar no navegador:

```bash
python scripts/fill_sisaeduc.py --dry-run --start-row 1 --limit 1
```

Executar a automação:

```bash
python scripts/fill_sisaeduc.py --start-row 1
```

Lote supervisionado, esperando sua confirmação manual e iniciando o próximo após 4 segundos:

```bash
python scripts/fill_sisaeduc.py --start-row 1 --limit 5 --start-delay 8 --resume-delay 4
```

Salvar a posição do primeiro campo:

```bash
python scripts/fill_sisaeduc.py --save-start-position
```

Depois disso, nas execuções normais o script clica automaticamente na posição salva antes de preencher.

## Ajuste da ordem de Tab

O arquivo `config/sisaeduc_tab_order.json` contém a ordem inicial dos campos. Ajuste esse arquivo para refletir exatamente a ordem da tela do Sisaeduc.

Cada entrada aceita:

- `field_id`: identificador do campo no fluxo
- `csv_column`: coluna do CSV usada como fonte
- `action_type`: `text`, `choice` ou `skip`
- `conversions`: lista opcional de conversões simples, como `digits` ou `upper`
- `choice_fallback`: define o comportamento quando não existe mapeamento em `choices`
- `tab_after`: se deve apertar `Tab` após preencher o campo
- `extra_tabs`: quantos `Tab` extras devem ser enviados
- `pause_after`: pausa adicional após o campo

## Dependências

Para a automação:

```bash
pip install -r requirements.txt
```

O extrator de `.docx` usa apenas a biblioteca padrão do Python.
