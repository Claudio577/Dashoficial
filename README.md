# Dashboard Executivo de Suporte

Dashboard em Streamlit para consolidar relatórios mensais de chamados e gerar uma visão executiva anual.

## Como usar

1. Coloque os arquivos `.xls` ou `.xlsx` dentro da pasta `data/`.
2. Rode o projeto:

```bash
pip install -r requirements.txt
streamlit run app.py
```

Também é possível subir novos arquivos direto pela barra lateral do app.

## Indicadores gerados

- Total de chamados
- Chamados abertos
- Chamados finalizados
- SLA cumprido
- Chamados fora do SLA
- TMA / tempo médio de atendimento
- FCR em até 1 hora
- Backlog
- Evolução mensal
- Comparativo mês atual x mês anterior
- Top clientes
- Top motivos/categorias
- Pontos de atenção e recomendações executivas

## Observação

O app tenta identificar automaticamente as colunas do relatório. Caso alguma coluna seja detectada incorretamente, ajuste na barra lateral em **Mapeamento das colunas**.
