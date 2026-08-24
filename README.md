# Fortneer — Devolutiva SST em Streamlit

Aplicação sem IA para:

1. receber o **Relatório de Empresas Geral** em PDF;
2. separar os registros por **SETOR/CARGO**;
3. extrair as conclusões explícitas de **insalubridade**;
4. extrair o **GFIP/LTCAT** por cargo;
5. extrair **ASO e Exames Obrigatórios** com periodicidade;
6. receber as **ações do PGR** em uma tabela editável;
7. gerar automaticamente a apresentação `Modelo_Devolutiva_SST_Fortneer.pptx`.

## Estrutura da pasta

```text
fortneer_streamlit_sst/
├── app.py
├── requirements.txt
└── Modelo_Devolutiva_SST_Fortneer.pptx
```

## Instalação

No terminal, dentro da pasta:

```bash
python -m venv .venv
```

Windows:

```bash
.venv\Scripts\activate
```

Instale as dependências:

```bash
pip install -r requirements.txt
```

Execute:

```bash
streamlit run app.py
```

## Observações importantes

- Não há uso de LLM/IA.
- A extração depende dos marcadores textuais do relatório-modelo, especialmente `SETOR:`, `CARGO:`, `AGENTE:`, `Conclusão do Laudo:`, `CONCLUSÃO LTCAT:`, `Atestado (ASO):`, `Exame:` e `Periodicidade:`.
- O app exibe os dados extraídos em tabelas editáveis antes da geração do PPTX.
- Quando um cargo possui vários agentes, o app mantém todos os códigos GFIP encontrados no resumo do cargo (por exemplo `GFIP 1, GFIP 4`) e disponibiliza as conclusões completas para conferência.
- A apresentação duplica automaticamente os slides de funções, exames e plano de ação quando o volume de registros excede a capacidade do slide-base.
