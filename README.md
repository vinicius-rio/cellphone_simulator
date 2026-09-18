# Chatbot Conversation Viewer

Visualizador local de conversas de chatbot a partir de uma base `.xlsx`. O projeto reconstrói cada sessão em uma interface inspirada em um smartphone, com mensagens do cliente, respostas do bot, horários, botões/links e metadados técnicos por interação.

## Estrutura

```text
chatbot_conversation_viewer/
├── data/
│   └── conversas_ficticias.xlsx
├── dist/
│   ├── index.html
│   └── conversas.json
├── src/
│   └── build_viewer.py
├── .gitignore
├── requirements.txt
├── gerar_viewer.bat
├── gerar_viewer.sh
└── README.md
```

## Colunas do XLSX

| Coluna | Nível | Obrigatória | Uso |
|---|---|---:|---|
| `id_sessao` | sessão | sim | Identificador da conversa |
| `seq_sessao` | interação | sim | Ordem das interações |
| `mensagem_cliente` | interação | sim | Mensagem do cliente |
| `resposta` | interação | sim | Resposta do chatbot |
| `datahora` | interação | sim | Data e hora da interação |
| `tool` | interação | não | API/tool generativa usada naquela interação |
| `nome_intencao` | interação | não | Tipo/intenção da resposta gerada pela tool |
| `demanda` | sessão | não | Demanda principal |
| `termino_conversa` | sessão | não | Ex.: resolutivo, transbordo, abandono |
| `nota_csat` | sessão | não | Nota CSAT |
| `comentario_csat` | sessão | não | Comentário CSAT |
| `tipo_csat` | sessão | não | Origem do CSAT: `SMS` ou `CHAT IATI` |

### Regra dos campos de sessão

`demanda`, `termino_conversa`, `nota_csat`, `comentario_csat` e `tipo_csat` devem repetir **o mesmo valor em todas as linhas do mesmo `id_sessao`**. O script valida essa regra e interrompe a geração caso encontre valores contraditórios dentro da mesma sessão.

`tool` e `nome_intencao`, por outro lado, são campos **da interação** e podem mudar a cada `seq_sessao`.

## Formato de botões e links

O parser reconhece:

```text
Sobre o que você quer falar de pix?
Botão 1. aumentar limite do pix
Botão 2. cadastrar chave pix
Link 1. ir pra página de pix
```

Botões e links aparecem no mesmo bloco visual. Botões usam `›` e links usam `↗`.

## Interface

- sidebar com scroll próprio para centenas de sessões;
- busca por `id_sessao` ou demanda;
- filtros por demanda, término, CSAT e `tipo_csat`;
- demanda limitada a 2 linhas na sidebar, com texto completo no hover;
- header do celular exibindo apenas o `id_sessao` abaixo de `Inteligência de Atendimento Itaú`;
- scroll independente da conversa;
- links renderizados junto aos botões, identificados por `↗`;
- hover em cada interação mostrando:
  - `seq_sessao`;
  - `tool`;
  - `nome_intencao`.

## Como executar

```bash
pip install -r requirements.txt
python src/build_viewer.py
```

No Windows, você também pode executar `gerar_viewer.bat`.

Para usar outro arquivo:

```bash
python src/build_viewer.py --input caminho/minha_base.xlsx --output dist/index.html
```

O HTML gerado é autocontido: os dados ficam embutidos no próprio arquivo e ele abre diretamente no navegador.
