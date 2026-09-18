# Chatbot Conversation Viewer

Visualizador local de conversas de chatbot a partir de uma base `.xlsx`. O projeto reconstrói cada sessão em uma interface inspirada em um smartphone, com mensagens do cliente, respostas do bot, horários, botões, links e metadados da sessão.

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

## Colunas esperadas no XLSX

| Coluna | Obrigatória | Uso |
|---|---|---|
| `id_sessao` | sim | Identificador da conversa |
| `seq_sessao` | sim | Ordem das interações |
| `mensagem_cliente` | sim | Mensagem do cliente |
| `resposta` | sim | Resposta do chatbot |
| `datahora` | sim | Data e hora da interação |
| `demanda` | não | Demanda principal |
| `termino_conversa` | não | Ex.: resolutivo, transbordo, abandono |
| `nota_csat` | não | Nota CSAT |
| `comentario_csat` | não | Comentário CSAT |
| `tipo_csat` | não | Origem do CSAT: `SMS` ou `CHAT IATI` |

## Formato de botões e links

O parser reconhece linhas no campo `resposta` neste padrão:

```text
Sobre o que você quer falar de pix?
Botão 1. aumentar limite do pix
Botão 2. cadastrar chave pix
Link 1. ir pra página de pix
```

No HTML, tanto `Botão N.` quanto `Link N.` aparecem dentro do mesmo bloco visual de opções. Botões usam a seta `›` e links usam `↗`, preservando a identificação sem poluir a interface. O prefixo e a numeração não aparecem para o analista.

## Como executar

Crie e ative um ambiente virtual, depois instale as dependências:

```bash
pip install -r requirements.txt
```

Gere o visualizador:

```bash
python src/build_viewer.py
```

No Windows, você também pode executar `gerar_viewer.bat`: ele cria o ambiente virtual, instala as dependências, gera o HTML e abre o resultado no navegador.

Abra depois:

```text
dist/index.html
```

Para usar outro arquivo:

```bash
python src/build_viewer.py --input caminho/minha_base.xlsx --output dist/index.html
```

## Base fictícia

`data/conversas_ficticias.xlsx` contém 10 sessões fictícias, com 3 a 8 interações por sessão, cobrindo demandas como Pix, cartão, boleto, limite, Seguro Pet, empréstimo, senha, investimentos, conta e entrega de cartão.

## Interface

- seletor de sessões no desktop;
- busca por `id_sessao` ou demanda;
- mockup de smartphone;
- scroll interno da conversa;
- mensagens do cliente à direita;
- respostas do chatbot à esquerda;
- reconstrução automática de botões e links no mesmo bloco visual;
- horário de cada interação;
- `seq_sessao` visível no hover;
- demanda, término da conversa, número de interações, duração, CSAT e `tipo_csat`;
- layout responsivo para abrir também em celular.

## Observação

O HTML gerado é autocontido: os dados ficam embutidos no próprio arquivo, então ele pode ser aberto diretamente no navegador sem servidor web.

## Filtros e navegação

A lateral de sessões possui altura fixa e scroll próprio, então a página não cresce mesmo com centenas de sessões. O viewer também oferece:

- busca por `id_sessao` ou demanda;
- filtro por demanda;
- filtro por término da conversa;
- filtro por nota CSAT, incluindo sessões sem nota;
- filtro por `tipo_csat` (`SMS` ou `CHAT IATI`);
- ordenação por data mais recente, mais antiga ou ID;
- contador de sessões exibidas após os filtros;
- botão para limpar todos os filtros.
