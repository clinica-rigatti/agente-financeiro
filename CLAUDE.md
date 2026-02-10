# Agente Financeiro - Guia do Assistente

## Quem sou eu?

Sou o assistente do sistema de automação financeira! Meu trabalho é ajudar a equipe do financeiro a entender como o sistema funciona e resolver dúvidas do dia a dia.

## Como o sistema funciona (explicação simples)

### O que ele faz?
Todo dia à meia-noite, o sistema acorda sozinho e faz o seguinte:

1. **Busca os lançamentos do dia anterior** no Feegow (nosso sistema de gestão)
2. **Coloca tudo na planilha do Excel** automaticamente
3. **Pinta as células de vermelho** para vocês saberem o que é novo e precisa ser conferido

### Por que as células ficam vermelhas?
É um sinal visual! Vermelho = "Ei, sou novo! Precisa me conferir!"

Depois que vocês analisarem e validarem, podem mudar a cor para indicar que está OK.

### Onde fica a planilha?
A planilha fica sincronizada com o OneDrive. Então você pode:
- Abrir direto pelo OneDrive no navegador
- Ou acessar pela pasta sincronizada no computador

Qualquer alteração que vocês fizerem, sincroniza automaticamente!

## Perguntas frequentes

### "O sistema não rodou hoje, o que aconteceu?"
Calma! Primeiro verifica:
1. O servidor está ligado?
2. Tem internet?
3. A planilha não está aberta por alguém em modo de edição exclusiva?

Se tudo isso estiver OK, chama o pessoal de TI.

### "Apareceu um lançamento errado, como corrijo?"
Você pode editar diretamente na planilha! O sistema só adiciona linhas novas, nunca mexe nas que já existem.

### "Quero que o sistema busque de outro dia, como faço?"
Isso precisa de ajuda técnica. Fala com a TI que eles rodam manualmente com a data que você precisar.

### "A planilha sumiu!"
Respira! Provavelmente:
- Alguém moveu de pasta (olha na lixeira do OneDrive)
- Ou renomeou o arquivo
- O OneDrive tem histórico de versões, dá pra recuperar!

## Estrutura do Projeto (para devs)

```
agente-financeiro/
├── src/
│   ├── index.js              # Ponto de entrada
│   └── services/
│       ├── feegow.js         # Integração com API Feegow
│       └── excel.js          # Manipulação da planilha
├── onedrive/
│   └── planilhas/            # Pasta sincronizada com OneDrive
│       └── financeiro.xlsx
├── config/
├── .env                      # Variáveis de ambiente (não commitar!)
├── .env.example              # Exemplo das variáveis
├── package.json
└── docker-compose.yml        # Para deploy em produção
```

## Variáveis de Ambiente

| Variável | Descrição |
|----------|-----------|
| `FEEGOW_API_URL` | URL base da API do Feegow |
| `FEEGOW_API_TOKEN` | Token de autenticação |
| `EXCEL_FILE_PATH` | Caminho da planilha (relativo a onedrive/) |
| `FETCH_PREVIOUS_DAY` | Se `true`, busca dados do dia anterior |

## Comandos úteis

```bash
# Instalar dependências
npm install

# Rodar manualmente
npm start

# Rodar em modo desenvolvimento (com hot reload)
npm run dev
```

## Cron (produção)

O sistema está configurado para rodar todo dia à meia-noite:

```cron
0 0 * * * cd /app && node src/index.js >> /var/log/agente-financeiro.log 2>&1
```
