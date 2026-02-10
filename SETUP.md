# Agente Financeiro - Guia de Setup (Teste Local com Docker)

## Objetivo

Validar que o OneDrive sincroniza corretamente com a planilha e que o agente
funciona em modo automático (cron) dentro do Docker, antes de subir para a VPS.

---

## Pré-requisitos

- Docker e Docker Compose instalados
- Conta Microsoft com acesso ao OneDrive
- Acesso à internet (para API Feegow e autenticação OneDrive)

---

## Passo 1: Preparar o OneDrive

Antes de subir os containers, a planilha precisa estar no OneDrive online.

1. Acesse [onedrive.live.com](https://onedrive.live.com)
2. Crie a pasta `planilhas/` na raiz do seu OneDrive
3. Faça upload da planilha de teste para dentro de `planilhas/`
4. Anote o **nome exato do arquivo** (ex: `COPIA ENTRADAS ANO 2026.xlsx`)

> O container do OneDrive vai sincronizar APENAS a pasta `planilhas/` (configurado
> via variável `ONEDRIVE_SINGLE_DIRECTORY=planilhas` no docker-compose).

---

## Passo 2: Configurar o .env

Copie o exemplo e edite:

```bash
cp .env.example .env
```

Edite o `.env` com os valores reais:

```env
# API Feegow
FEEGOW_API_URL=https://api.feegow.com/v1
FEEGOW_API_TOKEN=<seu_token_jwt>

# Planilha - usar o nome EXATO do arquivo no OneDrive
EXCEL_FILE_PATH=planilhas/COPIA ENTRADAS ANO 2026.xlsx

# Para testes: forçar uma aba específica (descomente a linha abaixo)
# EXCEL_ABA=FEVEREIRO 2026
#
# Para produção: deixe comentado e o sistema resolve pela data automaticamente
# Ex: data 05/02/2026 → aba "FEVEREIRO 2026"

EXCEL_LINHA_INICIAL=8
FETCH_PREVIOUS_DAY=true

# Logs em DEBUG para acompanhar tudo durante os testes
LOG_LEVEL=DEBUG
LOG_TO_FILE=true
```

---

## Passo 3: Limpar a pasta onedrive local

A pasta `./onedrive/` precisa estar vazia antes da primeira sincronização.
O container do OneDrive vai baixar os arquivos do servidor.

```bash
rm -rf ./onedrive/planilhas/*
```

---

## Passo 4: Autenticar o OneDrive (primeira vez)

Esse passo é **interativo** e só precisa ser feito uma vez.

```bash
docker compose run --rm onedrive
```

O container vai exibir algo como:

```
Authorize this app visiting:
https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=...

Enter the response uri:
```

1. Copie a URL completa e abra no navegador
2. Faça login com sua conta Microsoft
3. Autorize o aplicativo
4. O navegador vai redirecionar para uma URL em branco — **copie a URL inteira da barra de endereço**
5. Cole essa URL no terminal onde o container está esperando
6. O container vai fazer o primeiro sync e sair

Verifique se a planilha foi baixada:

```bash
ls -la ./onedrive/planilhas/
```

Deve aparecer o arquivo da planilha.

---

## Passo 5: Subir os containers

```bash
docker compose up -d
```

Verifique se ambos estão rodando:

```bash
docker compose ps
```

Esperado:

```
NAME                STATUS              HEALTH
onedrive-sync       Up (healthy)        healthy
agente-financeiro   Up                  -
```

> Se o `onedrive-sync` ficar `(health: starting)` por mais de 60 segundos,
> verifique os logs: `docker compose logs onedrive`

---

## Passo 6: Teste manual

Execute o agente manualmente para validar o fluxo completo:

```bash
docker compose exec agente-financeiro npm start
```

Acompanhe o que aconteceu:

```bash
# Ver logs do agente
docker compose logs agente-financeiro

# Ver log detalhado do dia
cat ./logs/agente-financeiro-$(date +%Y-%m-%d).log
```

Depois do agente rodar, verifique:

1. **A planilha local foi modificada?**
   ```bash
   ls -la ./onedrive/planilhas/
   ```
   O timestamp do arquivo deve ter atualizado.

2. **O OneDrive sincronizou de volta?**
   ```bash
   docker compose logs onedrive --tail=20
   ```
   Procure por linhas indicando upload do arquivo.

3. **Abra o OneDrive no navegador** e verifique se a planilha online
   tem os novos dados inseridos pelo agente.

---

## Passo 7: Testar o cron

O cron está configurado para rodar à meia-noite (horário de São Paulo).

Para confirmar que o cron está ativo:

```bash
docker compose exec agente-financeiro crontab -l
```

Deve mostrar:

```
0 0 * * * cd /app && node src/index.js >> /app/logs/cron.log 2>&1
```

No dia seguinte, verifique:

```bash
# Log do cron
cat ./logs/cron.log

# Log detalhado do dia
cat ./logs/agente-financeiro-$(date +%Y-%m-%d).log
```

---

## Comandos úteis

| Comando | O que faz |
|---------|-----------|
| `docker compose up -d` | Sobe tudo em background |
| `docker compose down` | Para tudo |
| `docker compose ps` | Status dos containers |
| `docker compose logs -f` | Acompanhar logs em tempo real |
| `docker compose logs onedrive --tail=50` | Últimas 50 linhas do OneDrive |
| `docker compose exec agente-financeiro npm start` | Executar agente manualmente |
| `docker compose restart agente-financeiro` | Reiniciar após mudança no .env |
| `docker compose up -d --build agente-financeiro` | Rebuild após mudança no código |

---

## Troubleshooting

### OneDrive não autentica

Reautentique:

```bash
docker compose down
docker volume rm agente-financeiro_onedrive-config
docker compose run --rm onedrive
```

### OneDrive não sincroniza a planilha

Verifique se o nome da pasta no OneDrive online é exatamente `planilhas` (minúsculo).

**Importante:** A planilha deve estar **fechada** no OneDrive online. Se alguém estiver
com o arquivo aberto, o OneDrive não consegue fazer upload (file lock).

Veja os logs com mais detalhe:

```bash
docker compose exec onedrive onedrive --display-config
docker compose logs onedrive
```

### OneDrive entra em loop de resync (exit code 126)

**NUNCA** monte arquivos de configuração em `/onedrive/conf/` via bind-mount.
Isso causa o erro `Application configuration change detected, --resync required` em loop.

A solução é usar apenas variáveis de ambiente no docker-compose:
- `ONEDRIVE_SINGLE_DIRECTORY=planilhas` (substitui o arquivo sync_list)
- `ONEDRIVE_RESYNC=1` (evita o loop de detecção de mudança de config)

### Agente roda mas não encontra a planilha

O `EXCEL_FILE_PATH` no `.env` precisa bater **exatamente** com o nome do arquivo
sincronizado em `./onedrive/planilhas/`. Verifique com:

```bash
ls ./onedrive/planilhas/
```

### Agente não encontra a aba

Se `EXCEL_ABA` está comentado, o sistema resolve pela data. Para 06/02/2026 com
`FETCH_PREVIOUS_DAY=true`, vai procurar a aba `FEVEREIRO 2026`. Verifique se
essa aba existe na planilha.

### Container agente-financeiro não sobe

Se o healthcheck do OneDrive não passa, o agente fica esperando. Verifique:

```bash
docker compose logs onedrive
docker compose ps
```

---

## Quando tudo estiver validado

Para migrar para produção de verdade:

1. Trocar `EXCEL_FILE_PATH` no `.env` para o nome da planilha real
2. Remover/comentar `EXCEL_ABA` (resolução automática por mês)
3. Trocar `LOG_LEVEL` para `INFO`
4. Subir o projeto na VPS e repetir a partir do Passo 4 (auth do OneDrive)
