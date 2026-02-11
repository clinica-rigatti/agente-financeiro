FROM node:20-alpine

WORKDIR /app

# Instala cron e timezone
RUN apk add --no-cache dcron tzdata
ENV TZ=America/Sao_Paulo

# Copia package.json e instala dependências
COPY package*.json ./
RUN npm ci --only=production

# Copia código fonte
COPY src ./src
COPY CLAUDE.md ./

# Cria diretório de logs
RUN mkdir -p /app/logs

# Configura cron para rodar à meia-noite (horário de São Paulo)
RUN echo "0 0 * * * cd /app && node src/index.js >> /app/logs/cron.log 2>&1" > /etc/crontabs/root

# Script de entrada
COPY docker-entrypoint.sh /docker-entrypoint.sh
RUN chmod +x /docker-entrypoint.sh

ENTRYPOINT ["/docker-entrypoint.sh"]
