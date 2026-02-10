#!/bin/sh

echo "==================================="
echo "Agente Financeiro - Container Init"
echo "==================================="
echo "Timezone: $(date)"
echo ""

# --- Log cleanup ---
# Remove log files older than 30 days
find /app/logs -name "*.log" -mtime +30 -delete 2>/dev/null
LOG_COUNT=$(find /app/logs -name "*.log" 2>/dev/null | wc -l)
echo "Log files: $LOG_COUNT (files older than 30 days are auto-deleted)"
echo ""

# --- Start cron ---
crond -b -l 2

echo "Cron iniciado. Próxima execução: meia-noite"
echo "Para executar manualmente: docker compose exec agente-financeiro npm start"
echo ""

# Keep container running
tail -f /app/logs/cron.log 2>/dev/null || tail -f /dev/null
