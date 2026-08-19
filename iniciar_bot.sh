#!/bin/bash
# Script para iniciar el bot de finanzas cargando las variables de entorno

set -a
source .env
set +a

echo "🔍 Verificando configuración..."
python3 verificar_config.py

if [ $? -eq 0 ]; then
    echo ""
    echo "✅ ¡Configuración correcta! Iniciando bot..."
    echo ""
    python3 bot_finanzas.py
else
    echo ""
    echo "❌ Hay errores en la configuración. Revisa el archivo .env"
    exit 1
fi
