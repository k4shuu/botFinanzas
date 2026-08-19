#!/usr/bin/env python3
"""
Script de verificación de configuración del bot
"""

import os
import sys

def verificar_configuracion():
    print("\n" + "="*60)
    print("🔍 VERIFICADOR DE CONFIGURACIÓN DEL BOT")
    print("="*60 + "\n")
    
    checks = {
        "TELEGRAM_TOKEN": "Token de Telegram",
        "TELEGRAM_CHAT_ID": "Chat ID de Telegram",
        "GMAIL_USER": "Usuario de Gmail",
        "GMAIL_PASS": "Contraseña de Gmail",
        "GEMINI_API_KEY": "API Key de Gemini",
    }
    
    resultados = {}
    
    for var, descripcion in checks.items():
        valor = os.getenv(var)
        if valor:
            # Mostrar parte del valor para verificación
            if "TOKEN" in var or "KEY" in var or "PASS" in var:
                valor_mostrado = valor[:10] + "..." if len(valor) > 10 else "***"
            else:
                valor_mostrado = valor
            
            print(f"✅ {descripcion:<30} {valor_mostrado}")
            resultados[var] = True
        else:
            print(f"❌ {descripcion:<30} NO CONFIGURADO")
            resultados[var] = False
    
    print("\n" + "-"*60)
    
    # Verificar archivos
    print("\n📁 Verificación de archivos:")
    
    archivos_necesarios = {
        "registro_finanzas.xlsx": "Archivo de transacciones",
        "bot_finanzas.py": "Script principal del bot",
        "requirements.txt": "Dependencias Python",
    }
    
    for archivo, descripcion in archivos_necesarios.items():
        ruta = f"/home/kashu/Documentos/BotFinanzas/{archivo}"
        if os.path.exists(ruta):
            tamaño = os.path.getsize(ruta)
            print(f"✅ {descripcion:<30} ({tamaño} bytes)")
        else:
            print(f"❌ {descripcion:<30} FALTA")
    
    print("\n" + "-"*60)
    print("\n📦 Verificación de dependencias Python:")
    
    dependencias = [
        "telegram",
        "openpyxl",
        "google",
    ]
    
    try:
        import importlib
        for dep in dependencias:
            try:
                importlib.import_module(dep)
                print(f"✅ {dep:<30} instalado")
            except ImportError:
                print(f"❌ {dep:<30} NO instalado")
    except Exception as e:
        print(f"⚠️  Error al verificar dependencias: {e}")
    
    print("\n" + "="*60)
    
    # Resumen
    variables_ok = sum(1 for v in resultados.values() if v)
    variables_total = len(resultados)
    
    if variables_ok == variables_total:
        print("✅ TODO ESTÁ CONFIGURADO CORRECTAMENTE")
        print("Puedes ejecutar: python bot_finanzas.py")
    else:
        print(f"⚠️  FALTAN CONFIGURAR: {variables_total - variables_ok}/{variables_total} variables")
        print("\nPara configurar, copia .env.example a .env y completa tus datos:")
        print("  cp .env.example .env")
        print("  nano .env  (o edita con tu editor favorito)")
        print("\nLuego ejecuta:")
        print("  export $(cat .env | xargs)")
        print("  python bot_finanzas.py")
    
    print("="*60 + "\n")

if __name__ == "__main__":
    verificar_configuracion()
