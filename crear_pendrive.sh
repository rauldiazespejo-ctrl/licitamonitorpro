#!/bin/bash
# ============================================================
#  Crea la carpeta pendrive lista para copiar al USB
#  Ejecutar en macOS/Linux: bash crear_pendrive.sh
# ============================================================

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
DEST="$SCRIPT_DIR/Pendrive_LicitaMonitor"

echo ""
echo "  ╔══════════════════════════════════════════════════════╗"
echo "  ║     LICITAMONITOR — Creando paquete pendrive        ║"
echo "  ╚══════════════════════════════════════════════════════╝"
echo ""

# Limpiar destino anterior
rm -rf "$DEST"
mkdir -p "$DEST"

# Archivos fuente
cp "$SCRIPT_DIR/LicitaMonitor.py"     "$DEST/"
cp "$SCRIPT_DIR/LicitaMonitor_GUI.py" "$DEST/"
cp "$SCRIPT_DIR/config.json"          "$DEST/"
cp "$SCRIPT_DIR/config.example.json"  "$DEST/"
cp "$SCRIPT_DIR/INSTALAR.bat"         "$DEST/"
cp "$SCRIPT_DIR/EJECUTAR.bat"         "$DEST/"
cp "$SCRIPT_DIR/LEEME.txt"            "$DEST/"
cp "$SCRIPT_DIR/build.bat"            "$DEST/"
cp "$SCRIPT_DIR/LicitaMonitor.spec"   "$DEST/"
[ -f "$SCRIPT_DIR/logo.png" ] && cp "$SCRIPT_DIR/logo.png" "$DEST/"
[ -f "$SCRIPT_DIR/icon.ico" ] && cp "$SCRIPT_DIR/icon.ico" "$DEST/"

echo "  Archivos copiados:"
ls -1 "$DEST/"

echo ""
echo "  ╔══════════════════════════════════════════════════════╗"
echo "  ║                  LISTO                              ║"
echo "  ╠══════════════════════════════════════════════════════╣"
echo "  ║  Carpeta: Pendrive_LicitaMonitor/                   ║"
echo "  ║                                                     ║"
echo "  ║  En el equipo Windows destino:                      ║"
echo "  ║    · Copiar carpeta al escritorio o disco           ║"
echo "  ║    · Doble clic en INSTALAR.bat                     ║"
echo "  ║    · Sigue las instrucciones en pantalla            ║"
echo "  ║                                                     ║"
echo "  ║  Para compilar .exe (opcional, en Windows):         ║"
echo "  ║    · Doble clic en build.bat                        ║"
echo "  ╚══════════════════════════════════════════════════════╝"
echo ""
