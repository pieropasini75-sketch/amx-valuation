#!/bin/sh
# Recalculate the workbook with LibreOffice (headless) so cached values exist for viewers that do not recalc.
set -e
IN="$1"; OUTDIR="$2"
soffice --headless --calc --convert-to xlsx --outdir "$OUTDIR" "$IN" >/dev/null 2>&1
