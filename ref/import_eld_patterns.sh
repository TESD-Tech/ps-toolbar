#!/bin/bash
# Script to copy ELD project patterns into ref/ for analysis
ELD_DIR="$HOME/Projects/eld-progress-report"
TARGET_DIR="./ref/eld-patterns"

mkdir -p "$TARGET_DIR"

echo "Copying config files..."
cp "$ELD_DIR/package.json" "$TARGET_DIR/" 2>/dev/null
cp "$ELD_DIR/vite.config.ts" "$TARGET_DIR/" 2>/dev/null
cp "$ELD_DIR/vite.config.js" "$TARGET_DIR/" 2>/dev/null
cp "$ELD_DIR/tsconfig.json" "$TARGET_DIR/" 2>/dev/null
cp "$ELD_DIR/GEMINI.md" "$TARGET_DIR/" 2>/dev/null

echo "Copying sample components..."
mkdir -p "$TARGET_DIR/src/components"
cp "$ELD_DIR/src/App.svelte" "$TARGET_DIR/src/" 2>/dev/null
# Copy one or two components from src/components/
ls "$ELD_DIR/src/components/" | head -n 3 | xargs -I {} cp "$ELD_DIR/src/components/{}" "$TARGET_DIR/src/components/" 2>/dev/null

echo "Done. Patterns copied to $TARGET_DIR"
