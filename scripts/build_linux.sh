#!/bin/bash
# Build script for Linux
set -e

VERSION=${1:-$(cat VERSION 2>/dev/null || echo "2.0.0")}
APP_NAME="EML-to-PDF"

echo "Building $APP_NAME v$VERSION for Linux..."

# Clean previous builds
rm -rf build dist

# Install dependencies if needed
pip install -r requirements-dev.txt

# Run build
python build.py --platform linux --package

# Make executable
chmod +x "dist/$APP_NAME-v$VERSION-Linux" 2>/dev/null || true

echo ""
echo "Build complete!"
echo "Output: dist/$APP_NAME-v$VERSION-Linux"
echo ""
echo "To run: ./dist/$APP_NAME-v$VERSION-Linux"
