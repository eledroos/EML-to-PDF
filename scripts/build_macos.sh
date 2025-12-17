#!/bin/bash
# Build script for macOS
set -e

VERSION=${1:-$(cat VERSION 2>/dev/null || echo "2.0.0")}
APP_NAME="EML-to-PDF"

echo "Building $APP_NAME v$VERSION for macOS..."

# Clean previous builds
rm -rf build dist

# Install dependencies if needed
pip install -r requirements-dev.txt

# Run build
python build.py --platform macos --package

# Ad-hoc code sign (for running without Gatekeeper issues locally)
if [ -d "dist/$APP_NAME.app" ]; then
    echo "Ad-hoc signing the app..."
    codesign --force --deep --sign - "dist/$APP_NAME.app" || echo "Code signing skipped (may require Xcode)"
fi

echo ""
echo "Build complete!"
echo "Output: dist/$APP_NAME-v$VERSION-macOS.zip"
echo ""
echo "Note: Users may need to right-click and select 'Open' on first launch"
echo "due to macOS Gatekeeper restrictions on unsigned applications."
