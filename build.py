#!/usr/bin/env python3
"""
Cross-platform build script for EML-to-PDF Converter.
Creates standalone executables for macOS, Windows, and Linux.

Usage:
    python build.py                    # Build for current platform
    python build.py --all              # Instructions for all platforms
    python build.py --platform macos   # Build for macOS (must run on Mac)
    python build.py --clean            # Clean build artifacts
    python build.py --release          # Create and push a release tag

IMPORTANT: PyInstaller cannot cross-compile. To build for all platforms:
  - Use GitHub Actions (automatic on version tag push)
  - Or run this script on each target platform separately
"""

import argparse
import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path

# Configuration
APP_NAME = "EML-to-PDF"
MAIN_SCRIPT = "eml_to_pdf.py"
BUNDLE_ID = "com.nassereledroos.eml-to-pdf"


def get_version():
    """Get version from the package."""
    version_file = Path(__file__).parent / "VERSION"
    if version_file.exists():
        return version_file.read_text().strip()

    # Fallback to reading from __init__.py
    try:
        from eml_to_pdf import __version__
        return __version__
    except ImportError:
        return "2.0.0"


VERSION = get_version()

# Platform detection
SYSTEM = platform.system()  # 'Darwin', 'Windows', 'Linux'
PLATFORM_NAMES = {
    'Darwin': 'macOS',
    'Windows': 'Windows',
    'Linux': 'Linux'
}


def clean():
    """Clean build artifacts."""
    dirs_to_remove = ['build', 'dist', '__pycache__', 'eml_to_pdf.egg-info']

    for dir_name in dirs_to_remove:
        path = Path(dir_name)
        if path.exists():
            print(f"Removing {dir_name}/")
            shutil.rmtree(path)

    # Clean __pycache__ in subdirectories
    for pycache in Path('.').rglob('__pycache__'):
        print(f"Removing {pycache}/")
        shutil.rmtree(pycache)

    # Clean .pyc files
    for pyc in Path('.').rglob('*.pyc'):
        pyc.unlink()

    print("Clean complete.")


def check_dependencies():
    """Check if required build dependencies are installed."""
    try:
        import PyInstaller
        return True
    except ImportError:
        print("Error: PyInstaller is not installed.")
        print("Install with: pip install -r requirements-dev.txt")
        return False


def get_pyinstaller_args(one_file=True):
    """Generate PyInstaller arguments based on platform."""
    base_args = [
        sys.executable, "-m", "PyInstaller",
        "--name", APP_NAME,
        "--windowed",
        "--clean",
        "--noconfirm",
    ]

    if one_file:
        base_args.append("--onefile")

    # Platform-specific additions
    icon_path = Path("assets")

    if SYSTEM == "Darwin":
        icon_file = icon_path / "icon.icns"
        if icon_file.exists():
            base_args.extend(["--icon", str(icon_file)])
        base_args.extend(["--osx-bundle-identifier", BUNDLE_ID])
    elif SYSTEM == "Windows":
        icon_file = icon_path / "icon.ico"
        if icon_file.exists():
            base_args.extend(["--icon", str(icon_file)])
    elif SYSTEM == "Linux":
        icon_file = icon_path / "icon.png"
        if icon_file.exists():
            base_args.extend(["--icon", str(icon_file)])

    # Hidden imports for dependencies
    hidden_imports = [
        "reportlab.graphics.barcode.common",
        "reportlab.graphics.barcode.code128",
        "reportlab.graphics.barcode.code39",
        "reportlab.graphics.barcode.code93",
        "reportlab.graphics.barcode.usps",
        "email.mime",
        "email.parser",
    ]

    for imp in hidden_imports:
        base_args.extend(["--hidden-import", imp])

    # Try to add PIL if available
    try:
        import PIL
        base_args.extend(["--hidden-import", "PIL"])
    except ImportError:
        pass

    # Collect data files
    base_args.extend(["--collect-data", "reportlab"])

    # Add main script
    base_args.append(MAIN_SCRIPT)

    return base_args


def build():
    """Execute the build process for current platform."""
    platform_name = PLATFORM_NAMES.get(SYSTEM, SYSTEM)
    print(f"Building {APP_NAME} v{VERSION} for {platform_name}...")
    print()

    if not check_dependencies():
        return False

    args = get_pyinstaller_args()
    print(f"Running: {' '.join(args[:5])}...")

    try:
        result = subprocess.run(args, check=True)
        print()
        print(f"Build successful!")
        print(f"Output: dist/{APP_NAME}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"\nBuild failed with error code {e.returncode}")
        return False


def package_release():
    """Package the build into a release artifact."""
    version_suffix = f"v{VERSION}"
    dist_path = Path("dist")
    platform_name = PLATFORM_NAMES.get(SYSTEM, SYSTEM)

    if not dist_path.exists():
        print("Error: dist/ directory not found. Run build first.")
        return False

    print(f"Packaging release for {platform_name}...")

    if SYSTEM == "Darwin":
        # Create ZIP for macOS
        app_path = dist_path / f"{APP_NAME}.app"
        exe_path = dist_path / APP_NAME

        if app_path.exists():
            output_name = f"{APP_NAME}-{version_suffix}-macOS"
            shutil.make_archive(
                str(dist_path / output_name),
                'zip',
                dist_path,
                f"{APP_NAME}.app"
            )
            print(f"Created: dist/{output_name}.zip")
        elif exe_path.exists():
            output_name = f"{APP_NAME}-{version_suffix}-macOS"
            output_path = dist_path / output_name
            shutil.copy(exe_path, output_path)
            os.chmod(output_path, 0o755)
            print(f"Created: dist/{output_name}")

    elif SYSTEM == "Windows":
        exe_path = dist_path / f"{APP_NAME}.exe"
        if exe_path.exists():
            new_name = f"{APP_NAME}-{version_suffix}-Windows.exe"
            new_path = dist_path / new_name
            if new_path.exists():
                new_path.unlink()
            shutil.copy(exe_path, new_path)
            print(f"Created: dist/{new_name}")

    elif SYSTEM == "Linux":
        exe_path = dist_path / APP_NAME
        if exe_path.exists():
            new_name = f"{APP_NAME}-{version_suffix}-Linux"
            new_path = dist_path / new_name
            if new_path.exists():
                new_path.unlink()
            shutil.copy(exe_path, new_path)
            os.chmod(new_path, 0o755)
            print(f"Created: dist/{new_name}")

    return True


def show_all_platforms_info():
    """Show information about building for all platforms."""
    print("""
================================================================================
                    Building for All Platforms
================================================================================

PyInstaller CANNOT cross-compile. You must build on each target platform.

OPTION 1: Use GitHub Actions (Recommended)
------------------------------------------
Push a version tag to trigger automatic builds for all platforms:

    git add .
    git commit -m "Release v{version}"
    git tag v{version}
    git push origin main --tags

This will:
  - Build for macOS, Windows, and Linux automatically
  - Create a GitHub Release with all artifacts
  - No need to own each platform!

OPTION 2: Manual Builds
-----------------------
Run this script on each platform:

  On macOS:   python build.py --package
  On Windows: python build.py --package
  On Linux:   python build.py --package

OPTION 3: Use Virtual Machines / Cloud
--------------------------------------
  - macOS: Use a Mac or macOS VM (macOS cannot be legally virtualized on non-Apple hardware)
  - Windows: Use a Windows VM or GitHub Actions windows-latest
  - Linux: Use a Linux VM, Docker, or GitHub Actions ubuntu-latest

Current Platform: {platform} ({system})
================================================================================
""".format(version=VERSION, platform=PLATFORM_NAMES.get(SYSTEM, SYSTEM), system=SYSTEM))


def create_release_tag():
    """Create and push a release tag to trigger GitHub Actions."""
    print(f"Creating release tag v{VERSION}...")

    # Check if we're in a git repo
    if not Path(".git").exists():
        print("Error: Not in a git repository")
        return False

    # Check for uncommitted changes
    result = subprocess.run(
        ["git", "status", "--porcelain"],
        capture_output=True, text=True
    )
    if result.stdout.strip():
        print("Warning: You have uncommitted changes.")
        response = input("Continue anyway? (y/N): ").strip().lower()
        if response != 'y':
            print("Aborted.")
            return False

    # Check if tag already exists
    result = subprocess.run(
        ["git", "tag", "-l", f"v{VERSION}"],
        capture_output=True, text=True
    )
    if result.stdout.strip():
        print(f"Error: Tag v{VERSION} already exists.")
        print("Update VERSION file to create a new release.")
        return False

    # Create tag
    try:
        subprocess.run(
            ["git", "tag", f"v{VERSION}"],
            check=True
        )
        print(f"Created tag: v{VERSION}")

        # Push tag
        response = input("Push tag to origin to trigger builds? (y/N): ").strip().lower()
        if response == 'y':
            subprocess.run(
                ["git", "push", "origin", f"v{VERSION}"],
                check=True
            )
            print(f"Pushed tag v{VERSION}")
            print()
            print("GitHub Actions will now build releases for all platforms.")
            print("Check progress at: https://github.com/YOUR_USERNAME/EML-to-PDF/actions")
        else:
            print(f"Tag created locally. Push manually with: git push origin v{VERSION}")

        return True

    except subprocess.CalledProcessError as e:
        print(f"Git command failed: {e}")
        return False


def create_spec_file():
    """Create a PyInstaller spec file for more control."""
    spec_content = f'''# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for {APP_NAME}.
Generated by build.py - Version {VERSION}
"""

import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# Collect reportlab data files (fonts, etc.)
reportlab_datas = collect_data_files('reportlab')

a = Analysis(
    ['{MAIN_SCRIPT}'],
    pathex=[],
    binaries=[],
    datas=reportlab_datas,
    hiddenimports=[
        'reportlab.graphics.barcode.common',
        'reportlab.graphics.barcode.code128',
        'reportlab.graphics.barcode.code39',
        'reportlab.graphics.barcode.code93',
        'reportlab.graphics.barcode.usps',
        'email.mime',
        'email.parser',
    ],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='{APP_NAME}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/icon.icns' if sys.platform == 'darwin' else ('assets/icon.ico' if sys.platform == 'win32' else 'assets/icon.png'),
)

# macOS app bundle
if sys.platform == 'darwin':
    app = BUNDLE(
        exe,
        name='{APP_NAME}.app',
        icon='assets/icon.icns',
        bundle_identifier='{BUNDLE_ID}',
        info_plist={{
            'CFBundleShortVersionString': '{VERSION}',
            'CFBundleVersion': '{VERSION}',
            'NSHighResolutionCapable': True,
            'LSMinimumSystemVersion': '10.13.0',
        }},
    )
'''

    spec_file = Path("eml_to_pdf.spec")
    spec_file.write_text(spec_content)
    print(f"Created {spec_file}")


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Build EML-to-PDF Converter for distribution",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python build.py                Build for current platform
  python build.py --package      Build and package for release
  python build.py --all          Show how to build for all platforms
  python build.py --release      Create a release tag (triggers GitHub Actions)
  python build.py --clean        Clean build artifacts

Note: PyInstaller cannot cross-compile. Use --all for multi-platform options.
        """
    )

    parser.add_argument(
        "--all",
        action="store_true",
        help="Show how to build for all platforms"
    )

    parser.add_argument(
        "--clean",
        action="store_true",
        help="Clean build artifacts"
    )

    parser.add_argument(
        "--package",
        action="store_true",
        help="Build and package as a release artifact"
    )

    parser.add_argument(
        "--release",
        action="store_true",
        help="Create a release tag to trigger GitHub Actions builds"
    )

    parser.add_argument(
        "--spec",
        action="store_true",
        help="Generate PyInstaller spec file only"
    )

    parser.add_argument(
        "--version",
        action="store_true",
        help="Show version and exit"
    )

    args = parser.parse_args()

    # Show version
    if args.version:
        print(f"{APP_NAME} v{VERSION}")
        return 0

    # Show all platforms info
    if args.all:
        show_all_platforms_info()
        return 0

    # Clean if requested
    if args.clean:
        clean()
        return 0

    # Generate spec file if requested
    if args.spec:
        create_spec_file()
        return 0

    # Create release tag
    if args.release:
        if create_release_tag():
            return 0
        return 1

    # Default: Build for current platform
    print(f"Current platform: {PLATFORM_NAMES.get(SYSTEM, SYSTEM)}")
    print()

    if not build():
        return 1

    # Package if requested
    if args.package:
        print()
        if not package_release():
            return 1

    print()
    print("=" * 60)
    print("Build complete!")
    print()
    print(f"To build for OTHER platforms, use: python build.py --all")
    print("=" * 60)

    return 0


if __name__ == "__main__":
    sys.exit(main())
