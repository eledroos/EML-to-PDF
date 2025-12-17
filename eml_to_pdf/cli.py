"""Command-line interface for EML to PDF Converter."""

import argparse
import sys
import os
import time
from typing import Optional

from .config import ConversionConfig
from .converter import convert_batch, create_skipped_files_report


def create_parser() -> argparse.ArgumentParser:
    """Create the argument parser for the CLI."""
    parser = argparse.ArgumentParser(
        prog='eml-to-pdf',
        description='Convert EML email files to PDF format',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s -i ~/emails                    Convert emails in ~/emails folder
  %(prog)s -i ~/emails -o ~/pdfs          Specify output folder
  %(prog)s -i ~/emails --extract-attachments  Save email attachments
  %(prog)s --gui                          Launch graphical interface
        """
    )

    parser.add_argument(
        '-i', '--input',
        type=str,
        metavar='FOLDER',
        help='Input folder containing EML files'
    )

    parser.add_argument(
        '-o', '--output',
        type=str,
        metavar='FOLDER',
        help='Output folder for PDFs (default: INPUT/PDF)'
    )

    parser.add_argument(
        '--page-size',
        choices=['letter', 'a4'],
        default='letter',
        help='PDF page size (default: letter)'
    )

    parser.add_argument(
        '--extract-attachments',
        action='store_true',
        help='Extract and save email attachments'
    )

    parser.add_argument(
        '--no-organize',
        action='store_true',
        help='Do not organize output by year/month folders'
    )

    parser.add_argument(
        '--no-weasyprint',
        action='store_true',
        help='Disable WeasyPrint HTML rendering (use simple text extraction)'
    )

    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Enable verbose output'
    )

    parser.add_argument(
        '-q', '--quiet',
        action='store_true',
        help='Suppress all output except errors'
    )

    parser.add_argument(
        '--gui',
        action='store_true',
        help='Launch graphical user interface'
    )

    parser.add_argument(
        '--version',
        action='version',
        version='%(prog)s 2.0.0'
    )

    return parser


def print_progress(current: int, total: int, filename: str, start_time: float, width: int = 40) -> None:
    """Print a terminal progress bar."""
    if total == 0:
        return

    percent = current / total
    filled = int(width * percent)
    bar = '=' * filled + '-' * (width - filled)

    # Calculate ETA
    elapsed = time.time() - start_time
    if current > 0:
        eta = (elapsed / current) * (total - current)
        eta_str = f"ETA: {int(eta)}s"
    else:
        eta_str = "ETA: --"

    # Truncate filename for display
    display_name = filename[:30] + '...' if len(filename) > 30 else filename.ljust(33)

    print(f'\r[{bar}] {current}/{total} {display_name} {eta_str}', end='', flush=True)


def run_cli(args: argparse.Namespace) -> int:
    """
    Run the CLI conversion.

    Args:
        args: Parsed command line arguments

    Returns:
        Exit code (0 for success, 1 for error)
    """
    # Validate input folder
    if not args.input:
        print("Error: Input folder is required. Use -i or --input to specify.", file=sys.stderr)
        print("Use --gui to launch the graphical interface instead.", file=sys.stderr)
        return 1

    if not os.path.isdir(args.input):
        print(f"Error: Input folder does not exist: {args.input}", file=sys.stderr)
        return 1

    # Build configuration
    config = ConversionConfig(
        page_size=args.page_size,
        organize_by_date=not args.no_organize,
        extract_attachments=args.extract_attachments,
        use_weasyprint=not args.no_weasyprint,
    )

    # Determine output folder
    output_folder = args.output or os.path.join(args.input, "PDF")

    if not args.quiet:
        print(f"Converting EML files from: {args.input}")
        print(f"Output folder: {output_folder}")
        if args.extract_attachments:
            print("Attachment extraction: enabled")
        print()

    # Track progress
    start_time = time.time()
    current_file = ""

    def progress_callback(current: int, total: int, filename: str) -> bool:
        nonlocal current_file
        current_file = filename

        if not args.quiet:
            print_progress(current, total, filename, start_time)

        if args.verbose and filename != "Complete":
            print(f"\n  Processing: {filename}")

        return True  # Continue processing

    # Run conversion
    try:
        result = convert_batch(
            input_folder=args.input,
            output_folder=output_folder,
            config=config,
            progress_callback=progress_callback
        )
    except KeyboardInterrupt:
        print("\n\nConversion cancelled by user.")
        return 1

    # Clear progress line
    if not args.quiet:
        print()
        print()

    # Handle results
    if result.total_files == 0:
        print("No EML files found in the specified folder.")
        return 1

    # Create skipped files report if needed
    if result.failed > 0:
        report_path = create_skipped_files_report(result.results, result.output_folder)

    # Print summary
    if not args.quiet:
        elapsed = time.time() - start_time
        print("=" * 50)
        print("Conversion Complete!")
        print("=" * 50)
        print(f"Total files:    {result.total_files}")
        print(f"Successful:     {result.successful}")
        print(f"Failed:         {result.failed}")
        print(f"Time elapsed:   {elapsed:.1f}s")
        print(f"Output folder:  {result.output_folder}")

        if result.failed > 0:
            print(f"\nSee Skipped_Files_Report.pdf for details on failed conversions.")

    if args.verbose:
        # Print details of failed files
        failed = [r for r in result.results if not r.success]
        if failed:
            print("\nFailed files:")
            for r in failed:
                print(f"  - {r.source_file}: {r.error_message}")

    return 0 if result.failed == 0 else 1


def main(args: Optional[list] = None) -> int:
    """
    Main entry point for the CLI.

    Args:
        args: Optional list of arguments (uses sys.argv if not provided)

    Returns:
        Exit code
    """
    parser = create_parser()
    parsed_args = parser.parse_args(args)

    # Launch GUI if requested or no arguments provided
    if parsed_args.gui or (not parsed_args.input and len(sys.argv) == 1):
        from .gui import launch_gui
        launch_gui()
        return 0

    return run_cli(parsed_args)


if __name__ == '__main__':
    sys.exit(main())
