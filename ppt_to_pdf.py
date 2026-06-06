"""
Batch convert PowerPoint files (.ppt / .pptx) to PDF on Windows.

This script drives the locally installed Microsoft PowerPoint application
through its COM automation interface (via the `comtypes` library). It walks an
input directory, exports every supported presentation to PDF in an output
directory, skips files that have already been converted, and survives individual
file failures without aborting the whole batch.

Requirements:
    * Windows
    * Microsoft PowerPoint installed (the script automates the real app)
    * comtypes  ->  pip install -r requirements.txt
"""

import os
import sys
import glob

import comtypes.client

# PowerPoint's SaveAs / ExportAsFixedFormat file-format constant for PDF.
# 32 == ppSaveAsPDF in the PowerPoint object model.
PP_SAVE_AS_PDF = 32

# Extensions we treat as convertible presentations.
SUPPORTED_EXTENSIONS = (".ppt", ".pptx")


def find_presentations(input_dir):
    """Return a sorted list of presentation file paths inside *input_dir*.

    Only top-level files are considered (no recursion). Temporary lock files
    that PowerPoint creates (names beginning with '~$') are ignored.
    """
    presentations = []
    for ext in SUPPORTED_EXTENSIONS:
        # glob is case-insensitive on Windows, so *.ppt also matches *.PPT.
        for path in glob.glob(os.path.join(input_dir, "*" + ext)):
            filename = os.path.basename(path)
            if filename.startswith("~$"):
                continue  # skip PowerPoint lock/temp files
            presentations.append(path)
    # De-duplicate (a path could match more than one pattern on some systems)
    # and sort for deterministic, predictable processing order.
    return sorted(set(presentations))


def convert_one(powerpoint, source_path, output_dir):
    """Convert a single presentation to PDF.

    Returns one of the strings: "converted", "skipped", or "failed" so the
    caller can keep a running tally.

    A *single* file failing (e.g. a corrupted file) must never crash the batch,
    so all per-file work is wrapped in try/except here.
    """
    filename = os.path.basename(source_path)
    name_without_ext = os.path.splitext(filename)[0]
    target_path = os.path.join(output_dir, name_without_ext + ".pdf")

    # I/O optimization: if the PDF already exists, don't redo the work.
    if os.path.exists(target_path):
        print(f"[SKIP] '{filename}' -> PDF already exists.")
        return "skipped"

    presentation = None
    try:
        print(f"[CONVERT] '{filename}' ...")

        # COM/PowerPoint is happiest with absolute paths.
        abs_source = os.path.abspath(source_path)
        abs_target = os.path.abspath(target_path)

        # Open the presentation.
        #   WithWindow=False keeps PowerPoint from flashing a window for each file.
        #   ReadOnly=True avoids touching the source and dodges some lock issues.
        presentation = powerpoint.Presentations.Open(
            abs_source,
            ReadOnly=True,
            Untitled=False,
            WithWindow=False,
        )

        # Export to PDF (file format 32 == PDF).
        presentation.SaveAs(abs_target, PP_SAVE_AS_PDF)

        print(f"[OK]   '{filename}' -> '{os.path.basename(abs_target)}'")
        return "converted"

    except Exception as exc:  # noqa: BLE001 - we intentionally catch everything per-file
        # Report and move on; one bad file should not stop the batch.
        print(f"[ERROR] Failed to convert '{filename}': {exc}")
        return "failed"

    finally:
        # Always close the presentation if it was opened, so we don't leak
        # open documents inside the PowerPoint process during a long batch.
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass  # closing failures are non-fatal; the app Quit() will clean up


def batch_convert(input_dir, output_dir):
    """Convert every supported presentation in *input_dir* into *output_dir*."""

    if not os.path.isdir(input_dir):
        print(f"[FATAL] Input directory does not exist: {input_dir}")
        return

    # Make sure the output directory exists.
    os.makedirs(output_dir, exist_ok=True)

    presentations = find_presentations(input_dir)
    if not presentations:
        print(f"No .ppt or .pptx files found in: {input_dir}")
        return

    print(f"Found {len(presentations)} presentation(s) in '{input_dir}'.\n")

    powerpoint = None
    tally = {"converted": 0, "skipped": 0, "failed": 0}

    try:
        # Launch / connect to the PowerPoint application via COM.
        # CreateObject starts PowerPoint if it isn't already running.
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")

        # NOTE: We deliberately do NOT force powerpoint.Visible = False.
        # Some PowerPoint versions raise an error when the main window is hidden
        # while automating. Using WithWindow=False on each Open() is the
        # reliable way to keep individual document windows from appearing.

        for source_path in presentations:
            result = convert_one(powerpoint, source_path, output_dir)
            tally[result] += 1

    finally:
        # CRUCIAL: always quit PowerPoint, success or failure, so we never leave
        # a phantom POWERPNT.EXE process holding memory in the background.
        if powerpoint is not None:
            try:
                powerpoint.Quit()
                print("\nPowerPoint application closed.")
            except Exception as exc:  # noqa: BLE001
                print(f"\n[WARN] PowerPoint.Quit() raised: {exc}")

    # Final summary.
    print("\n===== Summary =====")
    print(f"  Converted: {tally['converted']}")
    print(f"  Skipped  : {tally['skipped']}")
    print(f"  Failed   : {tally['failed']}")
    print("===================")


if __name__ == "__main__":
    # Default folders (relative to the current working directory). Override them
    # from the command line:  python ppt_to_pdf.py <input_dir> <output_dir>
    default_input = os.path.join(os.getcwd(), "input")
    default_output = os.path.join(os.getcwd(), "output")

    input_directory = sys.argv[1] if len(sys.argv) > 1 else default_input
    output_directory = sys.argv[2] if len(sys.argv) > 2 else default_output

    print(f"Input directory : {input_directory}")
    print(f"Output directory: {output_directory}\n")

    batch_convert(input_directory, output_directory)
