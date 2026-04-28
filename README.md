# ppt2img

A CLI tool for converting PowerPoint and PDF documents into per-page images.

- `PPT/PPTX -> PDF` is handled by LibreOffice
- `PDF -> images` is handled by `pdfium-render`
- CLI arguments are powered by `clap`
- `--json` output is available for server-side and batch job integration

The goal is to avoid a `pdftoppm` runtime dependency while still keeping rendering fast and scriptable. A Pdfium dynamic library is still required at runtime.

[中文文档](README.zh-CN.md)

## Dependencies

Runtime dependencies:

- `libreoffice` or `soffice`
- Pdfium dynamic library

Notes from `pdfium-render`:

- Pdfium is not bundled by the crate.
- The simplest deployment option is to put the Pdfium dynamic library next to the executable, or pass it with `--pdfium-lib`.
- Pdfium itself does not guarantee thread safety.
- `pdfium-render 0.8.37` with `pdfium_latest` currently targets Pdfium `7543`.

References:

- [pdfium-render documentation](https://docs.rs/pdfium-render/latest/pdfium_render/)
- [pdfium-binaries releases](https://github.com/bblanchon/pdfium-binaries/releases)

## Installation

### Ubuntu / Debian

Install LibreOffice:

```bash
sudo apt-get update
sudo apt-get install -y libreoffice
```

Download the matching Pdfium package from the official prebuilt releases and extract:

```bash
libpdfium.so
```

### CentOS / RHEL / Rocky / AlmaLinux

```bash
sudo dnf install -y libreoffice
```

Then download and extract the matching `libpdfium.so` from the official prebuilt releases.

### macOS

```bash
brew install --cask libreoffice
```

Then download the matching macOS Pdfium package. It contains:

```bash
libpdfium.dylib
```

## Pdfium Placement

There are two common deployment options:

1. Put `libpdfium.so`, `libpdfium.dylib`, or `pdfium.dll` next to the executable.
2. Pass the library explicitly with `--pdfium-lib`.

Linux example:

```bash
./target/release/ppt2img /data/ppt /data/output \
  --pdfium-lib /opt/pdfium/libpdfium.so
```

macOS example:

```bash
./target/release/ppt2img /data/ppt /data/output \
  --pdfium-lib /opt/pdfium/libpdfium.dylib
```

If `--pdfium-lib` points to a directory, the program will automatically append the platform-specific library filename.

## Compatibility Notes

Not every bundled or system `libpdfium` is usable.

If the Pdfium version is too old, startup may fail with errors such as:

- missing `FPDF_InitLibraryWithConfig`
- missing `FPDFFormObj_RemoveObject`

This usually means the dynamic library is too old or does not match the API expected by the current `pdfium-render` binding. The most reliable option is to use a recent prebuilt package from `pdfium-binaries`.

## Build

```bash
cd /Users/bran/project/ppt/example/ppt2img
cargo build --release
```

## Usage

```bash
./target/release/ppt2img /data/ppt /data/output
```

Show all options:

```bash
./target/release/ppt2img --help
```

Common options:

```text
Usage: ppt2img [OPTIONS] <INPUT_PATH> [OUTPUT_DIR]

Arguments:
  <INPUT_PATH>  A .ppt/.pptx/.pdf file or a directory to scan recursively
  [OUTPUT_DIR]  Directory where per-document image folders are written [default: ppt_images]

Options:
      --dpi <DPI>              Render density in DPI [default: 200]
      --format <FORMAT>        Output image format [default: webp] [possible values: png, jpeg, webp]
      --quality <QUALITY>      Lossy quality for jpeg/webp. Ignored for png. Defaults to 80
      --keep-pdf               Preserve intermediate PDFs for PPT/PPTX inputs
      --libreoffice <PATH>     LibreOffice/soffice executable path or command name
      --pdfium-lib <PATH>      Path to libpdfium, or a directory containing the platform library
      --json                   Emit a machine-readable JSON report to stdout
```

Select output format:

```bash
./target/release/ppt2img /data/ppt /data/output --format webp
```

Set quality for lossy formats:

```bash
./target/release/ppt2img /data/ppt /data/output --format jpeg --quality 82
./target/release/ppt2img /data/ppt /data/output --format webp --quality 80
```

Specify Pdfium manually:

```bash
./target/release/ppt2img /data/ppt /data/output \
  --pdfium-lib /path/to/libpdfium.dylib
```

You can also pass a directory. The program will look for the current platform's library filename inside it.

## Programmatic Usage

For server-side jobs or task queues, use `--json`:

```bash
./target/release/ppt2img /data/demo.pptx /data/output/job-20260428-001 \
  --format webp \
  --quality 80 \
  --json
```

On success, stdout contains only JSON:

```json
{
  "success": true,
  "input_path": "/data/demo.pptx",
  "output_root": "/data/output/job-20260428-001",
  "documents": [
    {
      "source_path": "/data/demo.pptx",
      "output_dir": "/data/output/job-20260428-001/demo",
      "page_count": 23,
      "dpi": 200,
      "output_format": "webp",
      "quality": 80,
      "pdf_export_ms": 8050,
      "image_render_ms": 2710,
      "intermediate_pdf": null,
      "files": [
        "/data/output/job-20260428-001/demo/slide-01.webp"
      ]
    }
  ]
}
```

On failure, the process exits with a non-zero status and writes the error to stderr. With `--json`, stdout also contains:

```json
{
  "success": false,
  "message": "failed to bind to Pdfium: ..."
}
```

Recommended integration rules:

- Use the exit code as the source of truth for success or failure.
- On success, parse `documents[].files` from stdout.
- On failure, log stderr first; parse stdout `message` only when structured error handling is needed.
- Use a dedicated `OUTPUT_DIR` for each job, such as `/data/output/<job-id>`. Do not share the same job directory across multiple jobs.

## Output Consistency

Each batch first renders into a sibling temporary directory. After every document succeeds, the tool publishes only the document output directories touched by this run. It does not replace the whole `OUTPUT_DIR`, so unrelated directories under the same output root are preserved.

This means:

- If any document in a batch fails, the current batch output is not published.
- Re-running a job removes stale `slide-*` files from that document output directory.
- If multiple inputs map to the same output directory, such as `foo.pptx` and `foo.pdf` in the same directory, the tool fails early.
- If multiple inputs would create nested output directories, such as `foo.pptx` and `foo/bar.pptx`, the tool fails early.
- If the output directory is inside the input directory, input scanning skips the output directory to avoid treating previous output PDFs as new inputs.
- If an output path already exists and is not a directory, the tool fails early and does not replace a regular file.
- When the input is a directory, `OUTPUT_DIR` must not equal `INPUT_PATH` and must not be an ancestor of `INPUT_PATH`.

## Output Formats And Quality

Supported formats:

- `png`
- `jpeg` / `jpg`
- `webp`

Notes:

- `--quality` applies to `jpeg` and `webp`.
- `png` is lossless and ignores `--quality`.
- `--quality` must be between `1` and `100`.
- The default format is `webp`.
- The default quality is `80`.

## Release Builds

The GitHub Actions workflow builds release packages for:

- Linux x64
- macOS arm64
- Windows x64

Tag a release with:

```bash
git tag v0.1.0
git push origin v0.1.0
```

The workflow will create a GitHub Release and upload platform packages containing:

- the `ppt2img` executable
- the matching Pdfium dynamic library
- `README.md`
- `LICENSE`

LibreOffice is not bundled. It must be installed on the target machine.

## Current Test Results

This version has been tested locally with:

- direct PDF input rendered to PNG
- PPTX converted through LibreOffice to PDF, then rendered with Pdfium

One 23-page sample took approximately:

- `PDF export`: `8.05s`
- `image render`: `2.71s`

The image rendering stage was noticeably faster than the older `pdftoppm` based version on the same sample.
