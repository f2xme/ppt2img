# ppt2img

一个适合被服务端程序调用的批量转换 CLI：

- `PPT/PPTX -> PDF` 继续使用 `LibreOffice`
- `PDF -> 图片` 使用 Rust crate `pdfium-render`
- CLI 参数使用 `clap` 管理
- 支持面向程序调用的 `--json` 结构化输出

它的目标是去掉 `pdftoppm` 依赖，但它仍然需要 `Pdfium` 动态库。

## 依赖

运行时需要：

- `libreoffice` 或 `soffice`
- `Pdfium` 动态库

`pdfium-render` 官方文档说明：

- 它不会自带 `Pdfium`
- 最简单的方式是运行时动态绑定系统库，或把 `Pdfium` 动态库和可执行文件放在一起
- `Pdfium` 本身不保证线程安全
- `pdfium-render 0.8.37` 的 `pdfium_latest` 特性当前对应 `Pdfium 7543`

来源：

- [pdfium-render 文档](https://docs.rs/pdfium-render/latest/pdfium_render/)
- [pdfium-binaries 预编译库](https://github.com/bblanchon/pdfium-binaries/releases)

## 安装

### Ubuntu / Debian

先安装 LibreOffice：

```bash
sudo apt-get update
sudo apt-get install -y libreoffice
```

再从官方预编译发布页下载与你系统匹配的 `Pdfium` 动态库，解压后保证能拿到类似：

```bash
libpdfium.so
```

### CentOS / RHEL / Rocky / AlmaLinux

```bash
sudo dnf install -y libreoffice
```

然后同样从官方预编译发布页下载并解压 `libpdfium.so`。

### macOS

```bash
brew install --cask libreoffice
```

然后从官方预编译发布页下载 macOS 对应的包，解压后会得到类似：

```bash
libpdfium.dylib
```

## 运行时放置方式

有两种常见方式：

1. 把 `libpdfium.so` / `libpdfium.dylib` 放在可执行文件旁边
2. 启动时显式传 `--pdfium-lib`

例如：

```bash
./target/release/ppt2img /data/ppt /data/output \
  --pdfium-lib /opt/pdfium/libpdfium.so
```

或者 macOS：

```bash
./target/release/ppt2img /data/ppt /data/output \
  --pdfium-lib /opt/pdfium/libpdfium.dylib
```

如果你传的是目录，程序会自动拼出当前平台对应的库文件名。

## 兼容性提示

不是所有机器里“碰巧带着的 `libpdfium`”都能直接拿来用。

如果 `Pdfium` 版本过旧，运行时可能会报类似下面的错误：

- 缺少 `FPDF_InitLibraryWithConfig`
- 缺少 `FPDFFormObj_RemoveObject`

这通常意味着动态库版本太老，或和当前 `pdfium-render` 绑定的 API 版本不匹配。此时最稳的做法是直接使用 `pdfium-binaries` 发布页里的新版本预编译库。

## 编译

```bash
cd /Users/bran/project/ppt/example/ppt2img
cargo build --release
```

## 用法

```bash
./target/release/ppt2img /data/ppt /data/output
```

查看完整参数：

```bash
./target/release/ppt2img --help
```

常用参数：

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

指定输出格式：

```bash
./target/release/ppt2img /data/ppt /data/output --format webp
```

控制有损格式质量：

```bash
./target/release/ppt2img /data/ppt /data/output --format jpeg --quality 82
./target/release/ppt2img /data/ppt /data/output --format webp --quality 80
```

如果 `Pdfium` 不在系统库路径里，可以手动指定：

```bash
./target/release/ppt2img /data/ppt /data/output \
  --pdfium-lib /path/to/libpdfium.dylib
```

也可以传一个目录，程序会自动在该目录里查找当前平台对应的库文件名。

## 程序调用

推荐服务端或任务队列调用时加 `--json`：

```bash
./target/release/ppt2img /data/demo.pptx /data/output/job-20260428-001 \
  --format webp \
  --quality 80 \
  --json
```

成功时 stdout 只输出 JSON，便于上层程序解析：

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

失败时进程返回非 0，stderr 输出错误；如果使用了 `--json`，stdout 也会输出：

```json
{
  "success": false,
  "message": "failed to bind to Pdfium: ..."
}
```

上层程序建议按以下规则处理：

- 只用退出码判断任务成功或失败
- 成功时解析 stdout 的 `documents[].files`
- 失败时优先记录 stderr；需要结构化错误时解析 stdout 的 `message`
- 每次任务传一个独立的 `OUTPUT_DIR`，例如 `/data/output/<job-id>`，不要让多个任务共享同一个任务目录

## 输出一致性

每次批量任务会先渲染到同级临时目录，所有文档全部成功后再发布本次涉及的文档输出目录。程序不会整体替换 `OUTPUT_DIR`，因此同一个输出根下其它无关目录会被保留。

这意味着：

- 批量任务中任意一个文档失败时，不会发布本次批量输出
- 重跑时，旧的 `slide-*` 文件会被清理，不会混入本次结果
- 如果多个输入会映射到同一个输出目录，例如同目录下同时存在 `foo.pptx` 和 `foo.pdf`，程序会提前报错，不会互相覆盖
- 如果多个输入会形成嵌套输出目录，例如同时存在 `foo.pptx` 和 `foo/bar.pptx`，程序会提前报错，避免发布顺序导致目录互相包含
- 当输出目录位于输入目录内部时，扫描输入文件会自动跳过输出目录，避免把上一次的输出 PDF 当作新输入
- 如果输出路径已经存在且不是目录，程序会提前报错，不会替换普通文件
- 当输入是目录时，`OUTPUT_DIR` 不能和 `INPUT_PATH` 相同，也不能是 `INPUT_PATH` 的上级目录，避免覆盖源文件目录

## 输出格式与质量

当前支持：

- `png`
- `jpeg` / `jpg`
- `webp`

说明：

- `--quality` 对 `jpeg` 和 `webp` 生效
- `png` 是无损格式，当前不使用 `--quality`
- `--quality` 范围是 `1-100`
- 默认格式是 `webp`
- 默认质量是 `80`

## 当前测试结果

本目录版本已经在当前机器上完成了两类测试：

- 直接读取 PDF 并导出 PNG
- 从 `pptx` 经过 `LibreOffice` 导出 PDF，再用 `pdfium-render` 导出 PNG

其中一份 `23` 页的样本在当前机器上的完整链路耗时约：

- `PDF export`: `8.05s`
- `image render`: `2.71s`

这比原先使用 `pdftoppm` 的版本在同样样本上的图片阶段明显更快。
