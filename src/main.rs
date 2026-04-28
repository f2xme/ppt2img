use clap::{Parser, ValueEnum, ValueHint};
use image::DynamicImage;
use image::codecs::jpeg::JpegEncoder;
use pdfium_render::prelude::*;
use serde::Serialize;
use std::collections::HashMap;
use std::env;
use std::ffi::OsString;
use std::fmt::{Display, Write as _};
use std::fs;
use std::fs::File;
use std::io::{BufWriter, Write as IoWrite};
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};
use webpx::{Encoder as WebpEncoder, Unstoppable};

type AppResult<T> = Result<T, String>;

#[derive(Debug, Parser)]
#[command(
    name = "ppt2img",
    version,
    about = "Batch convert PPT/PPTX/PDF files to per-page images using LibreOffice and Pdfium."
)]
struct Cli {
    #[arg(
        value_name = "INPUT_PATH",
        value_hint = ValueHint::AnyPath,
        help = "A .ppt/.pptx/.pdf file or a directory to scan recursively"
    )]
    input_path: PathBuf,

    #[arg(
        value_name = "OUTPUT_DIR",
        default_value = "ppt_images",
        value_hint = ValueHint::DirPath,
        help = "Directory where per-document image folders are written"
    )]
    output_root: PathBuf,

    #[arg(
        long,
        default_value_t = 200,
        value_parser = clap::value_parser!(u32).range(1..=1200),
        help = "Render density in DPI"
    )]
    dpi: u32,

    #[arg(
        long,
        value_enum,
        default_value_t = OutputFormat::Webp,
        help = "Output image format"
    )]
    format: OutputFormat,

    #[arg(
        long,
        value_parser = clap::value_parser!(u8).range(1..=100),
        help = "Lossy quality for jpeg/webp. Ignored for png. Defaults to 80"
    )]
    quality: Option<u8>,

    #[arg(long, help = "Preserve intermediate PDFs for PPT/PPTX inputs")]
    keep_pdf: bool,

    #[arg(
        long,
        value_name = "PATH",
        value_hint = ValueHint::CommandName,
        help = "LibreOffice/soffice executable path or command name"
    )]
    libreoffice: Option<OsString>,

    #[arg(
        long = "pdfium-lib",
        value_name = "PATH",
        value_hint = ValueHint::AnyPath,
        help = "Path to libpdfium, or a directory containing the platform library"
    )]
    pdfium_lib: Option<PathBuf>,

    #[arg(long, help = "Emit a machine-readable JSON report to stdout")]
    json: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, ValueEnum, Serialize)]
#[serde(rename_all = "lowercase")]
enum OutputFormat {
    Png,
    Jpeg,
    Webp,
}

impl OutputFormat {
    fn extension(self) -> &'static str {
        match self {
            Self::Png => "png",
            Self::Jpeg => "jpg",
            Self::Webp => "webp",
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Png => "png",
            Self::Jpeg => "jpeg",
            Self::Webp => "webp",
        }
    }
}

impl Display for OutputFormat {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}

#[derive(Debug)]
struct Config {
    input_path: PathBuf,
    output_root: PathBuf,
    dpi: u32,
    output_format: OutputFormat,
    quality: u8,
    keep_pdf: bool,
    libreoffice_bin: OsString,
    pdfium_lib: Option<PathBuf>,
    emit_json: bool,
}

#[derive(Debug)]
struct Job {
    source_path: PathBuf,
    relative_parent: PathBuf,
    stem: String,
}

#[derive(Debug)]
struct RenderedImages {
    page_count: usize,
    file_names: Vec<String>,
}

#[derive(Debug, Serialize)]
struct RunReport {
    success: bool,
    input_path: String,
    output_root: String,
    documents: Vec<DocumentReport>,
}

#[derive(Debug, Serialize)]
struct DocumentReport {
    source_path: String,
    output_dir: String,
    page_count: usize,
    dpi: u32,
    output_format: OutputFormat,
    quality: u8,
    pdf_export_ms: Option<u64>,
    image_render_ms: u64,
    intermediate_pdf: Option<String>,
    files: Vec<String>,
}

#[derive(Debug, Serialize)]
struct ErrorReport<'a> {
    success: bool,
    message: &'a str,
}

fn main() {
    let cli = Cli::parse();
    let emit_json = cli.json;

    match run(cli) {
        Ok(report) => {
            if emit_json {
                print_json_report(&report);
            }
        }
        Err(err) => {
            if emit_json {
                print_json_report(&ErrorReport {
                    success: false,
                    message: &err,
                });
            }
            eprintln!("Error: {err}");
            std::process::exit(1);
        }
    }
}

fn run(cli: Cli) -> AppResult<RunReport> {
    let config = Config::from_cli(cli)?;

    validate_output_root_safety(&config)?;
    let jobs = discover_jobs(&config.input_path, &config.output_root)?;
    if jobs.is_empty() {
        return Err(format!(
            "no .ppt, .pptx, or .pdf files found under {:?}",
            config.input_path
        ));
    }
    validate_unique_output_dirs(&config, &jobs)?;
    validate_output_dirs_not_nested(&config, &jobs)?;

    log_human(&config, format_args!("Found {} document(s)", jobs.len()));

    ensure_parent_dir(&config.output_root)?;
    let staging_root = TempDir::new_sibling(&config.output_root, "batch")?;
    let mut documents = Vec::with_capacity(jobs.len());
    for job in &jobs {
        documents.push(convert_job(
            &config,
            job,
            staging_root.path(),
            &config.output_root,
        )?);
    }
    publish_document_dirs(staging_root.path(), &config.output_root, &jobs)?;

    log_human(
        &config,
        format_args!("All done. Images are in {:?}", config.output_root),
    );

    Ok(RunReport {
        success: true,
        input_path: path_to_string(&config.input_path),
        output_root: path_to_string(&config.output_root),
        documents,
    })
}

impl Config {
    fn from_cli(cli: Cli) -> AppResult<Self> {
        let input_path = absolutize_existing_path(&cli.input_path)?;
        let output_root = absolutize_output_path(&cli.output_root)?;

        if output_root.exists() && !output_root.is_dir() {
            return Err(format!(
                "output path exists but is not a directory: {:?}",
                output_root
            ));
        }

        let quality_explicit = cli.quality.is_some();
        let quality = cli.quality.unwrap_or(80);
        if cli.format == OutputFormat::Png && quality_explicit {
            eprintln!("Warning: --quality is ignored when --format png");
        }

        let libreoffice_bin = cli
            .libreoffice
            .unwrap_or_else(|| resolve_first_available_binary(&["libreoffice", "soffice"]));

        Ok(Self {
            input_path,
            output_root,
            dpi: cli.dpi,
            output_format: cli.format,
            quality,
            keep_pdf: cli.keep_pdf,
            libreoffice_bin,
            pdfium_lib: cli.pdfium_lib,
            emit_json: cli.json,
        })
    }
}

fn discover_jobs(input_path: &Path, output_root: &Path) -> AppResult<Vec<Job>> {
    let mut jobs = Vec::new();

    if input_path.is_file() {
        if is_supported_input(input_path) {
            jobs.push(Job {
                source_path: input_path.to_path_buf(),
                relative_parent: PathBuf::new(),
                stem: file_stem_string(input_path)?,
            });
        }
        return Ok(jobs);
    }

    visit_dir(input_path, input_path, output_root, &mut jobs)?;
    jobs.sort_by(|a, b| a.source_path.cmp(&b.source_path));
    Ok(jobs)
}

fn visit_dir(
    root: &Path,
    current: &Path,
    output_root: &Path,
    jobs: &mut Vec<Job>,
) -> AppResult<()> {
    for entry in fs::read_dir(current)
        .map_err(|err| format!("failed to read directory {:?}: {err}", current))?
    {
        let entry = entry.map_err(|err| format!("failed to read directory entry: {err}"))?;
        let path = entry.path();
        let metadata = entry
            .metadata()
            .map_err(|err| format!("failed to inspect {:?}: {err}", path))?;

        if metadata.is_dir() {
            if should_skip_scan_dir(&path, output_root) {
                continue;
            }
            visit_dir(root, &path, output_root, jobs)?;
            continue;
        }

        if metadata.is_file() && is_supported_input(&path) {
            let relative_parent = path
                .strip_prefix(root)
                .map_err(|err| format!("failed to strip prefix from {:?}: {err}", path))?
                .parent()
                .unwrap_or_else(|| Path::new(""))
                .to_path_buf();

            jobs.push(Job {
                source_path: path.clone(),
                relative_parent,
                stem: file_stem_string(&path)?,
            });
        }
    }

    Ok(())
}

fn is_supported_input(path: &Path) -> bool {
    path.extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| matches!(ext.to_ascii_lowercase().as_str(), "ppt" | "pptx" | "pdf"))
        .unwrap_or(false)
}

fn is_pdf(path: &Path) -> bool {
    path.extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.eq_ignore_ascii_case("pdf"))
        .unwrap_or(false)
}

fn should_skip_scan_dir(path: &Path, output_root: &Path) -> bool {
    if paths_refer_to_same_location(path, output_root) {
        return true;
    }

    let Some(output_name) = output_root.file_name().and_then(|name| name.to_str()) else {
        return false;
    };
    let Some(candidate_name) = path.file_name().and_then(|name| name.to_str()) else {
        return false;
    };

    candidate_name.starts_with(&format!(".{output_name}.batch-"))
        || candidate_name.starts_with(&format!(".{output_name}.backup-"))
}

fn paths_refer_to_same_location(a: &Path, b: &Path) -> bool {
    match (fs::canonicalize(a), fs::canonicalize(b)) {
        (Ok(a), Ok(b)) => a == b,
        _ => a == b,
    }
}

fn validate_unique_output_dirs(config: &Config, jobs: &[Job]) -> AppResult<()> {
    let mut sources_by_output: HashMap<PathBuf, Vec<PathBuf>> = HashMap::new();

    for job in jobs {
        sources_by_output
            .entry(output_dir_for_job(&config.output_root, job))
            .or_default()
            .push(job.source_path.clone());
    }

    let conflicts: Vec<String> = sources_by_output
        .into_iter()
        .filter(|(_, sources)| sources.len() > 1)
        .map(|(output_dir, sources)| {
            let sources = sources
                .iter()
                .map(|source| path_to_string(source))
                .collect::<Vec<_>>()
                .join(", ");
            format!("{:?} <= {sources}", output_dir)
        })
        .collect();

    if conflicts.is_empty() {
        Ok(())
    } else {
        Err(format!(
            "multiple inputs would write to the same output directory: {}",
            conflicts.join(" | ")
        ))
    }
}

fn validate_output_dirs_not_nested(config: &Config, jobs: &[Job]) -> AppResult<()> {
    let mut outputs = jobs
        .iter()
        .map(|job| {
            (
                output_dir_for_job(&config.output_root, job),
                &job.source_path,
            )
        })
        .collect::<Vec<_>>();
    outputs.sort_by(|a, b| a.0.cmp(&b.0));

    let mut conflicts = Vec::new();
    for window in outputs.windows(2) {
        let [(parent_dir, parent_source), (child_dir, child_source)] = window else {
            continue;
        };

        if child_dir != parent_dir && child_dir.starts_with(parent_dir) {
            conflicts.push(format!(
                "{:?} from {:?} contains {:?} from {:?}",
                parent_dir, parent_source, child_dir, child_source
            ));
        }
    }

    if conflicts.is_empty() {
        Ok(())
    } else {
        Err(format!(
            "multiple inputs would create nested output directories: {}",
            conflicts.join(" | ")
        ))
    }
}

fn validate_output_root_safety(config: &Config) -> AppResult<()> {
    if !config.input_path.is_dir() {
        return Ok(());
    }

    if config.output_root == config.input_path {
        return Err(format!(
            "output directory must not be the same as input directory: {:?}",
            config.output_root
        ));
    }

    if config.input_path.starts_with(&config.output_root) {
        return Err(format!(
            "output directory must not be an ancestor of input directory: output={:?}, input={:?}",
            config.output_root, config.input_path
        ));
    }

    Ok(())
}

fn output_dir_for_job(output_root: &Path, job: &Job) -> PathBuf {
    output_root.join(&job.relative_parent).join(&job.stem)
}

fn publish_document_dirs(staging_root: &Path, output_root: &Path, jobs: &[Job]) -> AppResult<()> {
    fs::create_dir_all(output_root)
        .map_err(|err| format!("failed to create output directory {:?}: {err}", output_root))?;

    let mut moves = Vec::with_capacity(jobs.len());
    for job in jobs {
        let staged_dir = output_dir_for_job(staging_root, job);
        let final_dir = output_dir_for_job(output_root, job);

        if !staged_dir.is_dir() {
            return Err(format!(
                "staged output directory is missing or not a directory: {:?}",
                staged_dir
            ));
        }
        if final_dir.exists() && !final_dir.is_dir() {
            return Err(format!(
                "document output path exists but is not a directory: {:?}",
                final_dir
            ));
        }
        if let Some(parent) = final_dir.parent() {
            fs::create_dir_all(parent).map_err(|err| {
                format!(
                    "failed to create output parent directory {:?}: {err}",
                    parent
                )
            })?;
        }

        moves.push(PublishMove {
            staged_dir,
            final_dir,
            backup_dir: None,
        });
    }

    publish_moves(&mut moves)
}

struct PublishMove {
    staged_dir: PathBuf,
    final_dir: PathBuf,
    backup_dir: Option<PathBuf>,
}

fn publish_moves(moves: &mut [PublishMove]) -> AppResult<()> {
    let mut backup_count = 0_usize;
    for index in 0..moves.len() {
        if !moves[index].final_dir.exists() {
            continue;
        }

        let backup_dir = backup_path_for(&moves[index].final_dir)?;
        if let Err(err) = fs::rename(&moves[index].final_dir, &backup_dir) {
            restore_backups(moves, backup_count);
            return Err(format!(
                "failed to move existing output {:?} to backup {:?}: {err}",
                moves[index].final_dir, backup_dir
            ));
        }

        moves[index].backup_dir = Some(backup_dir);
        backup_count = index + 1;
    }

    let mut published_count = 0_usize;
    for index in 0..moves.len() {
        if let Err(err) = fs::rename(&moves[index].staged_dir, &moves[index].final_dir) {
            let rollback_errors = rollback_publish(moves, published_count);
            let mut message = format!(
                "failed to move rendered output from {:?} to {:?}: {err}",
                moves[index].staged_dir, moves[index].final_dir
            );
            if !rollback_errors.is_empty() {
                let _ = write!(
                    message,
                    "; rollback also failed: {}",
                    rollback_errors.join(" | ")
                );
            }
            return Err(message);
        }

        published_count = index + 1;
    }

    for publish_move in moves {
        if let Some(backup_dir) = &publish_move.backup_dir
            && let Err(err) = remove_path(backup_dir)
        {
            eprintln!(
                "Warning: failed to remove previous output backup {:?}: {err}",
                backup_dir
            );
        }
    }

    Ok(())
}

fn rollback_publish(moves: &[PublishMove], published_count: usize) -> Vec<String> {
    let mut errors = Vec::new();

    for publish_move in moves[..published_count].iter().rev() {
        if let Err(err) = remove_path(&publish_move.final_dir) {
            errors.push(format!(
                "failed to remove partially published output {:?}: {err}",
                publish_move.final_dir
            ));
        }
    }

    errors.extend(restore_backups(moves, moves.len()));
    errors
}

fn restore_backups(moves: &[PublishMove], count: usize) -> Vec<String> {
    let mut errors = Vec::new();

    for publish_move in moves[..count].iter().rev() {
        let Some(backup_dir) = &publish_move.backup_dir else {
            continue;
        };

        if publish_move.final_dir.exists() {
            continue;
        }

        if let Err(err) = fs::rename(backup_dir, &publish_move.final_dir) {
            errors.push(format!(
                "failed to restore previous output from backup {:?} to {:?}: {err}",
                backup_dir, publish_move.final_dir
            ));
        }
    }

    errors
}

fn backup_path_for(final_dir: &Path) -> AppResult<PathBuf> {
    Ok(final_dir.with_file_name(format!(
        ".{}.backup-{}",
        final_dir
            .file_name()
            .and_then(|name| name.to_str())
            .unwrap_or("output"),
        unique_suffix()?
    )))
}

fn file_stem_string(path: &Path) -> AppResult<String> {
    path.file_stem()
        .and_then(|name| name.to_str())
        .map(|name| name.to_string())
        .ok_or_else(|| format!("invalid file name: {:?}", path))
}

fn absolutize_existing_path(path: &Path) -> AppResult<PathBuf> {
    fs::canonicalize(path).map_err(|err| format!("failed to resolve input path {:?}: {err}", path))
}

fn absolutize_output_path(path: &Path) -> AppResult<PathBuf> {
    if path.exists() {
        return fs::canonicalize(path)
            .map_err(|err| format!("failed to resolve output path {:?}: {err}", path));
    }

    let absolute = if path.is_absolute() {
        path.to_path_buf()
    } else {
        env::current_dir()
            .map_err(|err| format!("failed to read current directory: {err}"))?
            .join(path)
    };

    normalize_path(&absolute)
}

fn normalize_path(path: &Path) -> AppResult<PathBuf> {
    let mut normalized = PathBuf::new();

    for component in path.components() {
        match component {
            std::path::Component::CurDir => {}
            std::path::Component::ParentDir => {
                if !normalized.pop() {
                    return Err(format!("failed to normalize path {:?}", path));
                }
            }
            _ => normalized.push(component.as_os_str()),
        }
    }

    Ok(normalized)
}

fn resolve_first_available_binary(candidates: &[&str]) -> OsString {
    for candidate in candidates {
        if binary_exists(candidate) {
            return OsString::from(candidate);
        }
    }

    OsString::from(candidates[0])
}

fn binary_exists(name: &str) -> bool {
    if name.contains(std::path::MAIN_SEPARATOR) {
        return Path::new(name).exists();
    }

    env::var_os("PATH")
        .map(|paths| {
            env::split_paths(&paths).any(|dir| {
                let full_path = dir.join(name);
                full_path.is_file()
            })
        })
        .unwrap_or(false)
}

fn bind_pdfium(path: Option<&Path>) -> AppResult<Pdfium> {
    let bindings = match path {
        Some(path) if path.is_dir() => {
            let candidate = Pdfium::pdfium_platform_library_name_at_path(path);
            Pdfium::bind_to_library(candidate)
        }
        Some(path) => Pdfium::bind_to_library(path),
        None => Pdfium::bind_to_library(Pdfium::pdfium_platform_library_name_at_path("./"))
            .or_else(|_| Pdfium::bind_to_system_library()),
    }
    .map_err(|err| format!("failed to bind to Pdfium: {err}"))?;

    Ok(Pdfium::new(bindings))
}

fn convert_job(
    config: &Config,
    job: &Job,
    staging_root: &Path,
    final_root: &Path,
) -> AppResult<DocumentReport> {
    let staging_out_dir = output_dir_for_job(staging_root, job);
    let final_out_dir = output_dir_for_job(final_root, job);
    fs::create_dir_all(&staging_out_dir).map_err(|err| {
        format!(
            "failed to create staging output directory {:?}: {err}",
            staging_out_dir
        )
    })?;

    log_human(config, format_args!("Converting {:?}", job.source_path));

    let mut _pdf_temp_dir = None;
    let mut pdf_export_ms = None;
    let pdf_path = if is_pdf(&job.source_path) {
        job.source_path.clone()
    } else {
        let dir = TempDir::new_system("ppt2img")?;
        let (path, elapsed) = export_presentation_to_pdf(config, job, dir.path())?;
        pdf_export_ms = Some(duration_ms(elapsed));
        _pdf_temp_dir = Some(dir);
        path
    };

    let render_started = Instant::now();
    let pdfium = bind_pdfium(config.pdfium_lib.as_deref())?;
    let rendered = render_pdf_to_images(&pdfium, &pdf_path, &staging_out_dir, config)?;
    let image_render_ms = duration_ms(render_started.elapsed());

    let intermediate_pdf = if !is_pdf(&job.source_path) && config.keep_pdf {
        let staged_pdf = staging_out_dir.join(format!("{}.pdf", job.stem));
        fs::copy(&pdf_path, &staged_pdf).map_err(|err| {
            format!(
                "failed to preserve intermediate PDF from {:?} to {:?}: {err}",
                pdf_path, staged_pdf
            )
        })?;
        let final_pdf = final_out_dir.join(format!("{}.pdf", job.stem));
        Some(path_to_string(&final_pdf))
    } else {
        None
    };

    log_human(
        config,
        format_args!(
            "  image render: {}ms (pages={}, dpi={}, format={}, quality={})",
            image_render_ms, rendered.page_count, config.dpi, config.output_format, config.quality
        ),
    );
    log_human(config, format_args!("Done -> {:?}", final_out_dir));

    Ok(DocumentReport {
        source_path: path_to_string(&job.source_path),
        output_dir: path_to_string(&final_out_dir),
        page_count: rendered.page_count,
        dpi: config.dpi,
        output_format: config.output_format,
        quality: config.quality,
        pdf_export_ms,
        image_render_ms,
        intermediate_pdf,
        files: rendered
            .file_names
            .iter()
            .map(|file_name| path_to_string(&final_out_dir.join(file_name)))
            .collect(),
    })
}

fn export_presentation_to_pdf(
    config: &Config,
    job: &Job,
    temp_dir: &Path,
) -> AppResult<(PathBuf, Duration)> {
    let temp_pdf = temp_dir.join(format!("{}.pdf", job.stem));
    let export_started = Instant::now();

    let mut libreoffice = Command::new(&config.libreoffice_bin);
    libreoffice
        .arg("--headless")
        .arg("--convert-to")
        .arg("pdf")
        .arg("--outdir")
        .arg(temp_dir)
        .arg(&job.source_path);

    run_command(&mut libreoffice, "libreoffice")?;

    if !temp_pdf.exists() {
        return Err(format!(
            "libreoffice finished but did not create {:?}",
            temp_pdf
        ));
    }

    let elapsed = export_started.elapsed();
    log_human(
        config,
        format_args!("  PDF export: {}ms", duration_ms(elapsed)),
    );
    Ok((temp_pdf, elapsed))
}

fn render_pdf_to_images(
    pdfium: &Pdfium,
    pdf_path: &Path,
    out_dir: &Path,
    config: &Config,
) -> AppResult<RenderedImages> {
    let document = pdfium
        .load_pdf_from_file(pdf_path, None)
        .map_err(|err| format!("failed to open PDF {:?}: {err}", pdf_path))?;

    let mut file_names = Vec::new();
    let mut page_count = 0_usize;

    for (index, page) in document.pages().iter().enumerate() {
        let width = render_target_for_page(&page, config.dpi);
        let image = page
            .render_with_config(
                &PdfRenderConfig::new()
                    .set_target_width(width)
                    .render_form_data(true),
            )
            .map_err(|err| {
                format!(
                    "failed to render page {} in {:?}: {err}",
                    index + 1,
                    pdf_path
                )
            })?
            .as_image();

        let output_path = out_dir.join(format!(
            "slide-{number:02}.{ext}",
            number = index + 1,
            ext = config.output_format.extension()
        ));
        write_image(&image, &output_path, config.output_format, config.quality)?;

        page_count = index + 1;
        file_names.push(
            output_path
                .file_name()
                .and_then(|name| name.to_str())
                .ok_or_else(|| format!("invalid output file name: {:?}", output_path))?
                .to_string(),
        );
    }

    Ok(RenderedImages {
        page_count,
        file_names,
    })
}

fn write_image(
    image: &DynamicImage,
    output_path: &Path,
    format: OutputFormat,
    quality: u8,
) -> AppResult<()> {
    match format {
        OutputFormat::Png => image
            .save(output_path)
            .map_err(|err| format!("failed to save {:?}: {err}", output_path)),
        OutputFormat::Jpeg => write_jpeg(image, output_path, quality),
        OutputFormat::Webp => write_webp(image, output_path, quality),
    }
}

fn write_jpeg(image: &DynamicImage, output_path: &Path, quality: u8) -> AppResult<()> {
    let file = File::create(output_path)
        .map_err(|err| format!("failed to create {:?}: {err}", output_path))?;
    let writer = BufWriter::new(file);
    let mut encoder = JpegEncoder::new_with_quality(writer, quality);
    let rgb = image.to_rgb8();
    encoder
        .encode_image(&rgb)
        .map_err(|err| format!("failed to encode {:?} as jpeg: {err}", output_path))
}

fn write_webp(image: &DynamicImage, output_path: &Path, quality: u8) -> AppResult<()> {
    let rgba = image.to_rgba8();
    let encoded = WebpEncoder::new_rgba(rgba.as_raw(), rgba.width(), rgba.height())
        .quality(quality as f32)
        .encode_owned(Unstoppable)
        .map_err(|err| format!("failed to encode {:?} as webp: {err}", output_path))?;

    fs::write(output_path, &*encoded)
        .map_err(|err| format!("failed to write {:?}: {err}", output_path))
}

fn render_target_for_page(page: &PdfPage, dpi: u32) -> i32 {
    let width_points = page.width().value;
    ((width_points / 72.0) * dpi as f32).round().max(1.0) as i32
}

struct TempDir {
    path: PathBuf,
    cleanup: bool,
}

impl TempDir {
    fn new_system(prefix: &str) -> AppResult<Self> {
        let mut path = env::temp_dir();
        path.push(unique_dir_name(prefix)?);
        create_temp_dir(path)
    }

    fn new_sibling(final_path: &Path, purpose: &str) -> AppResult<Self> {
        let parent = final_path
            .parent()
            .ok_or_else(|| format!("failed to resolve parent directory for {:?}", final_path))?;
        let stem = final_path
            .file_name()
            .and_then(|name| name.to_str())
            .unwrap_or("output");
        create_temp_dir(parent.join(unique_dir_name(&format!(".{stem}.{purpose}"))?))
    }

    fn path(&self) -> &Path {
        &self.path
    }
}

impl Drop for TempDir {
    fn drop(&mut self) {
        if self.cleanup {
            let _ = fs::remove_dir_all(&self.path);
        }
    }
}

fn create_temp_dir(path: PathBuf) -> AppResult<TempDir> {
    fs::create_dir_all(&path)
        .map_err(|err| format!("failed to create temp directory {:?}: {err}", path))?;
    Ok(TempDir {
        path,
        cleanup: true,
    })
}

fn unique_dir_name(prefix: &str) -> AppResult<String> {
    let mut temp_dir = env::temp_dir();
    let now = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map_err(|err| format!("system clock error: {err}"))?
        .as_nanos();
    temp_dir.push(format!("{prefix}-{}-{now}", std::process::id()));

    Ok(temp_dir
        .file_name()
        .and_then(|name| name.to_str())
        .unwrap_or(prefix)
        .to_string())
}

fn ensure_parent_dir(path: &Path) -> AppResult<()> {
    let Some(parent) = path.parent() else {
        return Ok(());
    };
    if parent.as_os_str().is_empty() {
        return Ok(());
    }

    fs::create_dir_all(parent)
        .map_err(|err| format!("failed to create parent directory {:?}: {err}", parent))
}

fn remove_path(path: &Path) -> std::io::Result<()> {
    if path.is_dir() {
        fs::remove_dir_all(path)
    } else {
        fs::remove_file(path)
    }
}

fn unique_suffix() -> AppResult<String> {
    let now = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map_err(|err| format!("system clock error: {err}"))?
        .as_nanos();
    Ok(format!("{}-{now}", std::process::id()))
}

fn run_command(command: &mut Command, name: &str) -> AppResult<()> {
    let output = command
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output()
        .map_err(|err| format!("failed to start {name}: {err}"))?;

    if output.status.success() {
        return Ok(());
    }

    let mut details = String::new();
    let stdout = String::from_utf8_lossy(&output.stdout).trim().to_string();
    let stderr = String::from_utf8_lossy(&output.stderr).trim().to_string();

    if !stdout.is_empty() {
        let _ = write!(details, "stdout: {stdout}");
    }
    if !stderr.is_empty() {
        if !details.is_empty() {
            details.push_str(" | ");
        }
        let _ = write!(details, "stderr: {stderr}");
    }

    if details.is_empty() {
        Err(format!("{name} exited with status {}", output.status))
    } else {
        Err(format!(
            "{name} exited with status {} ({details})",
            output.status
        ))
    }
}

fn duration_ms(duration: Duration) -> u64 {
    duration.as_millis().min(u64::MAX as u128) as u64
}

fn path_to_string(path: &Path) -> String {
    path.to_string_lossy().into_owned()
}

fn log_human(config: &Config, args: std::fmt::Arguments<'_>) {
    if !config.emit_json {
        println!("{args}");
    }
}

fn print_json_report<T: Serialize>(report: &T) {
    match serde_json::to_string_pretty(report) {
        Ok(json) => {
            let mut stdout = std::io::stdout().lock();
            if let Err(err) = writeln!(stdout, "{json}").and_then(|_| stdout.flush()) {
                eprintln!("Error: failed to write JSON report: {err}");
                std::process::exit(1);
            }
        }
        Err(err) => {
            eprintln!("Error: failed to serialize JSON report: {err}");
            std::process::exit(1);
        }
    }
}
