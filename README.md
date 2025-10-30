# ND2 Images to PowerPoint

A Python script that extracts fluorescence images from Nikon ND2 files and automatically generates PowerPoint presentations.

## Features

- **Automatic ND2 file detection**: Recursively searches for `.nd2` files in the specified directory and subdirectories
- **Automatic channel mapping**: 
  - DAPI → Blue channel
  - Alexa 488 antibody / Alexa488 → Green channel
  - Alx568 / Alexa568 → Red channel
  - Brightfield channels (brightfield/BF/TD, etc.) are automatically excluded
- **PowerPoint slide generation**: 
  - Uses Two Content layout
  - Displays relative path and filename in the title
  - Places image on the left side
  - Shows side length (µm) and channel information as bullet points on the right side
- **Image processing options**:
  - Maximum intensity projection (MIP) along Z-axis
  - Maximum intensity projection (MIP) along time axis
  - Contrast adjustment via percentile clipping
  - Image scaling

## Requirements

- Python 3.9 or higher
- The following Python packages:
  - `nd2` - ND2 file reading
  - `numpy` - Numerical computation
  - `Pillow` - Image processing
  - `python-pptx` - PowerPoint file generation

## Installation

1. Clone or download the repository:

```bash
git clone https://github.com/miuraTakashi/nd2ImagesToPowerpoint.git
cd nd2ImagesToPowerpoint
```

2. Install the package (this will also install dependencies and make the command available):

```bash
pip install -e .
```

Or install dependencies only:

```bash
pip install -r requirements.txt
```

Or install individually:

```bash
pip install nd2 numpy Pillow python-pptx
```

After installation with `pip install -e .`, you can use the `nd2ImagesToPowerpoint` command from any directory:

```bash
cd /path/to/nd2/files
nd2ImagesToPowerpoint
```

### Additional Installation Methods

#### Method 2: Using Shell Script Wrapper

Add the repository directory to your PATH. Add the following to `~/.zshrc` (or `~/.bashrc`):

```bash
export PATH="$PATH:/path/to/nd2ImagesToPowerpoint"
```

Then restart your shell or run:

```bash
source ~/.zshrc  # or source ~/.bashrc
```

#### Method 3: Create Symbolic Link

Create a symbolic link to make the command available system-wide:

```bash
sudo ln -s /path/to/nd2ImagesToPowerpoint/nd2ImagesToPowerpoint /usr/local/bin/nd2ImagesToPowerpoint
```

#### Method 4: Run Python Script Directly

If you don't want to install the command, you can always run the Python script directly:

```bash
python /path/to/nd2ImagesToPowerpoint/nd2ImagesToPowerpoint.py
```

Or if you're in the repository directory:

```bash
python nd2ImagesToPowerpoint.py
```

**Note:** Method 1 (`pip install -e .`) is recommended for most users as it properly manages dependencies and provides the cleanest installation. If you encounter permission errors, you may need to use `pip install --user -e .` instead.

## Usage

### Basic Usage

Search for `.nd2` files in the current directory and subdirectories, then generate a PowerPoint presentation:

```bash
nd2ImagesToPowerpoint
```

Or using Python directly:

```bash
python nd2ImagesToPowerpoint.py
```

### With Options

```bash
nd2ImagesToPowerpoint \
  --dir /path/to/nd2/files \
  --output MyPresentation.pptx \
  --mip-z \
  --mip-t \
  --clip-percent 0.3 \
  --scale 0.8 \
  --keep-jpgs \
  --verbose
```

### Command-Line Options

| Option | Description | Default |
|--------|-------------|---------|
| `--dir` | Target directory to search | Current directory |
| `--recursive` | Recursively search subdirectories | Enabled (default) |
| `--output` | Output file name (`.pptx`) | Directory name `.pptx` |
| `--mip-z` | Apply maximum intensity projection along Z-axis | Disabled |
| `--mip-t` | Apply maximum intensity projection along time axis | Disabled |
| `--clip-percent` | Percentile clipping percentage (e.g., 0.3) | 0.0 (disabled) |
| `--scale` | Image scale factor (e.g., 0.5 for 50% reduction) | 1.0 |
| `--max-slide-size` | Maximum image size on slide (pixels) | 1600 |
| `--keep-jpgs` | Keep intermediate JPG files | Disabled (auto-delete) |
| `--jpg-dir` | Directory to store intermediate JPG files | Temporary directory |
| `--verbose` | Display detailed channel mapping information | Disabled |

### Usage Examples

#### Example 1: Basic Generation

```bash
nd2ImagesToPowerpoint --dir ./sample
```

#### Example 2: Z-axis Projection and Contrast Adjustment

```bash
nd2ImagesToPowerpoint \
  --dir ./data \
  --mip-z \
  --clip-percent 0.3 \
  --output Results.pptx
```

#### Example 3: Image Size Adjustment and Debugging

```bash
nd2ImagesToPowerpoint \
  --dir ./experiments \
  --scale 0.6 \
  --max-slide-size 1200 \
  --keep-jpgs \
  --verbose
```

## Output Format

The generated PowerPoint slides have the following format:

- **Layout**: Two Content (2-column content)
- **Title**: Relative path + filename (e.g., `sample/x20x8_x500_1.nd2`)
- **Left side**: Fluorescence image (RGB composite)
- **Right side**: 
  - Side length (µm/side)
  - Channel information (Red/Green/Blue channel names)

## Channel Mapping Rules

The script automatically identifies channels using the following rules:

- **Blue channel**: DAPI, Hoechst, 405nm
- **Green channel**: Alexa 488 antibody, Alexa488, 488nm, GFP, FITC
- **Red channel**: Alx568, Alexa568, 568nm, 561nm, 555nm, 594nm, Cy3, mCherry, Texas Red
- **Excluded**: Brightfield channels such as brightfield, BF, TD, TL, transmitted, phase, PH, DIC

If channel names cannot be recognized, fallback processing assigns them appropriately.

## Troubleshooting

### Images Appear Blue

Use the `--verbose` option to check channel mapping:

```bash
nd2ImagesToPowerpoint --verbose
```

Check the output channel mapping information to verify that channel names are correctly recognized.

### Images Not Inserted into Placeholders

The script automatically attempts to insert images into placeholders, and removes placeholders if insertion fails. If problems persist, check your PowerPoint template.

### Dependency Package Errors

If you get an error that `python-pptx` is not found:

```bash
pip install python-pptx
```

## License

This project is provided under the MIT License.

## Contributing

Please report bugs or feature requests via Issues or Pull Requests.

## Related Files

- `nd2ImagesToPowerpoint.py` - Main script
- `requirements.txt` - Dependency package list
- `sample/` - Sample ND2 files (for testing)

---

[日本語版 / Japanese version](README_ja.md)
