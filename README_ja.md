# ND2 Images to PowerPoint

Nikon ND2ファイルから蛍光画像を抽出し、PowerPointプレゼンテーションを自動生成するPythonスクリプトです。

## 機能

- **ND2ファイルの自動検出**: 指定ディレクトリ以下（再帰的）から`.nd2`ファイルを検索
- **チャネル自動マッピング**: 
  - DAPI → 青チャネル
  - Alexa 488 antibody / Alexa488 → 緑チャネル
  - Alx568 / Alexa568 → 赤チャネル
  - 明視野チャネル（brightfield/BF/TD等）は自動的に除外
- **PowerPointスライド生成**: 
  - Two Contentレイアウトを使用
  - タイトルに相対パスとファイル名を表示
  - 左側に画像を配置
  - 右側に一辺の長さ（µm）とチャネル情報を箇条書きで表示
- **画像処理オプション**:
  - Z軸方向の最大強度投影（MIP）
  - 時間軸方向の最大強度投影（MIP）
  - パーセンタイルクリッピングによるコントラスト調整
  - 画像のスケーリング

## 必要要件

- Python 3.9 以上
- 以下のPythonパッケージ:
  - `nd2` - ND2ファイル読み込み
  - `numpy` - 数値計算
  - `Pillow` - 画像処理
  - `python-pptx` - PowerPointファイル生成

## インストール

1. リポジトリをクローンまたはダウンロード:

```bash
git clone https://github.com/miuraTakashi/nd2ImagesToPowerpoint.git
cd nd2ImagesToPowerpoint
```

2. 依存パッケージをインストール:

```bash
pip install -r requirements.txt
```

または個別にインストール:

```bash
pip install nd2 numpy Pillow python-pptx
```

## 使い方

### 基本的な使い方

カレントディレクトリとそのサブディレクトリから`.nd2`ファイルを検索し、PowerPointプレゼンテーションを生成:

```bash
python nd2ImagesToPowerpoint.py
```

### オプション指定

```bash
python nd2ImagesToPowerpoint.py \
  --dir /path/to/nd2/files \
  --output MyPresentation.pptx \
  --mip-z \
  --mip-t \
  --clip-percent 0.3 \
  --scale 0.8 \
  --keep-jpgs \
  --verbose
```

### コマンドラインオプション

| オプション | 説明 | デフォルト |
|-----------|------|----------|
| `--dir` | 検索対象ディレクトリ | カレントディレクトリ |
| `--recursive` | サブディレクトリを再帰的に検索 | 有効（デフォルト） |
| `--output` | 出力ファイル名（`.pptx`） | ディレクトリ名`.pptx` |
| `--mip-z` | Z軸方向の最大強度投影を適用 | 無効 |
| `--mip-t` | 時間軸方向の最大強度投影を適用 | 無効 |
| `--clip-percent` | パーセンタイルクリッピングのパーセンテージ（例: 0.3） | 0.0（無効） |
| `--scale` | 画像のスケールファクター（例: 0.5で50%に縮小） | 1.0 |
| `--max-slide-size` | スライド上の画像の最大サイズ（ピクセル） | 1600 |
| `--keep-jpgs` | 中間JPGファイルを保持 | 無効（自動削除） |
| `--jpg-dir` | 中間JPGファイルの保存ディレクトリ | 一時ディレクトリ |
| `--verbose` | 詳細なチャネルマッピング情報を表示 | 無効 |

### 使用例

#### 例1: 基本的な生成

```bash
python nd2ImagesToPowerpoint.py --dir ./sample
```

#### 例2: Z軸投影とコントラスト調整

```bash
python nd2ImagesToPowerpoint.py \
  --dir ./data \
  --mip-z \
  --clip-percent 0.3 \
  --output Results.pptx
```

#### 例3: 画像サイズ調整とデバッグ

```bash
python nd2ImagesToPowerpoint.py \
  --dir ./experiments \
  --scale 0.6 \
  --max-slide-size 1200 \
  --keep-jpgs \
  --verbose
```

## 出力形式

生成されるPowerPointスライドの形式:

- **レイアウト**: Two Content（2列コンテンツ）
- **タイトル**: 相対パス + ファイル名（例: `sample/x20x8_x500_1.nd2`）
- **左側**: 蛍光画像（RGB合成）
- **右側**: 
  - 一辺の長さ（µm/side）
  - チャネル情報（Red/Green/Blueそれぞれのチャネル名）

## チャネルマッピングルール

スクリプトは以下のルールでチャネルを自動的に識別します:

- **青チャネル**: DAPI, Hoechst, 405nm
- **緑チャネル**: Alexa 488 antibody, Alexa488, 488nm, GFP, FITC
- **赤チャネル**: Alx568, Alexa568, 568nm, 561nm, 555nm, 594nm, Cy3, mCherry, Texas Red
- **除外**: brightfield, BF, TD, TL, transmitted, phase, PH, DIC などの明視野チャネル

チャネル名が認識できない場合、フォールバック処理により適切に割り当てられます。

## トラブルシューティング

### 画像が青く表示される

`--verbose`オプションを使用してチャネルマッピングを確認:

```bash
python nd2ImagesToPowerpoint.py --verbose
```

出力されるチャネルマッピング情報を確認し、チャネル名が正しく認識されているか確認してください。

### プレースホルダに画像が挿入されない

スクリプトは自動的にプレースホルダへの挿入を試み、失敗した場合はプレースホルダを削除して画像を配置します。問題が続く場合は、PowerPointテンプレートを確認してください。

### 依存パッケージのエラー

`python-pptx`が見つからないエラーが出る場合:

```bash
pip install python-pptx
```

## ライセンス

このプロジェクトはMITライセンスの下で提供されています。

## 貢献

バグ報告や機能要望は、Issueまたはプルリクエストでお知らせください。

## 関連ファイル

- `nd2ImagesToPowerpoint.py` - メインスクリプト
- `requirements.txt` - 依存パッケージ一覧
- `sample/` - サンプルND2ファイル（テスト用）

