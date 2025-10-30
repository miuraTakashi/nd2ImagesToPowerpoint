# インストール方法 / Installation Guide

## 方法1: setup.pyを使用（推奨 / Recommended）

リポジトリのディレクトリで以下のコマンドを実行:

```bash
cd /path/to/nd2ImagesToPowerpoint
pip install -e .
```

これで、どこからでも `nd2ImagesToPowerpoint` コマンドが使えるようになります。

```bash
cd /path/to/nd2/files
nd2ImagesToPowerpoint
```

## 方法2: シェルスクリプトを使用

リポジトリのディレクトリをPATHに追加します。`~/.zshrc`（または`~/.bashrc`）に以下を追加:

```bash
export PATH="$PATH:/path/to/nd2ImagesToPowerpoint"
```

その後、シェルを再起動するか:

```bash
source ~/.zshrc  # または source ~/.bashrc
```

これで、どこからでも `nd2ImagesToPowerpoint` コマンドが使えるようになります。

## 方法3: シンボリックリンクを作成

```bash
sudo ln -s /path/to/nd2ImagesToPowerpoint/nd2ImagesToPowerpoint /usr/local/bin/nd2ImagesToPowerpoint
```

## アンインストール

方法1でインストールした場合:

```bash
pip uninstall nd2images-to-powerpoint
```

