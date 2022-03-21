## PDFファイルを扱うWindowsバッチファイル<!-- omit in toc -->


[Home](https://oasis3855.github.io/webpage/) > [Software](https://oasis3855.github.io/webpage/software/index.html) > [Software Download](https://oasis3855.github.io/webpage/software/software-download.html) > [win-scripts](../README.md) > ***batchfile_pdftool*** (this page)

<br />
<br />

Last Updated : Mar. 2022

- [ソフトウエアのダウンロード](#ソフトウエアのダウンロード)
- [概要](#概要)
  - [PDFを画像ファイルに変換.bat](#pdfを画像ファイルに変換bat)
  - [ディレクトリ内のjpgを結合しPDF化.bat](#ディレクトリ内のjpgを結合しpdf化bat)
  - [PDF縮小(Ghostscript).bat](#pdf縮小ghostscriptbat)
- [動作確認](#動作確認)
- [バージョンアップ履歴](#バージョンアップ履歴)
- [ライセンス](#ライセンス)

<br />
<br />

## ソフトウエアのダウンロード

- ![download icon](../readme_pics/soft-ico-download-darkmode.gif)    [このGitHubリポジトリを参照する（ソースコード）](../batchfile_pdftool/)

## 概要

Windowsファイルエクスプローラーのコンテキストメニュー「送る」から呼び出して使う、PDFファイルを扱うツールです。imagemagick と Ghostscript のコマンドライン版を簡単に使うためのものです。

### PDFを画像ファイルに変換.bat

PDFファイルを画像ファイルに変換します。PDFファイル1ページにつき、画像ファイル1つが出力されます。画像ファイルは jpg, png, bmp, tiff から選択可能です。

出力ファイルがjpgの場合は、圧縮率を設定できます。

### ディレクトリ内のjpgを結合しPDF化.bat

指定したディレクトリ内のすべてのjpgファイルを結合し、PDFファイルを作成します。

### PDF縮小(Ghostscript).bat

PDFファイルの圧縮率を変えることで、ファイルサイズの縮小を行います。

圧縮率は、次のものから選択できます。
- screen   (ファイルサイズ小, 高圧縮)
- ebook
- printer
- prepress   (ファイルサイズ大, 低圧縮)

## 動作確認

- Windows 10
- Windows 11

## バージョンアップ履歴

- Version 1.0 (2022/03/21)

## ライセンス

このスクリプトは [GNU General Public License v3ライセンスで公開する](https://gpl.mhatta.org/gpl.ja.html) フリーソフトウエア
