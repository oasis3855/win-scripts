## テキストファイルや画像ファイルを扱うWindows VBScript (Visual Basic Script)<!-- omit in toc -->


[Home](https://oasis3855.github.io/webpage/) > [Software](https://oasis3855.github.io/webpage/software/index.html) > [Software Download](https://oasis3855.github.io/webpage/software/software-download.html) > [win-scripts](../README.md) > ***VB_Script*** (this page)

<br />
<br />

Last Updated : Aug. 2021

- [ソフトウエアのダウンロード](#ソフトウエアのダウンロード)
- [概要](#概要)
  - [exiftool実行メニュー.vbs](#exiftool実行メニューvbs)
  - [テキストエディタを選択して開く.vbs](#テキストエディタを選択して開くvbs)
  - [ファイル名・更新日時変更.vbs](#ファイル名更新日時変更vbs)
- [動作確認](#動作確認)
- [バージョンアップ履歴](#バージョンアップ履歴)
- [ライセンス](#ライセンス)

<br />
<br />

## ソフトウエアのダウンロード

- ![download icon](../readme_pics/soft-ico-download-darkmode.gif)    [このGitHubリポジトリを参照する（ソースコード）](../VB_Script/)

## 概要

Windowsファイルエクスプローラーのコンテキストメニュー「送る」から呼び出して使う、各種ファイルを扱うツールです。

### exiftool実行メニュー.vbs

exiftool のコマンドライン版を簡単に使うためのツールです。次のようなことができます。

- Exifタグ情報を出力
- ファイルのタイムスタンプをExif日時に変更
- ファイル名をExif:CreateDateに変更
- Gimp編集前にMakeタグのPENTAXをPentax Corp.に書き換え
- Gimp編集後にSoftwareとModifyDateの書き戻し修正

### テキストエディタを選択して開く.vbs

複数のテキストエディタをファイルエクスプローラーの「送る」に登録せず、このスクリプトを呼び出すことでテキストエディタを選択できるようにするツール。

### ファイル名・更新日時変更.vbs

ファイル名の大文字・小文字化、タイムスタンプの変更などができます。

- ファイル名を小文字に変更
- ファイル名を大文字に変更
- 最終更新日を現在日時に変更
- 最終更新日を指定した日時に変更

## 動作確認

- Windows 10
- Windows 11

## バージョンアップ履歴

- Version 1.0 (2021/07/19)
  - テキストエディタを選択して開く.vbs 新規
  - ファイル名・更新日時変更.vbs 新規
- Version 1.0 (2021/08/09)
  - exiftool実行メニュー.vbs 新規

## ライセンス

このスクリプトは [GNU General Public License v3ライセンスで公開する](https://gpl.mhatta.org/gpl.ja.html) フリーソフトウエア
