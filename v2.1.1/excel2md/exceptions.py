# -*- coding: utf-8 -*-
"""excel2md 公開 API 用の例外階層。

Issue #16 に伴うライブラリ化で、内部処理の失敗を ``sys.exit`` ではなく
通常の Python 例外として呼び出し元に伝える必要が出たため導入。

CLI (``excel2md.cli.main``) は ``ExcelConversionError`` を捕まえて
従来どおり exit code 2 で終了する。ライブラリ呼び出し
(``ExcelConverter`` / ``convert_to_markdown``) ではプロセスを終了させず
例外として伝播する。
"""
from __future__ import annotations


class ExcelConversionError(Exception):
    """excel2md の変換処理に起因するすべての例外の基底クラス。"""


class WorkbookOpenError(ExcelConversionError):
    """xlsx / xlsm の読み込みに失敗したことを表す例外。

    openpyxl の ``load_workbook`` がファイル不在・破損・非対応形式などで
    例外を上げた場合に送出される。
    """
