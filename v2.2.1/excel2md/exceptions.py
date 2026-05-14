# -*- coding: utf-8 -*-
"""Exception types for the excel2md library API.

Issue #36: ライブラリ層は ``sys.exit`` を呼ばず、例外を上げる。
CLI 経路 (``cli.main``) でこれらを catch して終了コードに翻訳する。
"""


class ExcelConversionError(Exception):
    """Base exception for all excel2md conversion errors.

    ライブラリ利用者はこの基底クラスで一括捕捉できる。
    今後、開ける以外の段階で復帰不能エラーが追加された場合も
    このクラスを継承する。
    """


class WorkbookOpenError(ExcelConversionError):
    """Raised when a workbook cannot be opened (bad path / corrupt bytes / etc.).

    元の openpyxl 例外は ``__cause__`` (``raise ... from e``) として保持される。
    """
