# -*- coding: utf-8 -*-
"""ExcelConverter — xlsx を Markdown に変換する高水準 API。

Issue #16 Phase 2。ConversionConfig を受け取り、内部で runner.run を
呼び出す薄いラッパー。bytes 入力にも対応するため、必要に応じて一時
ファイルを経由する。

PR D で runner.run の内部が ConversionConfig ベースに置換された後は、
tempfile 経由のステップは不要になり、bytes / 文字列 / Path を直接
内部に渡せる予定。
"""
from __future__ import annotations

import tempfile
from dataclasses import replace
from pathlib import Path
from typing import Any, Dict, Optional, Union

from .config import ConversionConfig
from .runner import run as _runner_run

InputType = Union[bytes, str, Path]
OutputPathType = Union[str, Path]


class ExcelConverter:
    """xlsx を Markdown に変換する高水準 API。

    Examples:
        >>> from excel2md import ConversionConfig, ExcelConverter
        >>> cfg = ConversionConfig(hyperlink_mode="inline")
        >>> result = ExcelConverter(cfg).convert("spec.xlsx")
        >>> result["markdown"]
        '# 変換結果: spec.xlsx\\n...'
    """

    def __init__(self, config: Optional[ConversionConfig] = None):
        self.config = config if config is not None else ConversionConfig()

    def convert(
        self,
        source: InputType,
        output_path: Optional[OutputPathType] = None,
    ) -> Dict[str, Any]:
        """xlsx を Markdown に変換する。

        Args:
            source: xlsx の中身 (bytes) または既存ファイルへのパス (str / Path)。
            output_path: Markdown の出力先パス。省略時は一時ディレクトリに生成する。
                CSV Markdown モード (config.csv_markdown_enabled=True) では runner 側が
                csv_output_dir を使うため、このパラメータは出力名のヒントとしては
                使われない。

        Returns:
            次の辞書を返す:

            ``{
                "markdown": str,        # 生成された Markdown の中身
                "output_path": str,     # Markdown が実際に書き出されたパス
                "result": str,          # runner.run() の生の戻り値 (path / summary)
            }``

            split_by_sheet=True など複数ファイル出力モードでは、最初に見つかった
            Markdown ファイルが "markdown" / "output_path" に入る（取りこぼしは
            "result" に runner からの multi-line サマリーで残る）。
        """
        # ----- ①入力をディスク上のパスに正規化する -----
        workdir: Optional[Path] = None
        if isinstance(source, (bytes, bytearray)):
            workdir = Path(tempfile.mkdtemp(prefix="excel2md_"))
            input_path = str(workdir / "input.xlsx")
            Path(input_path).write_bytes(bytes(source))
        else:
            input_path = str(source)

        # ----- ②出力先を決める -----
        out_path_str: Optional[str] = str(output_path) if output_path is not None else None
        if out_path_str is None:
            if workdir is None:
                workdir = Path(tempfile.mkdtemp(prefix="excel2md_"))
            out_path_str = str(workdir / (Path(input_path).stem + ".md"))

        # ----- ③csv_output_dir の暫定差し替え -----
        # csv_markdown_enabled=True で csv_output_dir 未設定だと、runner は
        # input_path の親に CSV Markdown を書く。bytes 入力では親が tempdir に
        # なる、というのは妥当だが、明示的に workdir を渡しておくと挙動が
        # 予測しやすい。元 config は不変のまま、replace でコピーを作る。
        cfg = self.config
        if cfg.csv_markdown_enabled and cfg.csv_output_dir is None and workdir is not None:
            cfg = replace(cfg, csv_output_dir=str(workdir))

        # ----- ④runner.run に委譲 -----
        # runner.run は args.foo 形式の属性アクセスしか行わないため、
        # ConversionConfig インスタンスを args として直接渡せる。
        result = _runner_run(input_path, out_path_str, cfg)

        # ----- ⑤実際に書き出された Markdown ファイルを探す -----
        actual_md_path = self._resolve_output_path(result, out_path_str)
        markdown = actual_md_path.read_text(encoding="utf-8") if actual_md_path else ""

        return {
            "markdown": markdown,
            "output_path": str(actual_md_path) if actual_md_path else out_path_str,
            "result": result,
        }

    @staticmethod
    def _resolve_output_path(result: Any, fallback_path: str) -> Optional[Path]:
        """runner.run の戻り値から実際の Markdown ファイルパスを判定する。"""
        # ケース 1: 単一の path 文字列 (通常 Markdown モード / 単一 CSV Markdown モード)
        if isinstance(result, str) and "\n" not in result:
            candidate = Path(result)
            if candidate.exists() and candidate.is_file():
                return candidate

        # ケース 2: split_by_sheet などの multi-line サマリー — 最初の md を採用
        if isinstance(result, str) and "\n" in result:
            for line in result.splitlines()[1:]:  # 先頭行は通知文
                candidate = Path(line.strip())
                if candidate.exists() and candidate.is_file() and candidate.suffix == ".md":
                    return candidate

        # ケース 3: fallback_path にファイルが書かれている (通常 Markdown モードの一般)
        candidate = Path(fallback_path)
        if candidate.exists() and candidate.is_file():
            return candidate

        return None
