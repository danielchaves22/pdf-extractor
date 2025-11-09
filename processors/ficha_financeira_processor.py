"""Processador dedicado ao modelo de Ficha Financeira."""

from __future__ import annotations

import csv
import re
import unicodedata
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Sequence, Tuple

import pdfplumber


NumberByMonth = Dict[Tuple[int, int], Decimal]
LogCallback = Callable[[str], None]


@dataclass
class MonthColumn:
    """Representa as colunas Comp./Valor de um mês."""

    name: str
    month: int
    comp_center: Optional[float]
    valor_center: Optional[float]


@dataclass
class MonthBlock:
    """Informações de um bloco quadrimestral da ficha."""

    year: int
    months: List[MonthColumn]
    y_start: float
    y_end: float


class FichaFinanceiraProcessor:
    """Responsável por extrair dados da ficha financeira e gerar CSVs."""

    TARGET_CODES: Dict[str, Dict[str, object]] = {
        "1-Salario": {"column": 1},
        "6-Horas": {"column": 1},
        "8-Insalubridade": {"column": 2},
        "3123-Base": {"column": 2},
    }

    MONTH_MAP = {
        "janeiro": 1,
        "fevereiro": 2,
        "marco": 3,
        "março": 3,
        "abril": 4,
        "maio": 5,
        "junho": 6,
        "julho": 7,
        "agosto": 8,
        "setembro": 9,
        "outubro": 10,
        "novembro": 11,
        "dezembro": 12,
    }

    NUMBER_PATTERN = re.compile(r"^\d{1,3}(?:\.\d{3})*,\d+$|^\d+(?:,\d+)?$")

    def __init__(self, log_callback: Optional[LogCallback] = None) -> None:
        self._log_callback = log_callback

    # ------------------------------------------------------------------
    # API pública
    # ------------------------------------------------------------------
    def generate_proventos(
        self,
        pdf_paths: Sequence[Path],
        start_period: date,
        end_period: date,
        output_dir: Path,
    ) -> Dict[str, object]:
        """Processa PDFs e gera o arquivo PROVENTOS.csv."""

        if start_period > end_period:
            raise ValueError("Período inicial não pode ser maior que o final.")

        if not pdf_paths:
            raise ValueError("Informe ao menos um PDF para processamento.")

        aggregated: Dict[str, NumberByMonth] = {
            key: {} for key in self.TARGET_CODES.keys()
        }
        person_name: Optional[str] = None

        for pdf_path in pdf_paths:
            path = Path(pdf_path)
            if not path.exists():
                raise FileNotFoundError(f"Arquivo não encontrado: {path}")

            self._log(f"📄 Lendo {path.name}")
            parse_result = self._parse_pdf(path)

            if not person_name and parse_result["person_name"]:
                person_name = parse_result["person_name"]
            elif (
                person_name
                and parse_result["person_name"]
                and parse_result["person_name"].lower() != person_name.lower()
            ):
                self._log(
                    "⚠️ Nomes diferentes encontrados nos PDFs. Mantendo o primeiro identificado."
                )

            for code, values in parse_result["values"].items():
                target = aggregated.setdefault(code, {})
                for key, value in values.items():
                    if key in target and target[key] != value:
                        self._log(
                            f"⚠️ Valor duplicado para {code} em {key[1]:02d}/{key[0]}: substituindo {target[key]} por {value}."
                        )
                    target[key] = value

        if not person_name:
            person_name = pdf_paths[0].stem
            self._log("⚠️ Nome não encontrado no PDF. Usando nome do arquivo.")

        months = list(self._iterate_months(start_period, end_period))

        proventos_values: List[Tuple[int, int, Decimal]] = []
        for year, month in months:
            value = aggregated.get("3123-Base", {}).get((year, month), Decimal("0"))
            proventos_values.append((year, month, value))
            self._log(
                f"🧮 Mês {month:02d}/{year}: {self._format_decimal(value)} extraído para PROVENTOS."
            )

        output_dir.mkdir(parents=True, exist_ok=True)
        file_slug = self._slugify_name(person_name)
        output_path = output_dir / f"PROVENTOS_{file_slug}.csv"
        self._write_proventos_csv(output_path, proventos_values)

        self._log(f"✅ Arquivo gerado em {output_path}")

        return {
            "person_name": person_name,
            "output_path": output_path,
            "months": proventos_values,
        }

    # ------------------------------------------------------------------
    # Extração de dados
    # ------------------------------------------------------------------
    def _parse_pdf(self, pdf_path: Path) -> Dict[str, object]:
        values: Dict[str, NumberByMonth] = {
            key: {} for key in self.TARGET_CODES.keys()
        }
        person_name: Optional[str] = None

        with pdfplumber.open(str(pdf_path)) as pdf:
            if pdf.pages:
                person_name = self._extract_person_name(pdf.pages[0])

            for page_index, page in enumerate(pdf.pages):
                words = page.extract_words(use_text_flow=True, keep_blank_chars=False)
                if not words:
                    continue

                comp_centers, valor_centers = self._extract_column_centers(words)
                blocks = self._extract_month_blocks(
                    words, page.height, comp_centers, valor_centers
                )

                for block in blocks:
                    for prefix, cfg in self.TARGET_CODES.items():
                        column_index = int(cfg["column"])
                        occurrences = self._find_row_occurrences(words, prefix, block)
                        for row_words in occurrences:
                            extracted = self._extract_values_from_row(
                                row_words, block, column_index
                            )
                            target = values.setdefault(prefix, {})
                            for month_key, amount in extracted.items():
                                existing = target.get(month_key)
                                if existing is not None and existing != amount:
                                    self._log(
                                        f"⚠️ Valores conflitantes para {prefix} em {month_key[1]:02d}/{month_key[0]} (substituindo {existing} por {amount})."
                                    )
                                target[month_key] = amount

        return {"values": values, "person_name": person_name}

    def _find_row_occurrences(
        self, words: List[Dict[str, object]], prefix: str, block: MonthBlock
    ) -> List[List[Dict[str, object]]]:
        rows: List[List[Dict[str, object]]] = []
        tolerance = 0.8

        for word in words:
            text = word.get("text", "")
            if not text.startswith(prefix):
                continue

            row_center = (word["top"] + word["bottom"]) / 2
            if not (block.y_start <= row_center <= block.y_end):
                continue

            row_words = [
                candidate
                for candidate in words
                if abs(((candidate["top"] + candidate["bottom"]) / 2) - row_center)
                <= tolerance
            ]
            rows.append(row_words)

        return rows

    def _extract_values_from_row(
        self, row_words: List[Dict[str, object]], block: MonthBlock, column: int
    ) -> NumberByMonth:
        results: NumberByMonth = {}

        for word in row_words:
            text = word.get("text", "")
            if not self._is_number(text):
                continue

            value = self._to_decimal(text)
            center = (word["x0"] + word["x1"]) / 2

            best_month: Optional[MonthColumn] = None
            best_distance: Optional[float] = None

            for month in block.months:
                target_center = (
                    month.comp_center if column == 1 else month.valor_center
                )
                alternate_center = (
                    month.valor_center if column == 1 else month.comp_center
                )

                if target_center is None and alternate_center is None:
                    continue

                distance = (
                    abs(center - target_center)
                    if target_center is not None
                    else None
                )
                alt_distance = (
                    abs(center - alternate_center)
                    if alternate_center is not None
                    else None
                )

                chosen_center = None
                chosen_distance = None

                if distance is not None and (alt_distance is None or distance <= alt_distance):
                    chosen_center = target_center
                    chosen_distance = distance
                elif alt_distance is not None:
                    chosen_center = alternate_center
                    chosen_distance = alt_distance

                if chosen_center is None:
                    continue

                if chosen_distance is not None and chosen_distance <= 25:
                    if best_distance is None or chosen_distance < best_distance:
                        best_month = month
                        best_distance = chosen_distance

            if best_month is not None:
                results[(block.year, best_month.month)] = value

        return results

    def _extract_column_centers(
        self, words: List[Dict[str, object]]
    ) -> Tuple[List[float], List[float]]:
        comp_centers: List[float] = []
        valor_centers: List[float] = []

        for word in words:
            text = word.get("text", "")
            if text == "Comp.":
                comp_centers.append((word["x0"] + word["x1"]) / 2)
            elif text == "Valor":
                valor_centers.append((word["x0"] + word["x1"]) / 2)

        return comp_centers, valor_centers

    def _extract_month_blocks(
        self,
        words: List[Dict[str, object]],
        page_height: float,
        comp_centers: Sequence[float],
        valor_centers: Sequence[float],
    ) -> List[MonthBlock]:
        month_blocks: List[MonthBlock] = []

        words_sorted = sorted(
            words,
            key=lambda item: ((item["top"] + item["bottom"]) / 2, item["x0"]),
        )

        for word in words_sorted:
            text = word.get("text", "")
            if not (text.isdigit() and len(text) == 4):
                continue

            year = int(text)
            row_center = round((word["top"] + word["bottom"]) / 2, 1)
            same_row = [
                candidate
                for candidate in words_sorted
                if abs(round((candidate["top"] + candidate["bottom"]) / 2, 1) - row_center)
                < 0.2
            ]
            month_names = [item["text"] for item in same_row if item["text"] != text]
            if not month_names:
                continue

            months: List[MonthColumn] = []
            comp_idx = 0
            valor_idx = 0
            for name in month_names:
                normalized = name.strip().lower()
                if normalized == "*totais*":
                    valor_idx += 1
                    continue

                month_number = self.MONTH_MAP.get(normalized)
                if not month_number:
                    continue

                comp_center = (
                    comp_centers[comp_idx] if comp_idx < len(comp_centers) else None
                )
                valor_center = (
                    valor_centers[valor_idx] if valor_idx < len(valor_centers) else None
                )
                months.append(
                    MonthColumn(
                        name=name,
                        month=month_number,
                        comp_center=comp_center,
                        valor_center=valor_center,
                    )
                )
                comp_idx += 1
                valor_idx += 1

            if not months:
                continue

            month_blocks.append(
                MonthBlock(
                    year=year,
                    months=months,
                    y_start=row_center,
                    y_end=page_height,
                )
            )

        month_blocks.sort(key=lambda block: block.y_start)
        for index, block in enumerate(month_blocks):
            next_start = (
                month_blocks[index + 1].y_start if index + 1 < len(month_blocks) else page_height
            )
            block.y_end = next_start - 0.5

        return month_blocks

    def _extract_person_name(self, page: pdfplumber.page.Page) -> Optional[str]:
        text = page.extract_text() or ""
        lines = [line.strip() for line in text.splitlines() if line.strip()]

        for idx, line in enumerate(lines):
            if "Nome" in line and "Matr/Contr" in line:
                if idx + 1 < len(lines):
                    candidate = lines[idx + 1]
                    match = re.match(r"([A-Za-zÀ-ÿ'`\s]+?)\s+\d", candidate)
                    if match:
                        return match.group(1).strip()
                    return candidate.split("  ")[0].strip()
        return None

    # ------------------------------------------------------------------
    # Utilidades
    # ------------------------------------------------------------------
    def _write_proventos_csv(
        self, output_path: Path, months: Iterable[Tuple[int, int, Decimal]]
    ) -> None:
        header = [
            "MES_ANO",
            "VALOR",
            "FGTS",
            "FGTS_REC.",
            "CONTRIBUICAO_SOCIAL",
            "CONTRIBUICAO_SOCIAL_REC.",
            "",
            "",
            "",
            "",
        ]

        with output_path.open("w", newline="", encoding="utf-8") as csv_file:
            writer = csv.writer(csv_file, delimiter=";", lineterminator="\n")
            writer.writerow(header)
            for year, month, value in months:
                mes_ano = f"{month:02d}/{year}"
                formatted = self._format_decimal(value)
                writer.writerow(
                    [
                        mes_ano,
                        formatted,
                        "N",
                        "N",
                        "N",
                        "N",
                        "",
                        "",
                        "",
                        "",
                    ]
                )

    def _iterate_months(self, start: date, end: date) -> Iterable[Tuple[int, int]]:
        current_year = start.year
        current_month = start.month

        while (current_year, current_month) <= (end.year, end.month):
            yield current_year, current_month
            current_month += 1
            if current_month > 12:
                current_month = 1
                current_year += 1

    def _slugify_name(self, name: str) -> str:
        normalized = unicodedata.normalize("NFKD", name)
        ascii_text = "".join(
            ch for ch in normalized if not unicodedata.combining(ch)
        )
        ascii_text = ascii_text.replace(" ", "_")
        ascii_text = re.sub(r"[^A-Za-z0-9_\-]", "", ascii_text)
        return ascii_text or "resultado"

    def _format_decimal(self, value: Decimal) -> str:
        quantized = value.quantize(Decimal("0.01"))
        text = f"{quantized:.2f}".replace(".", ",")
        text = text.rstrip("0").rstrip(",")
        return text or "0"

    def _is_number(self, text: str) -> bool:
        return bool(self.NUMBER_PATTERN.match(text))

    def _to_decimal(self, text: str) -> Decimal:
        cleaned = text.replace(".", "").replace(",", ".")
        try:
            return Decimal(cleaned)
        except InvalidOperation:
            return Decimal("0")

    def _log(self, message: str) -> None:
        if self._log_callback:
            self._log_callback(message)

