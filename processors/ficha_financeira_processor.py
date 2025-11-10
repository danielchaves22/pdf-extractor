"""Processador dedicado ao modelo de Ficha Financeira."""

from __future__ import annotations

import csv
import re
import threading
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Sequence, Set, Tuple

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


@dataclass
class ActiveBlock:
    """Representa um bloco que pode se estender para a próxima página."""

    block: MonthBlock
    carry_count: int = 0


class FichaFinanceiraProcessor:
    """Responsável por extrair dados da ficha financeira e gerar CSVs."""

    TARGET_CODES: Dict[str, Dict[str, object]] = {
        "1-Salario": {"column": 1},
        "6-Horas": {"column": 1},
        "8-Insalubridade": {"column": 2},
        "3123-Base": {"column": 2},
        "173-Ferias": {"column": 2, "search_prefix": "173"},
        "174-Ferias": {"column": 2, "search_prefix": "174"},
        "527-INSS-Comp": {"column": 1, "search_prefix": "527"},
        "527-INSS-Valor": {"column": 2, "search_prefix": "527"},
    }

    OUTPUT_SPECS: Tuple[Dict[str, str], ...] = (
        {"code": "3123-Base", "label": "PROVENTOS", "log": "PROVENTOS"},
        {
            "code": "8-Insalubridade",
            "label": "ADIC. INSALUBRIDADE PAGO",
            "log": "INSALUBRIDADE",
        },
    )

    MAX_BLOCK_CARRY = 3

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
        *,
        max_workers: int = 1,
    ) -> Dict[str, object]:
        """Processa PDFs e gera o arquivo PROVENTOS.csv."""

        result = self.generate_csvs(
            pdf_paths=pdf_paths,
            start_period=start_period,
            end_period=end_period,
            output_dir=output_dir,
            max_workers=max_workers,
        )

        for output in result["outputs"]:
            if output["label"] == "PROVENTOS":
                return {
                    "person_name": result["person_name"],
                    "output_path": output["path"],
                    "months": output["months"],
                }

        raise RuntimeError("Arquivo PROVENTOS não foi gerado.")

    def generate_csvs(
        self,
        pdf_paths: Sequence[Path],
        start_period: date,
        end_period: date,
        output_dir: Path,
        *,
        max_workers: int = 1,
    ) -> Dict[str, object]:
        """Gera todos os CSVs configurados para a ficha financeira."""

        if start_period > end_period:
            raise ValueError("Período inicial não pode ser maior que o final.")

        aggregated, person_name = self._aggregate_pdfs(pdf_paths, max_workers=max_workers)

        self._apply_vacation_adjustments(aggregated)

        months_range = list(self._iterate_months(start_period, end_period))
        output_dir_path = Path(output_dir)
        output_dir_path.mkdir(parents=True, exist_ok=True)

        folder_slug = self._slugify_name(person_name)
        period_folder = f"{start_period.year:04d}_{start_period.month:02d}_a_{end_period.year:04d}_{end_period.month:02d}"

        target_dir = output_dir_path / folder_slug / period_folder
        target_dir.mkdir(parents=True, exist_ok=True)

        outputs: List[Dict[str, object]] = []

        for spec in self.OUTPUT_SPECS:
            values = self._collect_values_for_code(
                aggregated,
                spec["code"],
                months_range,
                spec["log"],
            )
            output_path = target_dir / f"{spec['label']}_{folder_slug}.csv"
            self._write_output_csv(output_path, values)
            self._log(f"✅ Arquivo gerado em {output_path}")
            outputs.append({
                "label": spec["label"],
                "path": output_path,
                "months": values,
            })

        return {
            "person_name": person_name,
            "output_folder": target_dir,
            "outputs": outputs,
        }

    # ------------------------------------------------------------------
    # Extração de dados
    # ------------------------------------------------------------------
    def _aggregate_pdfs(
        self,
        pdf_paths: Sequence[Path],
        max_workers: int = 1,
    ) -> Tuple[Dict[str, NumberByMonth], str]:
        if not pdf_paths:
            raise ValueError("Informe ao menos um PDF para processamento.")

        aggregated: Dict[str, NumberByMonth] = {
            key: {} for key in self.TARGET_CODES.keys()
        }
        person_name: Optional[str] = None

        resolved_paths: List[Path] = []
        for pdf_path in pdf_paths:
            path = Path(pdf_path)
            if not path.exists():
                raise FileNotFoundError(f"Arquivo não encontrado: {path}")
            resolved_paths.append(path)

        effective_workers = max(1, min(max_workers, len(resolved_paths)))
        merge_lock = threading.Lock()

        def merge_results(path: Path, parse_result: Dict[str, object]) -> None:
            nonlocal person_name
            with merge_lock:
                parsed_name = parse_result.get("person_name")

                if parsed_name:
                    if not person_name:
                        person_name = parsed_name
                    elif parsed_name.lower() != person_name.lower():
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

        if effective_workers == 1:
            for path in resolved_paths:
                self._log(f"📄 Lendo {path.name}")
                parse_result = self._parse_pdf(path)
                merge_results(path, parse_result)
        else:
            def parse_path(path: Path) -> Tuple[Path, Dict[str, object]]:
                self._log(f"📄 Lendo {path.name}")
                return path, self._parse_pdf(path)

            with ThreadPoolExecutor(max_workers=effective_workers) as executor:
                futures = {
                    executor.submit(parse_path, path): path
                    for path in resolved_paths
                }

                for future in as_completed(futures):
                    path, parse_result = future.result()
                    merge_results(path, parse_result)

        if not person_name:
            person_name = Path(pdf_paths[0]).stem
            self._log("⚠️ Nome não encontrado no PDF. Usando nome do arquivo.")

        return aggregated, person_name

    def _parse_pdf(self, pdf_path: Path) -> Dict[str, object]:
        values: Dict[str, NumberByMonth] = {
            key: {} for key in self.TARGET_CODES.keys()
        }
        person_name: Optional[str] = None

        with pdfplumber.open(str(pdf_path)) as pdf:
            if pdf.pages:
                person_name = self._extract_person_name(pdf.pages[0])

            pending_blocks: List[ActiveBlock] = []
            last_comp_centers: List[float] = []
            last_valor_centers: List[float] = []

            for page_index, page in enumerate(pdf.pages):
                words = page.extract_words(use_text_flow=True, keep_blank_chars=False)
                if not words:
                    continue

                comp_centers, valor_centers = self._extract_column_centers(words)
                if comp_centers:
                    last_comp_centers = list(comp_centers)
                else:
                    comp_centers = list(last_comp_centers)

                if valor_centers:
                    last_valor_centers = list(valor_centers)
                else:
                    valor_centers = list(last_valor_centers)
                extracted_blocks = self._extract_month_blocks(
                    words, page.height, comp_centers, valor_centers
                )

                active_blocks: List[Tuple[MonthBlock, ActiveBlock]] = []

                next_block_start = (
                    min(block.y_start for block in extracted_blocks)
                    if extracted_blocks
                    else page.height
                )

                for active in pending_blocks:
                    carry_block = MonthBlock(
                        year=active.block.year,
                        months=list(active.block.months),
                        y_start=0,
                        y_end=max(0, min(next_block_start - 0.5, page.height)),
                    )
                    active_blocks.append((carry_block, active))

                for block in extracted_blocks:
                    active_blocks.append(
                        (
                            block,
                            ActiveBlock(
                                block=MonthBlock(
                                    year=block.year,
                                    months=list(block.months),
                                    y_start=block.y_start,
                                    y_end=block.y_end,
                                )
                            ),
                        )
                    )

                next_pending: List[ActiveBlock] = []

                for block, state in active_blocks:
                    block_has_values = False

                    for code, cfg in self.TARGET_CODES.items():
                        column_index = int(cfg["column"])
                        search_prefix = str(cfg.get("search_prefix", code))
                        occurrences = self._find_row_occurrences(
                            words, search_prefix, block
                        )
                        if not occurrences:
                            continue

                        for row_words in occurrences:
                            extracted = self._extract_values_from_row(
                                row_words, block, column_index
                            )
                            if not extracted:
                                continue

                            block_has_values = True

                            target = values.setdefault(code, {})
                            for month_key, amount in extracted.items():
                                existing = target.get(month_key)
                                if existing is not None and existing != amount:
                                    self._log(
                                        f"⚠️ Valores conflitantes para {code} em {month_key[1]:02d}/{month_key[0]} (substituindo {existing} por {amount})."
                                    )
                                target[month_key] = amount

                    if not block_has_values:
                        next_count = state.carry_count + 1
                        if next_count <= self.MAX_BLOCK_CARRY:
                            next_pending.append(
                                ActiveBlock(block=state.block, carry_count=next_count)
                            )
                        else:
                            months_label = ", ".join(
                                f"{month.month:02d}" for month in state.block.months
                            )
                            self._log(
                                f"⚠️ Cabeçalho {state.block.year} ({months_label}) não apresentou valores após {self.MAX_BLOCK_CARRY} páginas."
                            )

                pending_blocks = next_pending

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
    def _collect_values_for_code(
        self,
        aggregated: Dict[str, NumberByMonth],
        code: str,
        months: Iterable[Tuple[int, int]],
        log_label: str,
    ) -> List[Tuple[int, int, Decimal]]:
        results: List[Tuple[int, int, Decimal]] = []

        for year, month in months:
            value = aggregated.get(code, {}).get((year, month), Decimal("0"))
            results.append((year, month, value))
            self._log(
                f"🧮 Mês {month:02d}/{year}: {self._format_decimal(value)} extraído para {log_label}."
            )

        return results

    def _apply_vacation_adjustments(
        self, aggregated: Dict[str, NumberByMonth]
    ) -> None:
        base_values = aggregated.setdefault("3123-Base", {})

        vacation_codes = ("173-Ferias", "174-Ferias")
        vacation_months: Set[Tuple[int, int]] = set()

        for code in vacation_codes:
            for month_key, value in aggregated.get(code, {}).items():
                if value != Decimal("0"):
                    vacation_months.add(month_key)

        if not vacation_months:
            return

        inss_comp = aggregated.get("527-INSS-Comp", {})
        inss_valor = aggregated.get("527-INSS-Valor", {})

        for year, month in vacation_months:
            key = (year, month)
            comp_value = inss_comp.get(key)
            valor_value = inss_valor.get(key)

            if (
                comp_value is None
                or valor_value is None
                or comp_value == Decimal("0")
            ):
                continue

            divisor = comp_value / Decimal("100")
            if divisor == Decimal("0"):
                continue

            additional = valor_value / divisor
            base_current = base_values.get(key, Decimal("0"))
            base_values[key] = base_current + additional

            self._log(
                "➕ Ajuste de férias aplicado em "
                f"{month:02d}/{year}: {self._format_decimal(additional)} somado ao 3123."
            )

    def _write_output_csv(
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

