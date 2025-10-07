"""Revision checker logic for comparing revision histories across inputs."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Dict, Iterable, List, Optional, Pattern, Tuple

import pandas as pd

DATE_FUZZY_REGEX = re.compile(
    r"(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4})|"
    r"(?:\d{1,2}\s*[-]?\s*[A-Za-z]{3,9}\s*[-]?\s*\d{2,4})",
    re.IGNORECASE,
)


def _append_comment(existing: str, new_comment: str) -> str:
    new_comment = (new_comment or "").strip()
    if not new_comment:
        return existing
    if existing:
        return f"{existing}\n{new_comment}"
    return new_comment


@dataclass
class HighlightState:
    rev: bool = False
    desc: bool = False
    date: bool = False
    original_text: str = ""


@dataclass
class RevEntry:
    rev: Optional[str]
    desc: Optional[str]
    date: Optional[date]
    raw_rev: str = ""
    raw_desc: str = ""
    raw_date: str = ""
    column: Optional[str] = None
    source: str = ""
    generated: bool = False


@dataclass
class PatternRule:
    name: str
    regex: Pattern[str]
    kind: str  # "non-incremental" or "incremental"
    fixed_revision: Optional[str] = None
    prefix: str = ""
    base: int = 10
    pad: int = 0
    core_group: int = 1
    start_value: int = 0
    step: int = 1

    def matches(self, value: str) -> bool:
        if value is None:
            return False
        value = str(value).strip()
        if not value:
            return False
        return bool(self.regex.fullmatch(value))

    def value_of(self, value: str) -> Optional[int]:
        if self.kind != "incremental":
            return None
        if value is None:
            return None
        match = self.regex.fullmatch(str(value).strip())
        if not match:
            return None
        token = match.group(self.core_group)
        if self.base == 10:
            try:
                return int(token)
            except ValueError:
                return None
        if self.base == 26:
            total = 0
            token = token.upper()
            for ch in token:
                if not ("A" <= ch <= "Z"):
                    return None
                total = total * 26 + (ord(ch) - ord("A"))
            return total
        return None

    def format_value(self, counter: int) -> str:
        if self.kind != "incremental":
            if self.fixed_revision is None:
                raise ValueError("Fixed revision not provided for non-incremental pattern")
            return self.fixed_revision
        if self.base == 10:
            core = f"{counter}".zfill(self.pad) if self.pad else str(counter)
        elif self.base == 26:
            if counter < 0:
                raise ValueError("Counter must be non-negative for alphabetical increments")
            digits: List[str] = []
            value = counter
            while True:
                digits.append(chr(ord("A") + (value % 26)))
                value //= 26
                if value <= 0:
                    break
            core = "".join(reversed(digits))
        else:
            raise ValueError("Unsupported base for pattern rule")
        return f"{self.prefix}{core}"

    def next_after(self, value: str) -> Optional[str]:
        if self.kind != "incremental":
            return None
        current = self.value_of(value)
        if current is None:
            return None
        return self.format_value(current + self.step)


@dataclass
class CustomPatternConfig:
    prefix: str = ""
    core_regex: str = r"(\d+)"
    padding: int = 0
    base: int = 10
    start: str = "0"
    step: int = 1


@dataclass
class RevCheckSettings:
    pattern_mode: str = "incremental"
    pattern_choice: str = "P0x"
    fixed_revision: Optional[str] = None
    custom_pattern: Optional[CustomPatternConfig] = None
    latest_desc_enabled: bool = False
    latest_desc_value: Optional[str] = None
    date_enabled: bool = False
    date_strict: bool = False
    date_format: Optional[str] = None
    latest_date_value: Optional[str] = None
    input1_rev_cols: List[str] = field(default_factory=list)
    input2_block_cols: List[str] = field(default_factory=list)
    generate_latest_for_input2: bool = True


def build_pattern_rule(settings: RevCheckSettings) -> PatternRule:
    if settings.pattern_mode == "non-incremental":
        fixed = settings.fixed_revision or ""
        return PatternRule(
            name="Fixed",
            regex=re.compile(re.escape(fixed)),
            kind="non-incremental",
            fixed_revision=fixed,
        )

    choice = (settings.pattern_choice or "").lower()
    if choice == "xx":
        return PatternRule(
            name="XX",
            regex=re.compile(r"^(\d{2})$"),
            kind="incremental",
            prefix="",
            base=10,
            pad=2,
            start_value=0,
            step=1,
        )
    if choice in {"alphabet only", "alphabet"}:
        return PatternRule(
            name="Alphabet",
            regex=re.compile(r"^([A-Z]+)$"),
            kind="incremental",
            prefix="",
            base=26,
            pad=0,
            start_value=0,
            step=1,
        )
    if choice in {"ifc (dae)", "ifc"}:
        return PatternRule(
            name="IFC (DAE)",
            regex=re.compile(r"^C(\d{2})$"),
            kind="incremental",
            prefix="C",
            base=10,
            pad=2,
            start_value=0,
            step=1,
        )
    if choice in {"p0x", "p0"}:
        return PatternRule(
            name="P0x",
            regex=re.compile(r"^P(\d{2})$"),
            kind="incremental",
            prefix="P",
            base=10,
            pad=2,
            start_value=0,
            step=1,
        )

    custom = settings.custom_pattern or CustomPatternConfig()
    core_regex = custom.core_regex or r"(\d+)"
    if "(" not in core_regex:
        core_regex = f"({core_regex})"
    pattern = re.compile(f"^{re.escape(custom.prefix)}{core_regex}$")

    if custom.base == 26:
        start_value = 0
        token = (custom.start or "A").upper()
        for ch in token:
            start_value = start_value * 26 + (ord(ch) - ord("A"))
    else:
        try:
            start_value = int(custom.start or 0)
        except ValueError:
            start_value = 0

    return PatternRule(
        name="Custom",
        regex=pattern,
        kind="incremental",
        prefix=custom.prefix,
        base=custom.base,
        pad=custom.padding,
        start_value=start_value,
        step=custom.step or 1,
    )


def _strict_formats() -> Dict[str, str]:
    return {
        "DD/MM/YY": "%d/%m/%y",
        "DD/MM/YYYY": "%d/%m/%Y",
        "DD-MMM-YYYY": "%d-%b-%Y",
        "YYYY-MM-DD": "%Y-%m-%d",
    }


def _normalize_date(
    value: Optional[str],
    settings: RevCheckSettings,
    *,
    for_header: bool = False,
) -> Tuple[Optional[date], Optional[str]]:
    if value is None:
        return None, None
    text = str(value).strip()
    if not text:
        return None, None

    if not settings.date_enabled and not for_header:
        return None, None

    fmt_error: Optional[str] = None

    if settings.date_enabled and settings.date_strict:
        fmt = settings.date_format or ""
        fmt_lookup = _strict_formats()
        fmt = fmt_lookup.get(fmt, fmt)
        try:
            parsed = datetime.strptime(text, fmt)
            return parsed.date(), None
        except Exception as exc:  # noqa: BLE001
            fmt_error = f"Invalid date '{text}': {exc}"
            return None, fmt_error

    match_value = text
    if settings.date_enabled and not settings.date_strict:
        matches = DATE_FUZZY_REGEX.findall(text)
        if matches:
            match_value = matches[-1]
    try:
        parsed = pd.to_datetime(match_value, dayfirst=True, errors="raise")
        if pd.isna(parsed):
            raise ValueError("Parsed NaT")
        return parsed.date(), None
    except Exception as exc:  # noqa: BLE001
        fmt_error = f"Unable to parse date '{text}': {exc}"
        return None, fmt_error


def _build_highlight_segments(state: HighlightState) -> List[Tuple[str, str]]:
    text = state.original_text or ""
    tokens = re.split(r"(\|)", text)
    if len(tokens) <= 1:
        parts = [state.original_text]
    else:
        parts = tokens

    segments: List[Tuple[str, str]] = []
    component_index = 0
    for token in parts:
        if token == "|":
            segments.append((token, "000000"))
            continue
        highlight = False
        if component_index == 0:
            highlight = state.rev
        elif component_index == 1:
            highlight = state.desc
        elif component_index == 2:
            highlight = state.date
        color = "FF0000" if highlight else "000000"
        segments.append((token, color))
        component_index += 1
    return segments


def _parse_input1_entries(
    row: pd.Series,
    settings: RevCheckSettings,
    pattern_rule: PatternRule,
) -> Tuple[List[RevEntry], Dict[str, HighlightState], List[str]]:
    entries: List[RevEntry] = []
    highlight_states: Dict[str, HighlightState] = {}
    errors: List[str] = []

    for col in settings.input1_rev_cols:
        raw_value = row.get(col, "")
        text = "" if pd.isna(raw_value) else str(raw_value)
        parts = [p.strip(" ,") for p in text.split("|")]
        while len(parts) < 3:
            parts.append("")
        rev_raw, desc_raw, date_raw = parts[:3]

        normalized_date, date_error = _normalize_date(date_raw, settings)
        if date_error:
            errors.append(f"invalid Date format in {col}")
            state = highlight_states.get(col)
            if state:
                state.date = True
            else:
                highlight_states[col] = HighlightState(date=True)

        entry = RevEntry(
            rev=rev_raw or None,
            desc=desc_raw or None,
            date=normalized_date,
            raw_rev=rev_raw,
            raw_desc=desc_raw,
            raw_date=date_raw,
            column=col,
            source="input1",
        )
        entries.append(entry)

        highlight_states[col] = HighlightState(original_text=text or "")

        if pattern_rule.kind == "incremental":
            previous_value: Optional[int] = None
            for entry in entries:
                if not entry.rev:
                    continue
                value = pattern_rule.value_of(entry.rev)
                if value is None:
                    continue
                if previous_value is not None:
                    expected = previous_value + pattern_rule.step
                    if value != expected:
                        errors.append(
                            f"unexpected increment at {entry.column} (expected {pattern_rule.format_value(expected)})"
                        )
                        state = highlight_states.get(entry.column)
                        if state:
                            state.rev = True
                previous_value = value

    return entries, highlight_states, errors


def _parse_input2_entries(
    row: pd.Series,
    settings: RevCheckSettings,
) -> Tuple[List[RevEntry], List[str]]:
    entries: List[RevEntry] = []
    errors: List[str] = []

    for col in settings.input2_block_cols:
        if "|" in col:
            header_parts = [p.strip() for p in col.split("|")]
            desc = header_parts[0]
            header_date_raw = header_parts[1] if len(header_parts) > 1 else ""
        else:
            desc = col.strip()
            header_date_raw = ""

        normalized_date, date_error = _normalize_date(header_date_raw, settings, for_header=True)
        if date_error and settings.date_enabled:
            errors.append(f"invalid reference date in column {col}")

        cell_value = row.get(col, "")
        cell_text = "" if pd.isna(cell_value) else str(cell_value).strip()
        entry = RevEntry(
            rev=cell_text or None,
            desc=desc or None,
            date=normalized_date,
            raw_rev=cell_text,
            raw_desc=desc,
            raw_date=header_date_raw,
            column=col,
            source="input2",
        )
        entries.append(entry)

    return entries, errors


def _generate_latest_entry(
    input2_entries: List[RevEntry],
    settings: RevCheckSettings,
    pattern_rule: PatternRule,
) -> Optional[RevEntry]:
    if not settings.generate_latest_for_input2:
        return None
    if pattern_rule.kind != "incremental":
        return None

    highest: Optional[int] = None
    for entry in input2_entries:
        if not entry.rev:
            continue
        value = pattern_rule.value_of(entry.rev)
        if value is None:
            continue
        highest = value if highest is None else max(highest, value)

    if highest is None:
        highest = pattern_rule.start_value - pattern_rule.step

    next_value = highest + pattern_rule.step
    try:
        next_rev = pattern_rule.format_value(next_value)
    except ValueError:
        return None

    desc = settings.latest_desc_value if settings.latest_desc_enabled else None
    normalized_date, date_error = _normalize_date(settings.latest_date_value, settings)
    if date_error and settings.date_enabled:
        # this error will be surfaced during comparison by returning None date but logging message
        pass

    return RevEntry(
        rev=next_rev,
        desc=desc,
        date=normalized_date,
        raw_rev=next_rev,
        raw_desc=desc or "",
        raw_date=settings.latest_date_value or "",
        column="GeneratedLatest",
        source="generated",
        generated=True,
    )


def _compress_input2(entries: Iterable[RevEntry]) -> List[RevEntry]:
    result: List[RevEntry] = []
    for entry in entries:
        if entry.rev:
            result.append(entry)
    return result


def _record_highlight(
    highlight_states: Dict[str, HighlightState],
    column: Optional[str],
    *,
    rev: bool = False,
    desc: bool = False,
    date: bool = False,
) -> None:
    if not column:
        return
    state = highlight_states.get(column)
    if not state:
        state = HighlightState()
        highlight_states[column] = state
    state.rev = state.rev or rev
    state.desc = state.desc or desc
    state.date = state.date or date


def apply_revision_checks(
    merged_df: pd.DataFrame,
    settings: Optional[RevCheckSettings],
) -> Tuple[pd.DataFrame, Dict[int, Dict[str, List[Tuple[str, str]]]]]:
    if settings is None or not settings.input1_rev_cols or not settings.input2_block_cols:
        if "Comments-Revision" not in merged_df.columns:
            merged_df["Comments-Revision"] = ""
        return merged_df, {}

    pattern_rule = build_pattern_rule(settings)
    comments: List[str] = []
    highlights: Dict[int, Dict[str, List[Tuple[str, str]]]] = {}

    for idx, row in merged_df.iterrows():
        row_comments = ""
        input1_entries, highlight_states, parse_errors = _parse_input1_entries(row, settings, pattern_rule)
        for err in parse_errors:
            row_comments = _append_comment(row_comments, err)

        input2_entries, input2_errors = _parse_input2_entries(row, settings)
        for err in input2_errors:
            row_comments = _append_comment(row_comments, err)

        generated = _generate_latest_entry(input2_entries, settings, pattern_rule)
        if generated:
            input2_entries.append(generated)

        compressed_input2 = _compress_input2(input2_entries)
        if not compressed_input2 and not generated:
            row_comments = _append_comment(row_comments, "no reference revisions in Input 2")

        max_len = max(len(input1_entries), len(compressed_input2))
        for pos in range(max_len):
            entry1 = input1_entries[pos] if pos < len(input1_entries) else None
            entry2 = compressed_input2[pos] if pos < len(compressed_input2) else None

            if entry1 is None:
                if entry2 is not None:
                    row_comments = _append_comment(
                        row_comments,
                        f"missing Rev in position {pos + 1}",
                    )
                continue

            column = entry1.column

            if entry1.rev:
                if not pattern_rule.matches(entry1.rev):
                    row_comments = _append_comment(row_comments, f"invalid Revision tag in {column}")
                    _record_highlight(highlight_states, column, rev=True)
            elif entry2 and entry2.rev:
                row_comments = _append_comment(row_comments, f"incorrect Rev in {column}")
                _record_highlight(highlight_states, column, rev=True, desc=True, date=True)
                continue

            if entry2 is None:
                if entry1.rev or entry1.desc or entry1.raw_date:
                    row_comments = _append_comment(row_comments, f"extra revision in {column}")
                    _record_highlight(highlight_states, column, rev=True, desc=True, date=True)
                continue

            if entry2.rev and not pattern_rule.matches(entry2.rev):
                row_comments = _append_comment(row_comments, f"Input 2 invalid revision at position {pos + 1}")

            if entry1.rev and entry2.rev and entry1.rev != entry2.rev:
                row_comments = _append_comment(row_comments, f"incorrect Rev in {column}")
                _record_highlight(highlight_states, column, rev=True)

            if settings.latest_desc_enabled or not entry2.generated:
                if entry1.desc and entry2.desc and entry1.desc != entry2.desc:
                    row_comments = _append_comment(row_comments, f"incorrect Description in {column}")
                    _record_highlight(highlight_states, column, desc=True)
                elif entry1.desc and entry2.desc is None and not entry2.generated:
                    row_comments = _append_comment(row_comments, f"incorrect Description in {column}")
                    _record_highlight(highlight_states, column, desc=True)

            if settings.date_enabled or not entry2.generated:
                if entry1.date and entry2.date and entry1.date != entry2.date:
                    row_comments = _append_comment(row_comments, f"incorrect Date in {column}")
                    _record_highlight(highlight_states, column, date=True)
                elif entry1.date and entry2.date is None and not entry2.generated:
                    row_comments = _append_comment(row_comments, f"incorrect Date in {column}")
                    _record_highlight(highlight_states, column, date=True)

        # Finalise highlight segments
        for col, state in list(highlight_states.items()):
            if state.rev or state.desc or state.date:
                highlights.setdefault(idx, {})[col] = _build_highlight_segments(state)

        comments.append(row_comments)

    merged_df["Comments-Revision"] = comments
    return merged_df, highlights
