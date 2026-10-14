# -*- coding: utf-8 -*-
"""Сотрудник — пользователь программы (раздел 4 ТЗ): выбирает себя один раз
при первом запуске, дальше п. 1 заполняется сам. Если отмечена галочка
«Запомнить меня на этом компьютере», выбор сохраняется в
dannye/nastrojki/polzovatel.json и программа больше не спрашивает; без
галочки — файл не создаётся (или удаляется), и при следующем запуске
программа спрашивает сотрудника заново.
"""
from __future__ import annotations

import json
from pathlib import Path


def put_nastrojek(papka_dannyh: str | Path) -> Path:
    return Path(papka_dannyh) / "nastrojki" / "polzovatel.json"


def zagruzit_zapomnennogo_polzovatelya(papka_dannyh: str | Path) -> str | None:
    """ФИО сотрудника, отмеченного галочкой «Запомнить меня», или None —
    тогда программа должна спросить сотрудника заново."""
    put = put_nastrojek(papka_dannyh)
    if not put.exists():
        return None
    dannye = json.loads(put.read_text(encoding="utf-8"))
    return dannye.get("fio")


def zapomnit_polzovatelya(papka_dannyh: str | Path, fio: str) -> None:
    """Сохраняет выбор сотрудника — вызывается, когда галочка отмечена."""
    put = put_nastrojek(papka_dannyh)
    put.parent.mkdir(parents=True, exist_ok=True)
    put.write_text(json.dumps({"fio": fio}, ensure_ascii=False, indent=2), encoding="utf-8")


def _put_poslednih_papok(papka_dannyh: str | Path) -> Path:
    return Path(papka_dannyh) / "nastrojki" / "poslednie_papki.json"


def poslednyaya_papka(papka_dannyh: str | Path, fio: str | None) -> str | None:
    """Папка, выбранная этим сотрудником в прошлый раз (окна выбора
    открываются в ней), или None."""
    put = _put_poslednih_papok(papka_dannyh)
    try:
        papka = json.loads(put.read_text(encoding="utf-8")).get(fio or "")
    except (OSError, ValueError):
        return None
    return papka if papka and Path(papka).is_dir() else None


def zapomnit_poslednyuyu_papku(papka_dannyh: str | Path, fio: str | None, papka: str | Path) -> None:
    put = _put_poslednih_papok(papka_dannyh)
    try:
        dannye = json.loads(put.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        dannye = {}
    dannye[fio or ""] = str(papka)
    put.parent.mkdir(parents=True, exist_ok=True)
    put.write_text(json.dumps(dannye, ensure_ascii=False, indent=2), encoding="utf-8")


def zabyt_polzovatelya(papka_dannyh: str | Path) -> None:
    """Убирает сохранённый выбор — вызывается, когда сотрудник снимает галочку
    или нажимает «Сменить сотрудника»."""
    put = put_nastrojek(papka_dannyh)
    if put.exists():
        put.unlink()
