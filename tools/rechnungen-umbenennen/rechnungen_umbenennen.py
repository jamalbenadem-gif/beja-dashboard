#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
Rechnungen automatisch umbenennen  (Windows / Python)
=====================================================

Liest gescannte Eingangsrechnungen (PDF oder Bild) per Anthropic Claude (Vision)
aus und benennt sie nach folgendem Schema um:

    JJJJ-MM-TT_Firma_Rechnungsnummer_Betrag.ext
    z. B.  2025-03-14_Musterbau GmbH_RE-2025-0421_1190,00EUR.pdf

Es werden ALLE Dateien im angegebenen Ordner und in allen Unterordnern
verarbeitet. Die Dateien bleiben in ihrem jeweiligen Ordner liegen – es ändert
sich nur der Dateiname.

------------------------------------------------------------------------------
EINRICHTUNG (einmalig, unter Windows)
------------------------------------------------------------------------------
1) Python 3.10+ installieren:  https://www.python.org/downloads/windows/
   -> beim Setup unbedingt "Add python.exe to PATH" anhaken.
2) Pakete installieren (in PowerShell oder Eingabeaufforderung):
       pip install -U anthropic pillow
3) API-Schlüssel auf  https://console.anthropic.com  erstellen und setzen:
       setx ANTHROPIC_API_KEY "sk-ant-..."
   Danach das Konsolenfenster EINMAL neu öffnen, damit der Schlüssel aktiv ist.

------------------------------------------------------------------------------
BENUTZUNG
------------------------------------------------------------------------------
Schritt 1 – Vorschaulauf (es wird NICHTS umbenannt, nur eine Tabelle erzeugt):

    python rechnungen_umbenennen.py "C:\Pfad\zu\2025 Rechnungseingang"

Danach die erzeugte Datei  umbenennungen_vorschau.csv  in Excel öffnen und
prüfen. Stimmt alles, folgt:

Schritt 2 – echtes Umbenennen:

    python rechnungen_umbenennen.py "C:\Pfad\zu\2025 Rechnungseingang" --apply

Rückgängig machen (mit der bei --apply erzeugten Protokolldatei):

    python rechnungen_umbenennen.py --undo umbenennungen_log.csv

Nützliche Optionen:
    --limit 5            nur die ersten 5 Dateien (günstiges Testen)
    --model NAME         anderes Modell (Standard: claude-opus-4-7).
                         Für viele Rechnungen ist  claude-haiku-4-5  deutlich
                         günstiger:   --model claude-haiku-4-5
    --max-edge 2200      Bilder vor dem Senden auf max. 2200 px verkleinern
                         (spart Kosten; 0 = nicht verkleinern)
    --keine-pruefung     auch bereits umbenannte Dateien erneut verarbeiten

------------------------------------------------------------------------------
WICHTIG (GoBD / DSGVO)
------------------------------------------------------------------------------
* Lege vor dem ersten echten Lauf eine Sicherungskopie der Originale an.
* Reines Umbenennen ersetzt KEINE GoBD-konforme Archivierung.
* Die Rechnungen werden zur Texterkennung an die Anthropic-API gesendet.
"""

from __future__ import annotations

import argparse
import base64
import csv
import io
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional, Literal

# --- stdout auf UTF-8 stellen (für Umlaute in der Windows-Konsole) ------------
try:
    sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
except Exception:
    pass

# --- optionale Abhängigkeit Pillow (für TIFF/BMP und Verkleinern) -------------
try:
    from PIL import Image  # type: ignore
    _PIL = True
except Exception:
    _PIL = False


DEFAULT_MODEL = "claude-opus-4-7"

# Von der Vision-API direkt unterstützte Bildformate
NATIVE_IMG = {
    "jpg": "image/jpeg", "jpeg": "image/jpeg",
    "png": "image/png", "gif": "image/gif", "webp": "image/webp",
}
# Formate, die erst per Pillow nach PNG gewandelt werden müssen
CONVERT_IMG = {"tif", "tiff", "bmp"}
ALL_EXT = set(NATIVE_IMG) | CONVERT_IMG | {"pdf"}

ILLEGAL = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
ALREADY = re.compile(r"^\d{4}-\d{2}-\d{2}_")  # erkennt bereits umbenannte Dateien

PREVIEW_CSV = "umbenennungen_vorschau.csv"
LOG_CSV = "umbenennungen_log.csv"


# ============================================================================
# Datenmodell für die strukturierte Ausgabe von Claude
# ============================================================================
try:
    from pydantic import BaseModel, Field
except Exception:
    print("FEHLER: Paket 'pydantic' fehlt. Bitte ausführen:  pip install -U anthropic")
    sys.exit(1)


class Rechnung(BaseModel):
    rechnungsdatum: Optional[str] = Field(
        default=None,
        description="Das Rechnungsdatum (NICHT Liefer-/Fälligkeitsdatum) im Format JJJJ-MM-TT. null, wenn nicht eindeutig erkennbar.",
    )
    firma: Optional[str] = Field(
        default=None,
        description="Name des rechnungsstellenden Unternehmens (Lieferant/Absender, meist im Briefkopf) – NICHT der Empfänger. null, wenn nicht erkennbar.",
    )
    rechnungsnummer: Optional[str] = Field(
        default=None,
        description="Rechnungsnummer (Rechnungs-Nr., RG-Nr.) oder ersatzweise Belegnummer. null, wenn keine vorhanden.",
    )
    betrag_brutto: Optional[float] = Field(
        default=None,
        description="Brutto-Gesamtbetrag (Zahlbetrag inkl. MwSt.) als Zahl mit Punkt als Dezimaltrennzeichen, OHNE Tausenderpunkte. null, wenn nicht erkennbar.",
    )
    waehrung: Optional[str] = Field(
        default=None,
        description="Währung als ISO-Code, z. B. EUR. Bei € immer EUR.",
    )
    konfidenz: Literal["hoch", "mittel", "niedrig"] = Field(
        description="Wie sicher die Erkennung insgesamt ist.",
    )


SYSTEM_PROMPT = (
    "Du bist ein präziser Assistent für die Buchhaltung eines deutschen "
    "Unternehmens und liest eingescannte EINGANGSRECHNUNGEN aus.\n\n"
    "Extrahiere ausschließlich die folgenden Felder und halte dich strikt an "
    "diese Regeln:\n\n"
    "1) rechnungsdatum: Das Datum der Rechnung (Beschriftung wie 'Rechnungsdatum', "
    "'Belegdatum', 'Datum'). NICHT das Lieferdatum/Leistungsdatum und NICHT das "
    "Fälligkeitsdatum. Gib es immer im Format JJJJ-MM-TT zurück (z. B. 14.03.2025 "
    "-> 2025-03-14).\n\n"
    "2) firma: Der Name des Unternehmens, das die Rechnung AUSGESTELLT hat "
    "(Rechnungssteller / Lieferant / Absender). Das steht meist im Briefkopf oben "
    "oder im Fußbereich (Bankverbindung, USt-IdNr.). Gib NICHT den Rechnungs"
    "empfänger zurück. Verwende einen kurzen, eindeutigen Firmennamen ohne "
    "Rechtsformzusätze wegzulassen (GmbH, AG, e.K. dürfen bleiben).\n\n"
    "3) rechnungsnummer: Die Rechnungsnummer ('Rechnungs-Nr.', 'Rechnung Nr.', "
    "'RG-Nr.', 'Invoice No.'). Falls keine vorhanden ist, ersatzweise eine "
    "Belegnummer. Kundennummer oder Bestellnummer sind NICHT die Rechnungsnummer.\n\n"
    "4) betrag_brutto: Der Brutto-Gesamtbetrag, also der tatsächlich zu zahlende "
    "Betrag inklusive Mehrwertsteuer ('Gesamtbetrag', 'Rechnungsbetrag', "
    "'Zahlbetrag', 'Brutto', 'zu zahlen'). Als Zahl mit Punkt als Dezimal"
    "trennzeichen und OHNE Tausenderpunkte (z. B. 1.190,00 EUR -> 1190.00). "
    "Wenn nur ein Nettobetrag erkennbar ist, gib dennoch den zu zahlenden "
    "Endbetrag zurück.\n\n"
    "5) waehrung: ISO-Code der Währung. Bei '€' oder fehlender Angabe: EUR.\n\n"
    "6) konfidenz: 'hoch', 'mittel' oder 'niedrig' – je nach Lesbarkeit des Scans "
    "und Eindeutigkeit der Felder.\n\n"
    "WICHTIG: Wenn ein Feld nicht zweifelsfrei erkennbar ist, gib dafür null "
    "zurück. RATE NICHT und erfinde keine Werte. Lieber null als ein falscher "
    "Wert – falsch benannte Belege verursachen in der Buchhaltung Fehler."
)

USER_PROMPT = (
    "Hier ist eine eingescannte Eingangsrechnung. Lies die geforderten Felder "
    "gemäß den Regeln aus. Gib für nicht eindeutig erkennbare Felder null zurück."
)


# ============================================================================
# Hilfsfunktionen
# ============================================================================
def clean_part(text: Optional[str], maxlen: int = 80) -> str:
    """Macht einen String windows-dateinamentauglich."""
    if not text:
        return ""
    s = ILLEGAL.sub("", text)
    s = re.sub(r"\s+", " ", s).strip()
    s = s.strip(" .")  # Windows: kein abschließender Punkt / Leerzeichen
    return s[:maxlen].strip(" .")


def norm_datum(d: Optional[str]) -> Optional[str]:
    """Normalisiert ein Datum auf JJJJ-MM-TT, sonst None."""
    if not d:
        return None
    d = d.strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y", "%Y/%m/%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(d, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def fmt_betrag(betrag: Optional[float], waehrung: Optional[str]) -> Optional[str]:
    """Formatiert den Betrag als '1190,00EUR'."""
    if betrag is None:
        return None
    cur = re.sub(r"[^A-Za-z]", "", (waehrung or "EUR")).upper()[:3] or "EUR"
    return f"{betrag:.2f}".replace(".", ",") + cur


def build_content_block(path: Path, max_edge: int) -> dict:
    """Erzeugt den Vision-Content-Block (PDF als document, Bild als image)."""
    ext = path.suffix.lower().lstrip(".")

    if ext == "pdf":
        data = base64.standard_b64encode(path.read_bytes()).decode("ascii")
        return {
            "type": "document",
            "source": {"type": "base64", "media_type": "application/pdf", "data": data},
        }

    if ext in NATIVE_IMG or ext in CONVERT_IMG:
        if _PIL:
            img = Image.open(path)
            if img.mode != "RGB":
                img = img.convert("RGB")
            if max_edge and max(img.size) > max_edge:
                img.thumbnail((max_edge, max_edge))
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            data = base64.standard_b64encode(buf.getvalue()).decode("ascii")
            return {
                "type": "image",
                "source": {"type": "base64", "media_type": "image/png", "data": data},
            }
        if ext in NATIVE_IMG:
            data = base64.standard_b64encode(path.read_bytes()).decode("ascii")
            return {
                "type": "image",
                "source": {"type": "base64", "media_type": NATIVE_IMG[ext], "data": data},
            }
        raise RuntimeError(f"Dateityp .{ext} benötigt das Paket 'pillow' (pip install pillow).")

    raise RuntimeError(f"Nicht unterstützter Dateityp: .{ext}")


def unique_target(target: Path) -> Path:
    """Hängt -2, -3 ... an, falls der Zielname schon existiert."""
    if not target.exists():
        return target
    stem, suffix, parent = target.stem, target.suffix, target.parent
    i = 2
    while True:
        cand = parent / f"{stem}-{i}{suffix}"
        if not cand.exists():
            return cand
        i += 1


# ============================================================================
# Kernverarbeitung
# ============================================================================
def extrahiere(client, model: str, path: Path, max_edge: int):
    """Ruft Claude auf und gibt (Rechnung|None, usage|None) zurück."""
    block = build_content_block(path, max_edge)
    system = [{"type": "text", "text": SYSTEM_PROMPT, "cache_control": {"type": "ephemeral"}}]
    resp = client.messages.parse(
        model=model,
        max_tokens=2048,
        system=system,
        messages=[{"role": "user", "content": [block, {"type": "text", "text": USER_PROMPT}]}],
        output_format=Rechnung,
    )
    return resp.parsed_output, getattr(resp, "usage", None)


def plane_namen(inv: Optional[Rechnung], original: Path):
    """Bestimmt aus den extrahierten Daten den neuen Namen + Status."""
    datum = norm_datum(inv.rechnungsdatum) if inv else None
    firma_raw = clean_part(inv.firma) if inv else ""
    nummer_raw = clean_part(inv.rechnungsnummer, maxlen=40) if inv else ""
    betrag_str = fmt_betrag(inv.betrag_brutto, inv.waehrung) if inv else None

    fehlend = []
    if not datum:
        fehlend.append("Datum")
    if not betrag_str:
        fehlend.append("Betrag")
    if not firma_raw and not nummer_raw:
        fehlend.append("Firma/Nummer")

    if fehlend:
        return None, "PRUEFEN (fehlt: " + ", ".join(fehlend) + ")"

    firma = firma_raw or "Firma-unbekannt"
    nummer = nummer_raw or "ohne-Nr"
    neu = f"{datum}_{firma}_{nummer}_{betrag_str}{original.suffix.lower()}"
    neu = clean_part(neu, maxlen=180) or neu
    # Endung sicher anhängen, falls clean_part sie gekürzt hat
    if not neu.lower().endswith(original.suffix.lower()):
        neu += original.suffix.lower()
    return neu, "OK"


def main():
    ap = argparse.ArgumentParser(
        description="Benennt gescannte Rechnungen per Claude (Vision) nach Datum/Firma/Nummer/Betrag um.",
    )
    ap.add_argument("ordner", nargs="?", help=r'Ordner mit den Scans, z. B. "C:\...\2025 Rechnungseingang"')
    ap.add_argument("--apply", action="store_true", help="Dateien wirklich umbenennen (sonst nur Vorschau).")
    ap.add_argument("--model", default=DEFAULT_MODEL, help=f"Claude-Modell (Standard: {DEFAULT_MODEL}).")
    ap.add_argument("--limit", type=int, default=0, help="Nur die ersten N Dateien verarbeiten (0 = alle).")
    ap.add_argument("--max-edge", type=int, default=2200, help="Bilder auf max. N Pixel verkleinern (0 = nicht).")
    ap.add_argument("--keine-pruefung", action="store_true", help="Auch bereits umbenannte Dateien verarbeiten.")
    ap.add_argument("--api-key", default=None, help="API-Schlüssel direkt übergeben (sonst Umgebungsvariable).")
    ap.add_argument("--undo", metavar="LOGDATEI", help="Umbenennungen aus einer Protokolldatei rückgängig machen.")
    args = ap.parse_args()

    if args.undo:
        return undo(Path(args.undo))

    if not args.ordner:
        ap.error('Bitte den Ordner angeben, z. B.:  python rechnungen_umbenennen.py "C:\\...\\2025 Rechnungseingang"')

    root = Path(args.ordner)
    if not root.is_dir():
        print(f"FEHLER: Ordner nicht gefunden: {root}")
        sys.exit(1)

    try:
        import anthropic
    except Exception:
        print("FEHLER: Paket 'anthropic' fehlt. Bitte ausführen:  pip install -U anthropic")
        sys.exit(1)

    try:
        client = anthropic.Anthropic(api_key=args.api_key) if args.api_key else anthropic.Anthropic()
    except Exception as e:
        print(f"FEHLER beim Initialisieren der API: {e}")
        print("Ist ANTHROPIC_API_KEY gesetzt? (setx ANTHROPIC_API_KEY \"sk-ant-...\", Fenster neu öffnen)")
        sys.exit(1)

    files = sorted(p for p in root.rglob("*")
                   if p.is_file() and p.suffix.lower().lstrip(".") in ALL_EXT)
    if not args.keine_pruefung:
        files = [p for p in files if not ALREADY.match(p.name)]
    if args.limit:
        files = files[: args.limit]

    if not files:
        print("Keine (noch nicht umbenannten) Rechnungen gefunden.")
        return

    print(f"{len(files)} Datei(en) werden verarbeitet  |  Modell: {args.model}  |  "
          f"Modus: {'UMBENENNEN' if args.apply else 'NUR VORSCHAU'}")
    if not _PIL:
        print("Hinweis: 'pillow' ist nicht installiert – TIFF/BMP werden übersprungen, "
              "Bilder werden nicht verkleinert.  (pip install pillow)")

    rows = []
    undo_rows = []
    zaehler = {"OK": 0, "PRUEFEN": 0, "FEHLER": 0, "UMBENANNT": 0}

    for i, path in enumerate(files, 1):
        rel = path.relative_to(root)
        try:
            inv, usage = extrahiere(client, args.model, path, args.max_edge)
            neu, status = plane_namen(inv, path)
            if i == 1 and usage is not None:
                cr = getattr(usage, "cache_read_input_tokens", 0)
                print(f"   (Tokens 1. Aufruf: input={getattr(usage,'input_tokens','?')}, cache_read={cr})")
        except Exception as e:
            zaehler["FEHLER"] += 1
            rows.append({"ordner": str(rel.parent), "alt_name": path.name, "neu_name": "",
                         "datum": "", "firma": "", "nummer": "", "betrag": "",
                         "waehrung": "", "konfidenz": "", "status": "FEHLER", "fehler": str(e)})
            print(f"[{i}/{len(files)}] FEHLER  {rel}  -> {e}")
            continue

        row = {
            "ordner": str(rel.parent),
            "alt_name": path.name,
            "neu_name": neu or "",
            "datum": (inv.rechnungsdatum or "") if inv else "",
            "firma": (inv.firma or "") if inv else "",
            "nummer": (inv.rechnungsnummer or "") if inv else "",
            "betrag": ("" if not inv or inv.betrag_brutto is None else f"{inv.betrag_brutto:.2f}"),
            "waehrung": (inv.waehrung or "") if inv else "",
            "konfidenz": (inv.konfidenz or "") if inv else "",
            "status": status,
            "fehler": "",
        }

        if status == "OK" and neu:
            zaehler["OK"] += 1
            if path.name == neu:
                print(f"[{i}/{len(files)}] schon korrekt  {rel}")
            elif args.apply:
                target = unique_target(path.with_name(neu))
                path.rename(target)
                undo_rows.append({"neu_pfad": str(target.resolve()), "alt_pfad": str(path.resolve())})
                zaehler["UMBENANNT"] += 1
                row["neu_name"] = target.name
                print(f"[{i}/{len(files)}] umbenannt  {rel}  ->  {target.name}")
            else:
                print(f"[{i}/{len(files)}] Vorschau   {rel}  ->  {neu}")
        else:
            zaehler["PRUEFEN"] += 1
            print(f"[{i}/{len(files)}] {status}  {rel}")

        rows.append(row)

    # Vorschau-/Ergebnis-Tabelle schreiben (UTF-8 mit BOM -> Excel-freundlich)
    out_csv = root / PREVIEW_CSV
    with out_csv.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["ordner", "alt_name", "neu_name", "datum", "firma",
                                          "nummer", "betrag", "waehrung", "konfidenz", "status", "fehler"])
        w.writeheader()
        w.writerows(rows)

    print("\n----------------------------------------------------------------")
    print(f"Fertig.  OK: {zaehler['OK']}   Prüfen: {zaehler['PRUEFEN']}   Fehler: {zaehler['FEHLER']}")
    print(f"Tabelle gespeichert: {out_csv}")

    if args.apply and undo_rows:
        log = root / LOG_CSV
        with log.open("w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=["neu_pfad", "alt_pfad"])
            w.writeheader()
            w.writerows(undo_rows)
        print(f"{zaehler['UMBENANNT']} Datei(en) umbenannt.  Rückgängig mit:")
        print(f'   python rechnungen_umbenennen.py --undo "{log}"')
    elif not args.apply:
        print("Das war nur die Vorschau. Zum echten Umbenennen denselben Befehl mit  --apply  ausführen.")


def undo(logfile: Path):
    """Macht Umbenennungen aus der Protokolldatei rückgängig."""
    if not logfile.is_file():
        print(f"FEHLER: Protokolldatei nicht gefunden: {logfile}")
        sys.exit(1)
    n = 0
    with logfile.open(newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            neu, alt = Path(r["neu_pfad"]), Path(r["alt_pfad"])
            if neu.exists() and not alt.exists():
                neu.rename(alt)
                n += 1
                print(f"zurück: {neu.name}  ->  {alt.name}")
            else:
                print(f"übersprungen: {neu.name} (Quelle fehlt oder Ziel existiert bereits)")
    print(f"\nFertig. {n} Datei(en) zurückbenannt.")


if __name__ == "__main__":
    main()
