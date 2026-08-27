#!/usr/bin/env python3
"""
Fetch brand colors (Brandfetch) and logos (logo.dev) for a company.

Colors come from Brandfetch (BRANDFETCH_API_KEY env var or .env file).
Logo images come from logo.dev instead — the same source already used
everywhere else in the app (tracker, agent picker, Slack/Teams cards) for a
consistent look, and it sidesteps Brandfetch quirks like a domain whose only
logo asset is a "light" variant meant for a dark background, which
disappears against this app's white document/page backgrounds.

Usage:
    from scripts.brand_color import get_brand_color, get_brand_logo
    colors = get_brand_color("Brightflag")
    # {"primary": "1A3C5E", "secondary": "00695C", "border": "2B6CB0", "fill": "EEF4FB"}
    logo = get_brand_logo("Brightflag")
    # {"bytes": b"...", "format": "png", "width": 512, "height": 512} or None
"""

import os
import re
from pathlib import Path
from urllib.parse import quote

try:
    import requests
except ImportError:
    requests = None  # type: ignore

# Fallback palette — Corey's default navy/teal scheme
DEFAULT_PALETTE = {
    "primary":   "1A3C5E",
    "secondary": "00695C",
    "border":    "2B6CB0",
    "fill":      "EEF4FB",
}

# Brandfetch search client token (autocomplete endpoint). Set BRANDFETCH_SEARCH_CLIENT in env.
_SEARCH_CLIENT = os.environ.get("BRANDFETCH_SEARCH_CLIENT", "")

# Brand colors are spliced into generated JS/DOCX-XML downstream — only accept
# well-formed 6-digit hex so a malformed API response can't break out of those contexts.
_HEX_RE = re.compile(r"^[0-9A-F]{6}$")

_COLOR_PRIORITY = ("accent", "dark", "light")

# Public logo.dev token — same one already hardcoded as a fallback in
# slack_bot.py and the frontend pages that render company logos.
_LOGODEV_PUBLIC_KEY = (
    os.environ.get("LOGODEV_PUBLIC_KEY")
    or os.environ.get("LOGODEV_API_KEY")
    or "pk_U3oIYbhyTvinmftvOvCTJg"
)


def _png_dimensions(data: bytes) -> tuple[int, int] | None:
    """Parse width/height straight from a PNG's IHDR chunk — no imaging library needed."""
    if len(data) < 24 or data[:8] != b"\x89PNG\r\n\x1a\n" or data[12:16] != b"IHDR":
        return None
    return int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big")


def _load_api_key() -> str:
    key = os.environ.get("BRANDFETCH_API_KEY", "")
    if key:
        return key
    env_path = Path(__file__).parent.parent / ".env"
    if env_path.exists():
        for line in env_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line.startswith("BRANDFETCH_API_KEY="):
                return line.split("=", 1)[1].strip()
    raise RuntimeError(
        "BRANDFETCH_API_KEY not found. Set it in the environment or in .env"
    )


def _lighten(hex6: str, factor: float) -> str:
    """Blend hex6 color with white. factor=0 → original, factor=1 → white."""
    r = int(hex6[0:2], 16)
    g = int(hex6[2:4], 16)
    b = int(hex6[4:6], 16)
    r = round(r + (255 - r) * factor)
    g = round(g + (255 - g) * factor)
    b = round(b + (255 - b) * factor)
    return f"{r:02X}{g:02X}{b:02X}"


def _darken(hex6: str, factor: float) -> str:
    """Blend hex6 color with black. factor=0 → original, factor=1 → black."""
    r = int(hex6[0:2], 16)
    g = int(hex6[2:4], 16)
    b = int(hex6[4:6], 16)
    r = round(r * (1 - factor))
    g = round(g * (1 - factor))
    b = round(b * (1 - factor))
    return f"{r:02X}{g:02X}{b:02X}"


# WCAG AA contrast ratio for normal text against a white page background.
_MIN_TEXT_CONTRAST = 4.5


def _relative_luminance(hex6: str) -> float:
    """WCAG relative luminance (0=black, 1=white) for an sRGB hex color."""
    def _channel(c: int) -> float:
        c /= 255
        return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4

    r, g, b = (int(hex6[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * _channel(r) + 0.7152 * _channel(g) + 0.0722 * _channel(b)


def _contrast_on_white(hex6: str) -> float:
    """WCAG contrast ratio of hex6 text against a white (#FFFFFF) background."""
    return (1.0 + 0.05) / (_relative_luminance(hex6) + 0.05)


def _ensure_readable(hex6: str, min_contrast: float = _MIN_TEXT_CONTRAST) -> str:
    """Darken hex6, preserving hue, until it reads clearly as text on a white
    page. Brandfetch sometimes returns an "accent" color meant for dark
    backgrounds or UI chrome — too light to use directly as document text."""
    if _contrast_on_white(hex6) >= min_contrast:
        return hex6

    darkened = hex6
    for step in range(1, 20):
        darkened = _darken(hex6, step * 0.05)
        if _contrast_on_white(darkened) >= min_contrast:
            break

    print(f"  ⚠ Brandfetch: #{hex6} too light for text "
          f"({_contrast_on_white(hex6):.1f}:1) — darkened to #{darkened}")
    return darkened


def _resolve_domain(company_name: str) -> str | None:
    """Look up company_name via Brandfetch's search endpoint. Returns the domain or None."""
    search_url = f"https://api.brandfetch.io/v2/search/{quote(company_name)}"
    search_resp = requests.get(search_url, params={"c": _SEARCH_CLIENT}, timeout=10)
    search_data = search_resp.json()

    if not isinstance(search_data, list) or not search_data:
        print(f"  ⚠ Brandfetch: no results for '{company_name}'")
        return None

    domain = search_data[0].get("domain", "")
    if not domain:
        print("  ⚠ Brandfetch: no domain in search result")
        return None

    print(f"  ✓ Brandfetch: resolved '{company_name}' → {domain}")
    return domain


def _fetch_brand_data(domain: str, api_key: str) -> dict:
    """Fetch the full Brandfetch brand record for a resolved domain."""
    brand_url = f"https://api.brandfetch.io/v2/brands/domain/{quote(domain, safe='.')}"
    brand_resp = requests.get(
        brand_url,
        headers={"Authorization": f"Bearer {api_key}"},
        timeout=10,
    )
    return brand_resp.json()


def _palette_from_hex(hex6: str, secondary_hex: str | None) -> dict:
    """Derive the four-color palette from a primary brand hex value.

    secondary_hex, if given, is a second distinct brand color pulled straight
    from Brandfetch. Otherwise a darkened shade of primary is used so it
    reads as a genuinely different accent, not just a lighter tint.
    """
    return {
        "primary":   hex6,
        "secondary": secondary_hex or _darken(hex6, 0.35),
        "border":    _lighten(hex6, 0.25),   # 25% lighter — section borders
        "fill":      _lighten(hex6, 0.85),    # 85% lighter — competency table bg
    }


def get_brand_color(company_name: str, domain: str | None = None) -> dict:
    """
    Look up the brand accent/dark/light colors for company_name via Brandfetch.

    If domain is given (e.g. the domain already resolved and stored on the
    application tracker record), it's used directly — company-name search is
    fuzzy and can match the wrong company entirely for a common/ambiguous
    name (a "Melior" search might not resolve to getmelior.com). Only falls
    back to a name search when no domain is known yet.

    Returns a palette dict: {primary, secondary, border, fill} as 6-char
    uppercase hex strings. Falls back to DEFAULT_PALETTE on any error.
    """
    if requests is None:
        print("  ⚠ brand_color: 'requests' not installed — using default colors")
        return DEFAULT_PALETTE

    try:
        api_key = _load_api_key()

        domain = domain.strip() if domain else _resolve_domain(company_name)
        if not domain:
            print("  ⚠ Brandfetch: using default colors")
            return DEFAULT_PALETTE

        brand_data = _fetch_brand_data(domain, api_key)

        # ── Pick primary color by priority, then look for a second, distinct
        #    color among the remaining priorities to use as secondary ────────
        colors = brand_data.get("colors", [])
        hex_val = None
        chosen_type = None
        for priority in _COLOR_PRIORITY:
            match = next((c for c in colors if c.get("type") == priority), None)
            if match and match.get("hex"):
                candidate = match["hex"].lstrip("#").upper()
                if _HEX_RE.match(candidate):
                    hex_val = candidate
                    chosen_type = priority
                    break
                print(f"  ⚠ Brandfetch: malformed hex '{match['hex']}' for type '{priority}' — skipping")

        if not hex_val:
            print("  ⚠ Brandfetch: no usable color in brand data — using default colors")
            return DEFAULT_PALETTE

        hex_val = _ensure_readable(hex_val)

        secondary_hex = None
        for priority in _COLOR_PRIORITY:
            if priority == chosen_type:
                continue
            match = next((c for c in colors if c.get("type") == priority), None)
            if match and match.get("hex"):
                candidate = match["hex"].lstrip("#").upper()
                if _HEX_RE.match(candidate) and candidate != hex_val:
                    secondary_hex = candidate
                    break

        if secondary_hex:
            secondary_hex = _ensure_readable(secondary_hex)

        palette = _palette_from_hex(hex_val, secondary_hex)
        print(f"  ✓ Brandfetch: {chosen_type} color #{hex_val} → "
              f"secondary #{palette['secondary']}, border #{palette['border']}, fill #{palette['fill']}")
        return palette

    except Exception as exc:
        print(f"  ⚠ Brandfetch lookup failed ({exc}) — using default colors")
        return DEFAULT_PALETTE


def get_brand_logo(company_name: str, domain: str | None = None) -> dict | None:
    """
    Download company_name's logo via logo.dev — see the module docstring for
    why logo.dev rather than Brandfetch's own logo assets.

    If domain isn't given, it's resolved via Brandfetch's company-name search
    first (that lookup already exists for get_brand_color(), so this adds no
    extra API dependency).

    Returns {"bytes": raw image bytes, "format": "png", "width": int|None,
    "height": int|None}, or None if no logo is available or anything fails.
    """
    if requests is None:
        return None

    try:
        domain = domain.strip() if domain else _resolve_domain(company_name)
        if not domain:
            return None

        url = f"https://img.logo.dev/{domain}?token={_LOGODEV_PUBLIC_KEY}&format=png&retina=true&size=256"
        img_resp = requests.get(url, timeout=10)
        if img_resp.status_code != 200 or not img_resp.content:
            print(f"  ⚠ logo.dev: no usable logo for '{domain}'")
            return None

        dims = _png_dimensions(img_resp.content)
        width, height = dims if dims else (None, None)
        print(f"  ✓ logo.dev: fetched logo for '{domain}' ({width}x{height})")
        return {"bytes": img_resp.content, "format": "png", "width": width, "height": height}

    except Exception as exc:
        print(f"  ⚠ logo.dev logo lookup failed ({exc}) — skipping logo")
        return None
