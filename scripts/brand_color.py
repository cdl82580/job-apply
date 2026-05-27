#!/usr/bin/env python3
"""
Fetch brand colors from Brandfetch API.

API key is read from the BRANDFETCH_API_KEY environment variable or .env file
in the project root. Never hardcode the key in source.

Usage:
    from scripts.brand_color import get_brand_color
    colors = get_brand_color("Brightflag")
    # {"primary": "1A3C5E", "border": "2B6CB0", "fill": "EEF4FB"}
"""

import os
from pathlib import Path
from urllib.parse import quote

try:
    import requests
except ImportError:
    requests = None  # type: ignore

# Fallback palette — Corey's default navy scheme
DEFAULT_PALETTE = {
    "primary": "1A3C5E",
    "border":  "2B6CB0",
    "fill":    "EEF4FB",
}

# Brandfetch search client token (autocomplete endpoint). Set BRANDFETCH_SEARCH_CLIENT in env.
_SEARCH_CLIENT = os.environ.get("BRANDFETCH_SEARCH_CLIENT", "")


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


def _palette_from_hex(hex6: str) -> dict:
    """Derive the three-color palette from a single brand hex value."""
    return {
        "primary": hex6,
        "border":  _lighten(hex6, 0.25),   # 25% lighter — section borders
        "fill":    _lighten(hex6, 0.85),    # 85% lighter — competency table bg
    }


def get_brand_color(company_name: str) -> dict:
    """
    Look up the brand accent/dark/light color for company_name via Brandfetch.
    Returns a palette dict: {primary, border, fill} as 6-char uppercase hex strings.
    Falls back to DEFAULT_PALETTE on any error.
    """
    if requests is None:
        print("  ⚠ brand_color: 'requests' not installed — using default colors")
        return DEFAULT_PALETTE

    try:
        api_key = _load_api_key()

        # ── Step 1: search for the company's domain ───────────────────────────
        search_url = f"https://api.brandfetch.io/v2/search/{quote(company_name)}"
        search_resp = requests.get(
            search_url,
            params={"c": _SEARCH_CLIENT},
            timeout=10,
        )
        search_data = search_resp.json()

        if not isinstance(search_data, list) or not search_data:
            print(f"  ⚠ Brandfetch: no results for '{company_name}' — using default colors")
            return DEFAULT_PALETTE

        domain = search_data[0].get("domain", "")
        if not domain:
            print(f"  ⚠ Brandfetch: no domain in search result — using default colors")
            return DEFAULT_PALETTE

        print(f"  ✓ Brandfetch: resolved '{company_name}' → {domain}")

        # ── Step 2: fetch brand data by domain ────────────────────────────────
        brand_url = f"https://api.brandfetch.io/v2/brands/domain/{quote(domain, safe='.')}"
        brand_resp = requests.get(
            brand_url,
            headers={"Authorization": f"Bearer {api_key}"},
            timeout=10,
        )
        brand_data = brand_resp.json()

        # ── Step 3: pick color by priority ────────────────────────────────────
        colors = brand_data.get("colors", [])
        hex_val = None
        chosen_type = None
        for priority in ("accent", "dark", "light"):
            match = next((c for c in colors if c.get("type") == priority), None)
            if match and match.get("hex"):
                hex_val = match["hex"].lstrip("#").upper()
                chosen_type = priority
                break

        if not hex_val:
            print("  ⚠ Brandfetch: no usable color in brand data — using default colors")
            return DEFAULT_PALETTE

        palette = _palette_from_hex(hex_val)
        print(f"  ✓ Brandfetch: {chosen_type} color #{hex_val} → "
              f"border #{palette['border']}, fill #{palette['fill']}")
        return palette

    except Exception as exc:
        print(f"  ⚠ Brandfetch lookup failed ({exc}) — using default colors")
        return DEFAULT_PALETTE
