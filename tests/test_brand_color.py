"""Unit tests for scripts/brand_color.py — brand color + logo lookup."""

from unittest.mock import MagicMock, patch

import pytest

from scripts import brand_color


def _resp(json_data=None, status_code=200, content=b""):
    r = MagicMock()
    r.status_code = status_code
    r.json.return_value = json_data if json_data is not None else {}
    r.content = content
    return r


@pytest.fixture(autouse=True)
def _api_key(monkeypatch):
    monkeypatch.setenv("BRANDFETCH_API_KEY", "test-key")


class TestGetBrandColor:
    def test_no_search_results_returns_default(self):
        with patch.object(brand_color.requests, "get", return_value=_resp([])):
            result = brand_color.get_brand_color("NobodyHeardOfThisCo")
        assert result == brand_color.DEFAULT_PALETTE

    def test_no_domain_returns_default(self):
        with patch.object(brand_color.requests, "get", return_value=_resp([{"domain": ""}])):
            result = brand_color.get_brand_color("Acme")
        assert result == brand_color.DEFAULT_PALETTE

    def test_no_usable_color_returns_default(self):
        search_resp = _resp([{"domain": "acme.com"}])
        brand_resp = _resp({"colors": []})
        with patch.object(brand_color.requests, "get", side_effect=[search_resp, brand_resp]):
            result = brand_color.get_brand_color("Acme")
        assert result == brand_color.DEFAULT_PALETTE

    def test_given_domain_skips_name_search(self):
        """Regression test: a company-name search is fuzzy and can resolve to
        the wrong company for an ambiguous name (e.g. "Melior" not resolving
        to getmelior.com). When the caller already knows the domain — e.g.
        from the application tracker record — it must be used directly,
        with no name-search call at all."""
        brand_resp = _resp({"colors": [{"type": "accent", "hex": "#123ABC"}]})
        with patch.object(brand_color.requests, "get", return_value=brand_resp) as mock_get:
            result = brand_color.get_brand_color("Melior", domain="getmelior.com")
        assert result["primary"] == "123ABC"
        assert mock_get.call_count == 1
        called_url = mock_get.call_args[0][0]
        assert "getmelior.com" in called_url
        assert "search" not in called_url

    def test_blank_domain_falls_back_to_name_search(self):
        search_resp = _resp([{"domain": "acme.com"}])
        brand_resp = _resp({"colors": [{"type": "accent", "hex": "#123ABC"}]})
        with patch.object(brand_color.requests, "get", side_effect=[search_resp, brand_resp]) as mock_get:
            result = brand_color.get_brand_color("Acme", domain="")
        assert result["primary"] == "123ABC"
        assert mock_get.call_count == 2

    def test_single_color_derives_distinct_secondary(self):
        """With only one brand color, secondary should be a darkened shade —
        not just a fixed fallback and not identical to primary or border."""
        search_resp = _resp([{"domain": "acme.com"}])
        brand_resp = _resp({"colors": [{"type": "accent", "hex": "#3366CC"}]})
        with patch.object(brand_color.requests, "get", side_effect=[search_resp, brand_resp]):
            result = brand_color.get_brand_color("Acme")
        assert result["primary"] == "3366CC"
        assert result["secondary"] not in (result["primary"], result["border"])
        assert result["secondary"] == brand_color._darken("3366CC", 0.35)

    def test_two_distinct_colors_uses_real_secondary(self):
        """When Brandfetch returns two distinct colors, secondary should be
        the real second brand color, not a derived tint."""
        search_resp = _resp([{"domain": "acme.com"}])
        brand_resp = _resp({"colors": [
            {"type": "accent", "hex": "#3366CC"},
            {"type": "dark", "hex": "#112233"},
        ]})
        with patch.object(brand_color.requests, "get", side_effect=[search_resp, brand_resp]):
            result = brand_color.get_brand_color("Acme")
        assert result["primary"] == "3366CC"
        assert result["secondary"] == "112233"

    def test_malformed_hex_is_skipped(self):
        search_resp = _resp([{"domain": "acme.com"}])
        brand_resp = _resp({"colors": [{"type": "accent", "hex": "not-a-color"}]})
        with patch.object(brand_color.requests, "get", side_effect=[search_resp, brand_resp]):
            result = brand_color.get_brand_color("Acme")
        assert result == brand_color.DEFAULT_PALETTE

    def test_light_primary_is_darkened_for_readability(self):
        """A pale 'accent' color (fine for UI chrome, unreadable as page text)
        must be darkened before it's used as the primary text color."""
        search_resp = _resp([{"domain": "acme.com"}])
        brand_resp = _resp({"colors": [{"type": "accent", "hex": "#FFF3B0"}]})
        with patch.object(brand_color.requests, "get", side_effect=[search_resp, brand_resp]):
            result = brand_color.get_brand_color("Acme")
        assert result["primary"] != "FFF3B0"
        assert brand_color._contrast_on_white(result["primary"]) >= brand_color._MIN_TEXT_CONTRAST

    def test_dark_primary_passes_through_unchanged(self):
        search_resp = _resp([{"domain": "acme.com"}])
        brand_resp = _resp({"colors": [{"type": "accent", "hex": "#112233"}]})
        with patch.object(brand_color.requests, "get", side_effect=[search_resp, brand_resp]):
            result = brand_color.get_brand_color("Acme")
        assert result["primary"] == "112233"

    def test_light_secondary_is_also_darkened(self):
        search_resp = _resp([{"domain": "acme.com"}])
        brand_resp = _resp({"colors": [
            {"type": "accent", "hex": "#112233"},
            {"type": "dark", "hex": "#FFF3B0"},
        ]})
        with patch.object(brand_color.requests, "get", side_effect=[search_resp, brand_resp]):
            result = brand_color.get_brand_color("Acme")
        assert result["secondary"] != "FFF3B0"
        assert brand_color._contrast_on_white(result["secondary"]) >= brand_color._MIN_TEXT_CONTRAST

    def test_exception_returns_default(self):
        with patch.object(brand_color.requests, "get", side_effect=Exception("network down")):
            result = brand_color.get_brand_color("Acme")
        assert result == brand_color.DEFAULT_PALETTE

    def test_requests_unavailable_returns_default(self):
        with patch.object(brand_color, "requests", None):
            result = brand_color.get_brand_color("Acme")
        assert result == brand_color.DEFAULT_PALETTE


# A real, valid 1x1 PNG — used to exercise the IHDR dimension parser end to end.
_TINY_PNG = (
    b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
    b"\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx\x9cc\xf8\xcf\xc0\x00"
    b"\x00\x03\x01\x01\x00\x18\xdd\x8d\xb0\x00\x00\x00\x00IEND\xaeB`\x82"
)


class TestGetBrandLogo:
    """get_brand_logo now fetches the actual image from logo.dev — the same
    source used everywhere else in the app for company logos — rather than
    Brandfetch's logo assets, which can be a "light" variant that disappears
    against a white document background. Brandfetch's name search is still
    used as a fallback to resolve a domain when none is given, since
    get_brand_color() already depends on that same lookup."""

    def test_no_domain_returns_none(self):
        with patch.object(brand_color.requests, "get", return_value=_resp([])):
            assert brand_color.get_brand_logo("Acme") is None

    def test_given_domain_skips_name_search_and_fetches_from_logodev(self):
        img_resp = _resp(content=_TINY_PNG)
        with patch.object(brand_color.requests, "get", return_value=img_resp) as mock_get:
            result = brand_color.get_brand_logo("Melior", domain="getmelior.com")
        assert result["bytes"] == _TINY_PNG
        assert result["format"] == "png"
        assert result["width"] == 1 and result["height"] == 1
        mock_get.assert_called_once()
        called_url = mock_get.call_args[0][0]
        assert called_url.startswith("https://img.logo.dev/getmelior.com?")
        assert "search" not in called_url

    def test_blank_domain_resolves_via_name_search_then_logodev(self):
        search_resp = _resp([{"domain": "acme.com"}])
        img_resp = _resp(content=_TINY_PNG)
        with patch.object(brand_color.requests, "get", side_effect=[search_resp, img_resp]) as mock_get:
            result = brand_color.get_brand_logo("Acme", domain="")
        assert result["bytes"] == _TINY_PNG
        assert mock_get.call_count == 2
        assert "img.logo.dev/acme.com" in mock_get.call_args_list[-1][0][0]

    def test_parses_real_png_dimensions(self):
        with patch.object(brand_color.requests, "get", return_value=_resp(content=_TINY_PNG)):
            result = brand_color.get_brand_logo("Acme", domain="acme.com")
        assert result["width"] == 1
        assert result["height"] == 1

    def test_unparseable_bytes_returns_none_dimensions(self):
        with patch.object(brand_color.requests, "get", return_value=_resp(content=b"not-a-real-png")):
            result = brand_color.get_brand_logo("Acme", domain="acme.com")
        assert result["width"] is None
        assert result["height"] is None

    def test_image_download_failure_returns_none(self):
        with patch.object(brand_color.requests, "get", return_value=_resp(status_code=404, content=b"")):
            assert brand_color.get_brand_logo("Acme", domain="acme.com") is None

    def test_exception_returns_none(self):
        with patch.object(brand_color.requests, "get", side_effect=Exception("network down")):
            assert brand_color.get_brand_logo("Acme", domain="acme.com") is None

    def test_requests_unavailable_returns_none(self):
        with patch.object(brand_color, "requests", None):
            assert brand_color.get_brand_logo("Acme") is None


class TestDarkenLighten:
    def test_darken_moves_toward_black(self):
        assert brand_color._darken("FFFFFF", 0.5) == "808080"
        assert brand_color._darken("336699", 0.0) == "336699"

    def test_lighten_moves_toward_white(self):
        assert brand_color._lighten("000000", 0.5) == "808080"
        assert brand_color._lighten("336699", 0.0) == "336699"


class TestEnsureReadable:
    def test_already_readable_is_unchanged(self):
        assert brand_color._ensure_readable("1A3C5E") == "1A3C5E"

    def test_white_is_darkened_to_readable(self):
        result = brand_color._ensure_readable("FFFFFF")
        assert result != "FFFFFF"
        assert brand_color._contrast_on_white(result) >= brand_color._MIN_TEXT_CONTRAST

    def test_pale_yellow_is_darkened_to_readable(self):
        result = brand_color._ensure_readable("FFF3B0")
        assert brand_color._contrast_on_white(result) >= brand_color._MIN_TEXT_CONTRAST

    def test_black_has_max_contrast(self):
        assert brand_color._contrast_on_white("000000") == pytest.approx(21.0, rel=0.01)

    def test_default_palette_primary_is_already_readable(self):
        assert brand_color._contrast_on_white(
            brand_color.DEFAULT_PALETTE["primary"]
        ) >= brand_color._MIN_TEXT_CONTRAST
