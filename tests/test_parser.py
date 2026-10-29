"""Parser: tokens, classification order, m resolution, refused and accepted odd formats."""

from decimal import Decimal

import pytest

from xlformat import FormatError, NumberFormat, SectionKind, format_cell, format_value
from xlformat._parser import TokenKind, split_sections, tokenize


def kinds(section: str) -> list[tuple[str, str]]:
    return [(t.kind.value, t.text) for t in tokenize(section)]


@pytest.mark.parametrize(
    ("source", "expected"),
    [
        ('"Total: "', [("literal", "Total: ")]),
        ("\\x", [("literal", "x")]),
        ("$-+():!^&~{}<>= ", [("literal", c) for c in "$-+():!^&~{}<>= "]),
        ("[Red]", [("color", "Red")]),
        ("[blue]", [("color", "Blue")]),
        ("[Color 3]", [("color", "Color3")]),
        ("[COLOR56]", [("color", "Color56")]),
        ("[<=100]", [("condition", "[<=100]")]),
        ("[$€-x-euro2]", [("currency", "€")]),
        ("[$-en-US]", [("currency", "")]),
        ("[h]", [("elapsed", "h")]),
        ("[mm]", [("elapsed", "mm")]),
        ("[SSS]", [("elapsed", "sss")]),
        ("General", [("general", "General")]),
        ("GENERAL", [("general", "General")]),
        ("@", [("at", "@")]),
        ("0#?", [("digit", "0"), ("digit", "#"), ("digit", "?")]),
        (".", [("point", ".")]),
        (",", [("comma", ",")]),
        ("%", [("percent", "%")]),
        ("E+", [("exp", "E+")]),
        ("e-", [("exp", "e-")]),
        ("/", [("slash", "/")]),
        ("yyy", [("date", "yyyy")]),
        ("Y", [("date", "yy")]),
        ("mmmmmm", [("date", "mmmm")]),
        ("ddddd", [("date", "dddd")]),
        ("HH", [("date", "hh")]),
        ("sss", [("date", "ss")]),
        ("e", [("date", "e")]),
        ("AM/PM", [("ampm", "AM/PM")]),
        ("a/p", [("ampm", "a/p")]),
        ("_(", [("pad", "(")]),
        ("*-", [("fill", "-")]),
        ("x", [("literal", "x")]),
    ],
)
def test_token_classes(source: str, expected: list[tuple[str, str]]) -> None:
    assert kinds(source) == expected


def test_condition_value_is_decimal() -> None:
    (tok,) = tokenize("[>=1.5]")
    assert tok.condition == (">=", Decimal("1.5"))


def test_currency_locale_is_recorded() -> None:
    (tok,) = tokenize("[$€-x-euro2]")
    assert tok.locale == "x-euro2"
    assert NumberFormat.parse("[$-en-US]h:mm").sections[0].locale == "en-US"


def test_split_sections_respects_quotes_escapes_and_brackets() -> None:
    assert split_sections(r'"a;b";\;;[>1]0;@') == ['"a;b"', r"\;", "[>1]0", "@"]


@pytest.mark.parametrize(
    ("source", "kind"),
    [
        ("", SectionKind.GENERAL),  # an empty first section behaves as General
        ("General", SectionKind.GENERAL),
        ('"Yes"', SectionKind.GENERAL),
        ("@", SectionKind.TEXT),
        ("0.00", SectionKind.NUMBER),
        (".", SectionKind.NUMBER),
        ("0.00E+00", SectionKind.SCIENTIFIC),
        ("# ?/?", SectionKind.FRACTION),
        ("m/d/yy", SectionKind.DATETIME),
        ("[h]:mm:ss", SectionKind.DATETIME),
        ("mmss.0", SectionKind.DATETIME),
        ("e+", SectionKind.DATETIME),  # bare e is a year code, + a literal
        ("0%E+0", SectionKind.SCIENTIFIC),
        ("( General )", SectionKind.GENERAL),
    ],
)
def test_classification(source: str, kind: SectionKind) -> None:
    assert NumberFormat.parse(source).sections[0].kind is kind


def test_empty_later_section_is_empty_but_colored_one_is_general() -> None:
    fmt = NumberFormat.parse("0;;[Red]")
    assert [s.kind for s in fmt.sections] == [
        SectionKind.NUMBER,
        SectionKind.EMPTY,
        SectionKind.GENERAL,
    ]


def date_codes(source: str) -> list[tuple[str, str]]:
    return [
        (t.kind.value, t.text)
        for t in NumberFormat.parse(source).sections[0].tokens
        if t.kind in (TokenKind.DATE, TokenKind.MINUTE, TokenKind.ELAPSED)
    ]


@pytest.mark.parametrize(
    ("source", "expected"),
    [
        ("m/d/yy", [("date", "m"), ("date", "d"), ("date", "yy")]),
        ("h:mm", [("date", "h"), ("minute", "mm")]),
        ("mm:ss", [("minute", "mm"), ("date", "ss")]),
        (
            "m/d/yy h:mm",
            [("date", "m"), ("date", "d"), ("date", "yy"), ("date", "h"), ("minute", "mm")],
        ),
        ("h:mm:ss m", [("date", "h"), ("minute", "mm"), ("date", "ss"), ("date", "m")]),
        ("[h]:mm:ss", [("elapsed", "h"), ("minute", "mm"), ("date", "ss")]),
        ("mmss.0", [("minute", "mm"), ("date", "ss.0")]),
        (
            "yyyy-mm-dd hh:mm:ss.000",
            [
                ("date", "yyyy"),
                ("date", "mm"),
                ("date", "dd"),
                ("date", "hh"),
                ("minute", "mm"),
                ("date", "ss.000"),
            ],
        ),
    ],
)
def test_m_resolution(source: str, expected: list[tuple[str, str]]) -> None:
    assert date_codes(source) == expected


REFUSED = [
    "0.00e+00",
    "0.00e-00",
    "##0.0e+0",
    "0.00E+00;0.00e+00",
    "0;0;0;0;0",
    "[Red",
    '0.00"',
    "[Red][Blue]0",
    "[>1][<5]0",
    "@;@",
    "0;@;0",
    "yyyy 0",
    "0 yyyy",
    "0.00E+",
    "# ?/",
    "[hh]:mm:[ss]",
    "General General",
]


@pytest.mark.parametrize("source", REFUSED)
def test_formats_excel_refuses_in_a_file_raise(source: str) -> None:
    with pytest.raises(FormatError):
        NumberFormat.parse(source)


def test_accepted_odd_formats_render_like_excel() -> None:
    assert format_value(1234.5678, "*-0*-") == "1235"  # fill emits nothing without a width
    assert format_value(1234.5678, "0.0.0") == "1234.5.7"
    assert format_value(1234.5678, "0 x") == "1235 x"
    assert format_value(-0.5, "x0") == "-x1"
    assert format_value("abc", "x0") == "abc"
    assert format_cell(1234.5678, "[Color 3]0").color == "Color3"
    assert NumberFormat.parse("[Red]").sections[0].kind is SectionKind.GENERAL  # [Red]General
    assert NumberFormat.parse("e+").sections[0].kind is SectionKind.DATETIME
    assert NumberFormat.parse("0%E+0").sections[0].kind is SectionKind.SCIENTIFIC
    assert NumberFormat.parse("*-0*-").sections[0].kind is SectionKind.NUMBER


def test_parse_is_cached_and_empty_means_general() -> None:
    assert NumberFormat.parse("#,##0.00") is NumberFormat.parse("#,##0.00")
    assert NumberFormat.parse("").is_general
    assert NumberFormat.parse("General").is_general
    assert NumberFormat.parse("m/d/yy;@").is_date
    assert NumberFormat.parse("@").is_text
    assert not NumberFormat.parse("0;@").is_text


def test_number_spec_comma_rules() -> None:
    spec = NumberFormat.parse("#,##0.00,").sections[0].number
    assert spec is not None
    assert spec.grouping
    assert (spec.scale, spec.n_integer, spec.n_decimal) == (1, 4, 2)
    spec = NumberFormat.parse(",0").sections[0].number
    assert spec is not None
    assert spec.leading == (("lit", ","),)
    assert not spec.grouping


def test_width_character_mode() -> None:
    assert format_value(1234567.891, "General", width=8) == "1234568"
    assert format_value(12345678, "General", width=5) == "1E+07"
    assert format_value(0.000123, "General", width=4) == "0"
    assert format_value(-1, "m/d/yy", width=6) == "######"
    assert format_value("hello world", "@", width=3) == "hello world"
    assert format_value(1234.5, "0.00*-", width=10) == "1234.50---"
    assert format_value(True, "General", width=3) == "###"
    with pytest.raises(ValueError, match="width"):
        format_value(1, "0", width=0)


def test_width_excel_mode() -> None:
    # Calibri 11: digits 6 points, '-' 3, ',' and '.' 3, '%' 8, '#' 5; a column of w units has
    # 6*w points for text (CAPTURING.md section 2 items 24 and 25)
    assert format_value(1000, "#,##0.0", column_width=6) == "1,000.0"  # 36 points in 36
    assert format_value(-12345678, "General", column_width=8.43) == "-12345678"  # 8.5 units
    assert format_value(100, "0%", column_width=6) == "#######"  # 38 points > 36; 36 // 5 = 7
    assert format_value(1234.5, "0.00*-", column_width=30, font=("Calibri", 12)) == (
        "1234.50" + "-" * 35
    )
    assert format_value(-1234.5678, "General", column_width=6) == "-1235"  # by characters
    assert format_value("Hello World", "@", column_width=3) == "Hello World"
    cell = format_cell(-1, "m/d/yy", column_width=6)
    assert cell.overflow
    assert cell.text == "#######"
    assert cell.full_text == "#" * 255
    with pytest.raises(ValueError, match="not both"):
        format_value(1, "0", width=6, column_width=6)
    with pytest.raises(LookupError, match="Aptos Narrow"):
        format_value(1, "0", column_width=6, font="Comic Sans")
    assert format_value(12345, "0", width=10, measure=lambda t: 2.0 * len(t)) == "12345"
    # the budget is in the measure's units: 12 > 10 overflows, and 10 // 2 gives five '#'
    assert format_value(123456, "0", width=10, measure=lambda t: 2.0 * len(t)) == "#####"
