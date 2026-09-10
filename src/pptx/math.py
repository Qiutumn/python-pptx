"""Editable Office Math (OMML) support for PowerPoint text."""

from __future__ import annotations

import re
from copy import deepcopy
from importlib import import_module
from typing import TYPE_CHECKING, Any, cast

from lxml import etree
from lxml.etree import XMLSyntaxError, _Element  # pyright: ignore[reportPrivateUsage]

from pptx.exc import PythonPptxError
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsmap, qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.shapes import Subshape
from pptx.util import Emu, Length

if TYPE_CHECKING:
    from pptx.dml.color import RGBColor
    from pptx.types import ProvidesPart


_A14_NS = "http://schemas.microsoft.com/office/drawing/2010/main"
_DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_MAX_OMML_LENGTH = 1_048_576
_SHAPE_TAGS = frozenset(
    {
        qn("p:sp"),
        qn("p:grpSp"),
        qn("p:graphicFrame"),
        qn("p:cxnSp"),
        qn("p:pic"),
        qn("p:contentPart"),
    }
)


class MathConversionError(PythonPptxError):
    """Raised when a math expression cannot be converted to valid PowerPoint OMML."""


class Math(Subshape):
    """An editable Office Math zone contained in a text paragraph."""

    def __init__(self, math_elm: _Element, parent: ProvidesPart):
        super(Math, self).__init__(parent)
        self._element = math_elm

    @property
    def display(self) -> bool:
        """True when this is a standalone (display) equation."""
        return self._omml.tag == qn("m:oMathPara")

    @property
    def element(self) -> _Element:
        """The underlying ``a14:m`` XML element."""
        return self._element

    @property
    def omml(self) -> str:
        """The equation as a standalone OMML XML fragment."""
        return etree.tostring(self._omml, encoding="unicode")

    @property
    def text(self) -> str:
        """Linear Unicode text collected from the equation's ``m:t`` elements."""
        return "".join(t.text or "" for t in self._omml.iter(qn("m:t")))

    @property
    def _omml(self) -> _Element:
        children = list(self._element)
        if len(children) != 1:  # pragma: no cover - guarded during construction
            raise MathConversionError("a14:m must contain exactly one OMML root")
        return children[0]


def add_math_to_paragraph(
    paragraph: _Element,
    parent: ProvidesPart,
    omml: str,
    *,
    display: bool | None = None,
    font_size: Length | None = None,
    color: RGBColor | None = None,
) -> Math:
    """Append an editable OMML expression to ``paragraph`` and return its proxy."""
    math_root = _parse_omml(omml, display=display)
    _apply_math_style(math_root, font_size, color)
    _ensure_math_shape_wrapper(paragraph)
    math_elm = cast(_Element, OxmlElement("a14:m", nsmap=nsmap("a14", "m", "a")))
    math_elm.append(math_root)
    endParaRPr = paragraph.find(qn("a:endParaRPr"))
    if endParaRPr is None:
        paragraph.append(math_elm)
    else:
        endParaRPr.addprevious(math_elm)
    return Math(math_elm, parent)


def latex_to_omml(latex: str, *, display: bool = False) -> str:
    """Convert a LaTeX expression to an editable PowerPoint OMML fragment.

    The optional ``math`` extra supplies the pure-Python converters used here::

        pip install "python-pptx[math]"
    """
    if not latex.strip():
        raise MathConversionError("LaTeX expression must be a non-empty string")

    latex2mathml = _optional_module("latex2mathml.converter")
    try:
        mathml = latex2mathml.convert(latex, display="block" if display else "inline")
    except Exception as exc:
        raise MathConversionError(f"could not convert LaTeX to MathML: {exc}") from exc
    return mathml_to_omml(mathml, display=display)


def mathml_to_omml(mathml: str, *, display: bool = False) -> str:
    """Convert a MathML expression to an editable PowerPoint OMML fragment."""
    if not mathml.strip():
        raise MathConversionError("MathML expression must be a non-empty string")

    mathml2omml = _optional_module("mathml2omml")
    try:
        omml = mathml2omml.convert(mathml)
    except Exception as exc:
        raise MathConversionError(f"could not convert MathML to OMML: {exc}") from exc
    # mathml2omml 0.0.2 closes groupChrPr with </m:groupChr> for stretchy
    # over/underscripts. Normalize only that converter-emitted property shape;
    # direct user-supplied OMML still goes through strict XML validation.
    omml = re.sub(
        r'(<m:groupChrPr><m:chr m:val="[^"]*"/><m:pos m:val="(?:top|bot)"/>)</m:groupChr>',
        r"\1</m:groupChrPr>",
        omml,
    )
    root = _parse_omml(omml, display=display)
    _normalize_converted_accents(root)
    return _serialize_omml(root)


def _normalize_converted_accents(root: _Element) -> None:
    """Keep converted accents on the base's baseline and stretch over its width."""
    for element in list(root.iter()):
        tag = etree.QName(element).localname
        kind, character, position = None, None, None
        if tag == "groupChr":
            properties = element.find(qn("m:groupChrPr"))
            if properties is None:
                continue
            marker, placement = properties.find(qn("m:chr")), properties.find(qn("m:pos"))
            character = marker.get(qn("m:val")) if marker is not None else None
            position = placement.get(qn("m:val"), "bot") if placement is not None else "bot"
            if character in {"¯", "―"}:
                kind = "bar"
            elif position == "top" and character in {"←", "→", "↔"}:
                kind = "acc"
                character = {"←": "\u20d6", "→": "\u20d7", "↔": "\u20e1"}[character]
            else:
                # With a top group character, Office otherwise puts the base
                # below the surrounding baseline and shrinks it like a script.
                if properties.find(qn("m:vertJc")) is None:
                    align = etree.SubElement(properties, qn("m:vertJc"))
                    align.set(qn("m:val"), "bot" if position == "top" else "top")
        elif tag in {"limUpp", "limLow"}:
            limit = element.find(qn("m:lim"))
            if limit is not None and len(limit) == 1 and limit[0].tag == qn("m:r"):
                text = limit[0].find(qn("m:t"))
                if text is not None and text.text == "―":
                    kind, position = "bar", "top" if tag == "limUpp" else "bot"
        base = element.find(qn("m:e"))
        parent = element.getparent()
        if kind is None or base is None or parent is None:
            continue
        replacement = etree.Element(qn("m:" + kind))
        properties = etree.SubElement(replacement, qn("m:" + kind + "Pr"))
        option = etree.SubElement(properties, qn("m:pos" if kind == "bar" else "m:chr"))
        option.set(qn("m:val"), position if kind == "bar" else character)
        replacement.append(base)
        parent.replace(element, replacement)


def _apply_math_style(
    omml: _Element,
    font_size: Length | None,
    color: RGBColor | None,
) -> None:
    """Apply PowerPoint run properties to all visible leaves and controls in ``omml``."""
    if font_size is not None:
        size_centipoints = Emu(font_size).centipoints
        if size_centipoints <= 0:
            raise ValueError("math font size must be greater than zero")
    else:
        size_centipoints = None
    for run in omml.iter(qn("m:r")):
        _merge_drawingml_run_properties(run, font_size=size_centipoints, color=color)

    control_property_owners = (
        "m:accPr",
        "m:barPr",
        "m:borderBoxPr",
        "m:boxPr",
        "m:dPr",
        "m:fPr",
        "m:groupChrPr",
        "m:naryPr",
        "m:radPr",
    )
    for owner_tag in control_property_owners:
        for owner in omml.iter(qn(owner_tag)):
            ctrlPr = owner.find(qn("m:ctrlPr"))
            if ctrlPr is None:
                ctrlPr = OxmlElement("m:ctrlPr")
                owner.append(ctrlPr)
            _merge_drawingml_run_properties(
                ctrlPr, font_size=size_centipoints, color=color, control=True
            )


def _ensure_math_shape_wrapper(paragraph: _Element) -> None:
    """Place the paragraph's shape in the MCE Choice required by ``a14:m``."""
    shape: _Element | None = None
    node: _Element | None = paragraph
    while node is not None:
        if node.tag in _SHAPE_TAGS:
            shape = node
            break
        node = node.getparent()

    # -- A detached paragraph is useful in unit tests and fragment-building workflows. --
    if shape is None or shape.getparent() is None:
        return

    parent = shape.getparent()
    assert parent is not None
    if parent.tag == qn("mc:Choice"):
        return
    if parent.tag not in {qn("p:spTree"), qn("p:grpSp")}:
        raise MathConversionError("equations require a slide-local shape or group shape")

    alternate = cast(_Element, OxmlElement("mc:AlternateContent", nsmap={"mc": _MC_NS}))
    choice = cast(_Element, OxmlElement("mc:Choice", nsmap={"a14": _A14_NS}))
    choice.set("Requires", "a14")

    index = parent.index(shape)
    parent.remove(shape)
    choice.append(shape)
    alternate.append(choice)
    parent.insert(index, alternate)


def _merge_drawingml_run_properties(
    owner: _Element,
    *,
    font_size: int | None,
    color: RGBColor | None,
    control: bool = False,
) -> None:
    """Merge defaults into one ``a:rPr`` under an OMML run or control."""
    rPr = owner.find(qn("a:rPr"))
    if rPr is None:
        rPr = cast(_Element, OxmlElement("a:rPr"))
        insert_at = 1 if not control and len(owner) and owner[0].tag == qn("m:rPr") else 0
        owner.insert(insert_at, rPr)

    if rPr.get("lang") is None:
        rPr.set("lang", "en-US")
    rPr.set("dirty", "0")
    if font_size is not None:
        rPr.set("sz", str(font_size))

    if color is not None:
        fill_tags = {
            qn("a:blipFill"),
            qn("a:gradFill"),
            qn("a:grpFill"),
            qn("a:noFill"),
            qn("a:pattFill"),
            qn("a:solidFill"),
        }
        for child in list(rPr):
            if child.tag in fill_tags:
                rPr.remove(child)
        solidFill = OxmlElement("a:solidFill")
        srgbClr = OxmlElement("a:srgbClr")
        srgbClr.set("val", str(color))
        solidFill.append(srgbClr)
        rPr.insert(0, solidFill)

    for font_tag in ("a:latin", "a:ea", "a:cs"):
        font = rPr.find(qn(font_tag))
        if font is None:
            font = OxmlElement(font_tag)
            font.set("typeface", "Cambria Math")
            rPr.append(font)


def _optional_module(module_name: str) -> Any:
    try:
        return import_module(module_name)
    except ImportError as exc:
        raise MathConversionError(
            'LaTeX and MathML conversion requires the optional "math" extra; '
            'install it with: pip install "python-pptx[math]"'
        ) from exc


def _parse_omml(omml: str, *, display: bool | None = None) -> _Element:
    """Parse, validate, and normalize one OMML expression."""
    if not omml.strip():
        raise MathConversionError("OMML expression must be a non-empty string")
    if len(omml) > _MAX_OMML_LENGTH:
        raise MathConversionError(f"OMML expression exceeds {_MAX_OMML_LENGTH} characters")
    upper = omml.upper()
    if "<!DOCTYPE" in upper or "<!ENTITY" in upper:
        raise MathConversionError("DOCTYPE and ENTITY declarations are not allowed in OMML")

    wrapper_xml = (
        f'<pptx-math-root xmlns:m="{_MATH_NS}" xmlns:a="{_DML_NS}">{omml}</pptx-math-root>'
    )
    try:
        wrapper = parse_xml(wrapper_xml)
    except (XMLSyntaxError, ValueError) as exc:
        raise MathConversionError(f"invalid OMML XML: {exc}") from exc
    if len(wrapper) != 1 or (wrapper.text or "").strip() or (wrapper[0].tail or "").strip():
        raise MathConversionError("OMML must contain exactly one root element")

    root = deepcopy(wrapper[0])
    if root.tag not in {qn("m:oMath"), qn("m:oMathPara")}:
        raise MathConversionError("OMML root must be m:oMath or m:oMathPara")

    for element in root.iter():
        if not isinstance(element.tag, str):
            raise MathConversionError("OMML comments and processing instructions are not allowed")
        namespace = element.tag[1:].split("}", 1)[0] if element.tag.startswith("{") else None
        if namespace not in {_MATH_NS, _DML_NS}:
            raise MathConversionError(f"unsupported namespace in OMML: {namespace!r}")

    if root.tag == qn("m:oMathPara"):
        expressions = [child for child in root if child.tag == qn("m:oMath")]
        if len(expressions) != 1:
            raise MathConversionError("m:oMathPara must contain exactly one m:oMath expression")

    if display is True and root.tag == qn("m:oMath"):
        expression = root
        root = cast(_Element, OxmlElement("m:oMathPara", nsmap=nsmap("m", "a")))
        properties = cast(_Element, OxmlElement("m:oMathParaPr"))
        justification = cast(_Element, OxmlElement("m:jc"))
        justification.set(qn("m:val"), "centerGroup")
        properties.append(justification)
        root.append(properties)
        root.append(expression)
    elif display is False and root.tag == qn("m:oMathPara"):
        expression = root.find(qn("m:oMath"))
        assert expression is not None
        root = deepcopy(expression)

    if not any((text.text or "") for text in root.iter(qn("m:t"))):
        raise MathConversionError("OMML expression contains no math text")
    return root


def _serialize_omml(omml: _Element) -> str:
    return etree.tostring(omml, encoding="unicode")


__all__ = [
    "Math",
    "MathConversionError",
    "add_math_to_paragraph",
    "latex_to_omml",
    "mathml_to_omml",
]
