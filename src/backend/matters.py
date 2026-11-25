"""Copy announcement/matters slides from a previous deck into a generated deck.

The copy is done at the OPC/XML level so pictures, tables, grouped shapes, charts
(including embedded workbooks), background images, and other related parts come
along. The target slide keeps the generated deck's ``matters`` layout, so the
template's title and decorations are preserved.
"""

from __future__ import annotations

import copy
import re
from io import BytesIO
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.constants import RELATIONSHIP_TARGET_MODE as RTM
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part, _Relationship
from pptx.opc.packuri import PackURI
from pptx.parts.slide import SlidePart

REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Structural relationships are not part of the slide content. Notes are skipped
# because cloning a notes slide would create a back-reference to the source slide.
SKIPPED_RELTYPES = {RT.SLIDE_LAYOUT, RT.NOTES_SLIDE, RT.NOTES_MASTER}


def _layout_name(slide) -> str:
    return slide.slide_layout.name.strip().lower()


def _slide_has_content(slide) -> bool:
    """Return True when a slide contains something worth copying."""
    if slide._element.xpath(".//p:bg//a:blip"):
        return True

    content_types = {
        MSO_SHAPE_TYPE.PICTURE,
        MSO_SHAPE_TYPE.TABLE,
        MSO_SHAPE_TYPE.CHART,
        MSO_SHAPE_TYPE.GROUP,
    }
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text.strip():
            return True
        if shape.shape_type in content_types:
            return True
        # Picture-filled placeholders are reported as placeholders, not pictures.
        if shape._element.xpath(".//a:blip"):
            return True
    return False


def find_matters_slides(prs: Presentation) -> list:
    """Return the source deck's announcement slides in presentation order."""
    slides = list(prs.slides)

    exact = [
        slide
        for slide in slides
        if _layout_name(slide) == "matters" and _slide_has_content(slide)
    ]
    if exact:
        return exact

    announced = [
        slide
        for slide in slides
        if "announce" in _layout_name(slide) and _slide_has_content(slide)
    ]
    if announced:
        return announced

    # Older decks may not use a dedicated matters layout. Avoid the worship-order
    # slide, which also mentions "事项报告 / Announcements".
    text_matches = [
        slide
        for slide in slides
        if _layout_name(slide) != "order"
        and any(
            label in "\n".join(
                shape.text for shape in slide.shapes if shape.has_text_frame
            )
            for label in ("事项报告", "事项汇报")
        )
    ]
    return text_matches


def _partname_template(partname) -> str:
    """Return a ``%d`` partname template matching ``partname``."""
    directory, _, filename = str(partname).rpartition("/")
    stem, dot, ext = filename.rpartition(".")
    if not dot:
        stem, ext = filename, ""

    # `next_partname` uses %-formatting, so escape any literal percent signs.
    stem = stem.replace("%", "%%")
    if re.search(r"\d+$", stem):
        stem = re.sub(r"\d+$", "%d", stem)
    else:
        stem = f"{stem}_copy%d"

    return f"{directory}/{stem}.{ext}" if ext else f"{directory}/{stem}"


def _next_partname(dst_pkg, tmpl: str, used: set[str]) -> PackURI:
    """Return the next unused partname for ``tmpl``.

    Newly-created parts are not reachable from the package until their parent
    relationship is added, so track reserved names separately.
    """
    n = 1
    while True:
        candidate = tmpl % n
        if candidate not in used:
            used.add(candidate)
            return PackURI(candidate)
        n += 1


def _clone_part(
    source_part,
    dst_pkg,
    memo: dict,
    source_slide_part,
    dst_slide_part,
    used: set[str],
    warnings: list[str],
):
    """Deep-copy an opaque related part and its relationships.

    Parts are copied as generic ``Part`` objects so their original binary/XML
    payload is preserved verbatim. Relationship ids are kept unchanged because
    the cloned payload still references them.
    """
    if source_part is source_slide_part:
        return dst_slide_part
    if source_part in memo:
        return memo[source_part]
    if isinstance(source_part, SlidePart):
        warnings.append(f"Skipped related slide part {source_part.partname}.")
        return None

    cloned = Part.load(
        _next_partname(dst_pkg, _partname_template(source_part.partname), used),
        source_part.content_type,
        dst_pkg,
        source_part.blob,
    )
    memo[source_part] = cloned

    for rId, rel in source_part.rels.items():
        if rel.reltype in SKIPPED_RELTYPES:
            continue

        if rel.is_external:
            target, is_external = rel.target_ref, True
        else:
            target_part = rel.target_part
            if target_part is source_slide_part:
                target, is_external = dst_slide_part, False
            elif isinstance(target_part, SlidePart):
                warnings.append(
                    f"Skipped relationship {rId} to slide {target_part.partname}."
                )
                continue
            else:
                target = _clone_part(
                    target_part,
                    dst_pkg,
                    memo,
                    source_slide_part,
                    dst_slide_part,
                    used,
                    warnings,
                )
                if target is None:
                    continue
                is_external = False

        cloned._rels._rels[rId] = _Relationship(
            cloned._rels._base_uri,
            rId,
            rel.reltype,
            RTM.EXTERNAL if is_external else RTM.INTERNAL,
            target,
        )

    return cloned


def _copy_slide_into(dst_slide, source_slide, warnings: list[str]) -> None:
    """Replace ``dst_slide``'s content with ``source_slide``'s content."""
    source_part = source_slide.part
    dst_part = dst_slide.part
    dst_pkg = dst_part.package

    # Remove content relationships from the blank target slide, keeping its
    # relationship to the target matters layout.
    for rId, rel in list(dst_part.rels.items()):
        if rel.reltype != RT.SLIDE_LAYOUT:
            dst_part._rels.pop(rId)

    new_element = copy.deepcopy(source_part._element)
    rId_map: dict[str, str] = {}
    memo: dict = {}
    used = {str(part.partname) for part in dst_pkg.iter_parts()}

    for rId, rel in source_part.rels.items():
        if rel.reltype in SKIPPED_RELTYPES:
            continue

        if rel.is_external:
            new_rId = dst_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            cloned = _clone_part(
                rel.target_part,
                dst_pkg,
                memo,
                source_part,
                dst_part,
                used,
                warnings,
            )
            if cloned is None:
                continue
            new_rId = dst_part.relate_to(cloned, rel.reltype)

        rId_map[rId] = new_rId

    unmapped: set[str] = set()
    for element in new_element.iter():
        for attr, value in list(element.attrib.items()):
            if not attr.startswith(f"{{{REL_NS}}}"):
                continue
            if value in rId_map:
                element.set(attr, rId_map[value])
            else:
                unmapped.add(value)

    if unmapped:
        warnings.append(
            "Some copied slide references could not be remapped: "
            + ", ".join(sorted(unmapped))
        )

    dst_part._element = new_element


def copy_matters_slides(
    target_pptx: str | Path,
    source_pptx: bytes | str | Path,
) -> tuple[int, list[str]]:
    """Copy announcement slides from ``source_pptx`` into ``target_pptx``.

    Returns ``(copied_slide_count, warnings)``. The first source announcement
    slide replaces the generated deck's blank matters slide; any additional
    source slides are inserted immediately after it using the same target
    layout. The target file is saved in place.
    """
    warnings: list[str] = []

    source_bytes = (
        source_pptx
        if isinstance(source_pptx, bytes)
        else Path(source_pptx).read_bytes()
    )
    source_prs = Presentation(BytesIO(source_bytes))
    source_slides = find_matters_slides(source_prs)
    if not source_slides:
        return 0, [
            "No announcement/matters slide was found in last week's PPT; "
            "the generated matters slide was left blank."
        ]
    if any(_layout_name(slide) != "matters" for slide in source_slides):
        warnings.append(
            "Some copied announcement slides did not use the 'matters' layout; "
            "check their placement in the generated deck."
        )

    target_prs = Presentation(str(target_pptx))
    if (source_prs.slide_width, source_prs.slide_height) != (
        target_prs.slide_width,
        target_prs.slide_height,
    ):
        warnings.append(
            "Last week's PPT slide size differs from the generated deck; "
            "copied content may be misaligned."
        )
    target_slides = [
        slide for slide in target_prs.slides if _layout_name(slide) == "matters"
    ]
    if not target_slides:
        return 0, ["No matters slide was found in the generated PPT."]

    anchor = target_slides[0]
    _copy_slide_into(anchor, source_slides[0], warnings)

    sldIdLst = target_prs.slides._sldIdLst
    for source_slide in source_slides[1:]:
        new_slide = target_prs.slides.add_slide(anchor.slide_layout)
        _copy_slide_into(new_slide, source_slide, warnings)

        anchor_index = list(target_prs.slides).index(anchor)
        new_sldId = list(sldIdLst)[-1]
        sldIdLst.remove(new_sldId)
        sldIdLst.insert(anchor_index + 1, new_sldId)
        anchor = new_slide

    target_prs.save(str(target_pptx))
    return len(source_slides), warnings
