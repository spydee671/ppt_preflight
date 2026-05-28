#!/usr/bin/env python3
"""
AV Preflight Tool for PowerPoint presentations.

Analyzes a .pptx file and reports on everything that matters before
taking it to a live event: fonts, embedded vs linked media, transitions,
animations, hidden slides, hyperlinks, and embedded objects.

Usage:
    python av_preflight.py presentation.pptx
    python av_preflight.py presentation.pptx --display 3840x2160

The --display flag defaults to 1920x1080. The deck's aspect ratio is
compared against the target display and a mismatch is flagged.

Requires:
    pip install python-pptx
"""

import argparse
import sys
import zipfile
from pathlib import Path, PurePosixPath
from collections import defaultdict

try:
    from pptx import Presentation
    from pptx.util import Emu
    from pptx.oxml.ns import qn
except ImportError:
    print("python-pptx not installed. Run: pip install python-pptx")
    sys.exit(1)


# ── Constants ─────────────────────────────────────────────────────────────────

EMU_PER_INCH = 914400

VIDEO_EXTENSIONS  = {'.mp4', '.mov', '.avi', '.wmv', '.mpg', '.mpeg',
                     '.m4v', '.mkv', '.webm', '.flv'}
AUDIO_EXTENSIONS  = {'.mp3', '.wav', '.aac', '.m4a', '.ogg', '.wma', '.flac'}
IMAGE_EXTENSIONS  = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff',
                     '.tif', '.emf', '.wmf', '.svg', '.ico'}

SAFE_VIDEO_FORMATS = {'.mp4', '.mov'}  # broadly supported on event systems
SAFE_AUDIO_FORMATS = {'.mp3', '.wav', '.aac', '.m4a'}


# ── Formatting helpers ────────────────────────────────────────────────────────

def fmt_size(n):
    if n is None:
        return "?"
    for unit in ('B', 'KB', 'MB', 'GB'):
        if n < 1024:
            return f"{n:.0f} {unit}" if unit == 'B' else f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} TB"


def aspect_label(w_emu, h_emu):
    r = w_emu / h_emu
    for val, name in ((16/9, "16:9"), (4/3, "4:3"), (16/10, "16:10"), (1.0, "1:1")):
        if abs(r - val) < 0.02:
            return name
    return f"custom ({r:.3f}:1)"


def parse_display(value):
    """Parse 'WxH' or 'W×H' into (width, height) integers."""
    for sep in ('x', '×', 'X'):
        if sep in value:
            w, h = value.split(sep, 1)
            return int(w.strip()), int(h.strip())
    raise argparse.ArgumentTypeError(f"expected WxH format (e.g. 1920x1080), got: {value!r}")


def aspect_match(deck_w_emu, deck_h_emu, disp_w, disp_h, tol=0.02):
    """Return (match, deck_ratio, display_ratio)."""
    deck_r = deck_w_emu / deck_h_emu
    disp_r = disp_w / disp_h
    return abs(deck_r - disp_r) < tol, deck_r, disp_r


def section_header(title):
    print(f"\n{title}")
    print("─" * 60)


# ── Font collection ───────────────────────────────────────────────────────────

def collect_fonts(element, slide_num, bucket):
    """
    Walk element tree collecting explicit font names from a:latin/@typeface.
    Theme font placeholders (+mj-lt, +mn-lt) are skipped here and reported
    separately via theme_font_names().
    """
    for latin in element.iter(qn('a:latin')):
        tf = latin.get('typeface', '')
        if tf and not tf.startswith('+'):
            bucket[tf].add(slide_num)


def theme_font_names(prs):
    """Return (major_font, minor_font) from the slide master's theme."""
    names = {}
    try:
        master = prs.slide_master
        theme_part = master.part.part_related_by(
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'
        )
        ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        fs = theme_part.element.find(f'{{{ns}}}themeElements/{{{ns}}}fontScheme')
        if fs is not None:
            for role in ('majorFont', 'minorFont'):
                node = fs.find(f'{{{ns}}}{role}')
                if node is not None:
                    latin = node.find(f'{{{ns}}}latin')
                    if latin is not None:
                        names[role] = latin.get('typeface', '')
    except Exception:
        pass
    return names


# ── Media extraction ──────────────────────────────────────────────────────────

def _embedded_size(rel, zf):
    try:
        zip_path = str(rel.target_part.partname).lstrip('/')
        return zf.getinfo(zip_path).file_size
    except Exception:
        return None


def slide_media(slide, slide_num, zf):
    """Return list of video/audio items from a slide's relationships."""
    items = []
    for rel in slide.part.rels.values():
        rt = rel.reltype.lower()
        target = rel.target_ref
        ext = PurePosixPath(target).suffix.lower()

        if 'video' in rt or ext in VIDEO_EXTENSIONS:
            kind = 'video'
        elif 'audio' in rt or ext in AUDIO_EXTENSIONS:
            kind = 'audio'
        else:
            continue

        embedded = not rel.is_external
        items.append(dict(
            slide=slide_num,
            kind=kind,
            name=PurePosixPath(target).name,
            ext=ext,
            embedded=embedded,
            linked_path=target if rel.is_external else None,
            size=_embedded_size(rel, zf) if embedded else None,
        ))
    return items


def slide_images(slide, slide_num, zf):
    items = []
    for rel in slide.part.rels.values():
        if 'image' not in rel.reltype.lower():
            continue
        target = rel.target_ref
        ext = PurePosixPath(target).suffix.lower()
        embedded = not rel.is_external
        items.append(dict(
            slide=slide_num,
            name=PurePosixPath(target).name,
            ext=ext,
            embedded=embedded,
            linked_path=target if rel.is_external else None,
            size=_embedded_size(rel, zf) if embedded else None,
        ))
    return items


# ── Transitions & animations ──────────────────────────────────────────────────

def slide_transition(slide, n):
    t = slide.element.find('.//' + qn('p:transition'))
    if t is None:
        return None
    children = [c.tag.split('}')[-1] for c in t if '}' in c.tag]
    return dict(
        slide=n,
        kind=children[0] if children else 'none',
        click=t.get('advClick', '1') != '0',
        auto_ms=t.get('advTm'),
        dur_ms=t.get('dur'),
    )


def slide_animation_count(slide):
    """Return count of animation effect nodes on a slide (0 = no animations)."""
    timing = slide.element.find(qn('p:timing'))
    if timing is None:
        return 0
    anim_tags = [qn(t) for t in
                 ('p:anim', 'p:animEffect', 'p:animMotion', 'p:animScale', 'p:animRot')]
    return sum(len(timing.findall('.//' + tag)) for tag in anim_tags)


# ── Presentation-level properties ─────────────────────────────────────────────

def show_properties(prs):
    """Extract playback mode and loop setting from p:showPr."""
    props = {'mode': 'standard', 'loop': False, 'no_narration': False}
    showPr = prs.element.find(qn('p:showPr'))
    if showPr is None:
        return props
    props['loop'] = showPr.get('loop', '0') == '1'
    props['no_narration'] = showPr.get('showNarration', '1') == '0'
    if showPr.find(qn('p:kiosk')) is not None:
        props['mode'] = 'kiosk (auto-advance, no menu)'
    elif showPr.find(qn('p:browse')) is not None:
        props['mode'] = 'browsed in window'
    else:
        props['mode'] = 'presented by speaker'
    return props


def presentation_sections(prs):
    """Return list of section names if the deck uses sections."""
    sections = []
    sectionLst = prs.element.find('.//' + qn('p:sectionLst'))
    if sectionLst is not None:
        for sec in sectionLst.findall(qn('p:section')):
            name = sec.get('name', '(unnamed)')
            sections.append(name)
    return sections


# ── Main ──────────────────────────────────────────────────────────────────────

def analyze(path, display_wh=(1920, 1080)):
    path = Path(path)
    if not path.exists():
        print(f"File not found: {path}")
        sys.exit(1)

    prs = Presentation(str(path))
    zf  = zipfile.ZipFile(str(path))

    print(f"\n{'━'*60}")
    print(f"  AV PREFLIGHT — {path.name}")
    print(f"{'━'*60}")

    # ── Deck overview ──────────────────────────────────────────────────────────
    n_slides = len(prs.slides)
    w_emu, h_emu = prs.slide_width, prs.slide_height
    w_in, h_in   = w_emu / EMU_PER_INCH, h_emu / EMU_PER_INCH
    w_px, h_px   = round(w_in * 96), round(h_in * 96)

    hidden_slides = [i+1 for i, s in enumerate(prs.slides)
                     if s.element.get('show') == '0']
    show_props = show_properties(prs)
    sections   = presentation_sections(prs)

    section_header("DECK OVERVIEW")
    print(f"  Slides          {n_slides}" +
          (f"  ({len(hidden_slides)} hidden: slides {hidden_slides})" if hidden_slides else ""))
    print(f"  Dimensions      {w_in:.2f}\" × {h_in:.2f}\"  "
          f"({w_px} × {h_px} px @ 96 dpi)")
    print(f"  Aspect ratio    {aspect_label(w_emu, h_emu)}")

    disp_w, disp_h = display_wh
    ratio_ok, deck_r, disp_r = aspect_match(w_emu, h_emu, disp_w, disp_h)
    disp_label = f"{disp_w}×{disp_h}  ({aspect_label(disp_w, disp_h)})"
    if ratio_ok:
        match_str = "✓ matches"
    else:
        # Describe the visual consequence
        if deck_r < disp_r:
            consequence = "pillarboxed (black bars left/right)"
        else:
            consequence = "letterboxed (black bars top/bottom)"
        match_str = f"✗ MISMATCH — {consequence}"
    print(f"  Target display  {disp_label}")
    print(f"  Ratio match     {match_str}")

    print(f"  File size       {fmt_size(path.stat().st_size)}")
    print(f"  Show mode       {show_props['mode']}")
    if show_props['loop']:
        print(f"  Loop            YES")
    if sections:
        print(f"  Sections        {len(sections)}: {', '.join(sections)}")

    # ── Fonts ──────────────────────────────────────────────────────────────────
    fonts = defaultdict(set)
    collect_fonts(prs.slide_master.element, 0, fonts)
    for i, slide in enumerate(prs.slides, 1):
        collect_fonts(slide.element, i, fonts)

    theme_fonts  = theme_font_names(prs)
    # Embedded font files stored inside the .pptx zip
    embedded_font_files = [f for f in zf.namelist() if f.startswith('ppt/fonts/')]

    section_header(f"FONTS  ({len(fonts)} explicit)")
    for name, slides in sorted(fonts.items()):
        on_slides  = sorted(s for s in slides if s > 0)
        tags       = []
        if 0 in slides:
            tags.append("master")
        if name == theme_fonts.get('majorFont'):
            tags.append("theme heading")
        if name == theme_fonts.get('minorFont'):
            tags.append("theme body")
        tag_str = f"  [{', '.join(tags)}]" if tags else ""
        print(f"  {name:<40} slides {on_slides}{tag_str}")

    if theme_fonts:
        print(f"\n  Theme heading font  {theme_fonts.get('majorFont', '?')}")
        print(f"  Theme body font     {theme_fonts.get('minorFont', '?')}")

    if embedded_font_files:
        print(f"\n  Embedded font files ({len(embedded_font_files)}):")
        for f in embedded_font_files:
            print(f"    {Path(f).name}  ({fmt_size(zf.getinfo(f).file_size)})")
    else:
        print(f"\n  ⚠  No embedded fonts — all fonts must be installed on "
              f"the playback machine")

    # ── Video ──────────────────────────────────────────────────────────────────
    all_media  = []
    all_images = []
    for i, slide in enumerate(prs.slides, 1):
        all_media.extend(slide_media(slide, i, zf))
        all_images.extend(slide_images(slide, i, zf))

    videos = [m for m in all_media if m['kind'] == 'video']
    audios  = [m for m in all_media if m['kind'] == 'audio']

    def print_media_section(label, items):
        section_header(f"{label}  ({len(items)})")
        if not items:
            print("  (none)")
            return
        for m in items:
            status = "embedded" if m['embedded'] else "LINKED ⚠"
            sz     = f"  [{fmt_size(m['size'])}]" if m['size'] else ""
            fmt_warn = ""
            if m['kind'] == 'video' and m['ext'] not in SAFE_VIDEO_FORMATS:
                fmt_warn = "  ⚠ format may need codec"
            if m['kind'] == 'audio' and m['ext'] not in SAFE_AUDIO_FORMATS:
                fmt_warn = "  ⚠ check codec support"
            print(f"  slide {m['slide']:>2}  {status:<12}  "
                  f"{m['ext']:<6}  {m['name']}{sz}{fmt_warn}")
            if not m['embedded']:
                print(f"           ↳ linked path: {m['linked_path']}")

    print_media_section("VIDEO", videos)
    print_media_section("AUDIO", audios)

    # ── Images ─────────────────────────────────────────────────────────────────
    # Deduplicate by filename (same asset used on multiple slides)
    img_map = {}
    for img in all_images:
        k = img['name']
        if k not in img_map:
            img_map[k] = {**img, 'slides': [img['slide']]}
        else:
            img_map[k]['slides'].append(img['slide'])

    ext_counts     = defaultdict(int)
    total_img_size = 0
    for img in img_map.values():
        ext_counts[img['ext']] += 1
        if img['size']:
            total_img_size += img['size']

    linked_imgs = [img for img in img_map.values() if not img['embedded']]

    section_header(f"IMAGES  ({len(img_map)} unique)")
    for ext, count in sorted(ext_counts.items()):
        print(f"  {(ext or '(none)'):<10}  {count} file{'s' if count > 1 else ''}")
    if total_img_size:
        print(f"\n  Total image data   {fmt_size(total_img_size)}")
    if linked_imgs:
        print(f"\n  ⚠  {len(linked_imgs)} linked image(s) — must travel with the deck:")
        for img in linked_imgs:
            print(f"     slides {img['slides']}  {img['name']}")

    # ── Transitions & animations ───────────────────────────────────────────────
    transitions  = []
    anim_slides  = []
    auto_advance = []

    for i, slide in enumerate(prs.slides, 1):
        t = slide_transition(slide, i)
        if t:
            transitions.append(t)
            if t['auto_ms']:
                auto_advance.append(i)
        count = slide_animation_count(slide)
        if count > 0:
            anim_slides.append((i, count))

    section_header("TRANSITIONS & ANIMATIONS")
    print(f"  Slides with transitions    {len(transitions)}/{n_slides}")
    for t in transitions:
        parts = [f"slide {t['slide']:>2}", f"{t['kind']:<20}"]
        if t['click']:
            parts.append("advance on click")
        if t['auto_ms']:
            parts.append(f"auto after {t['auto_ms']} ms")
        print(f"  {'  '.join(parts)}")

    print(f"\n  Slides with animations     {len(anim_slides)}/{n_slides}")
    for slide_n, count in anim_slides:
        print(f"  slide {slide_n:>2}  {count} animation effect(s)")

    if auto_advance:
        print(f"\n  ⚠  Auto-advancing slides: {auto_advance}  — verify timing on event system")

    # ── Speaker notes ──────────────────────────────────────────────────────────
    notes_slides = [
        i+1 for i, s in enumerate(prs.slides)
        if s.has_notes_slide and s.notes_slide.notes_text_frame.text.strip()
    ]
    section_header("SPEAKER NOTES")
    print(f"  {len(notes_slides)}/{n_slides} slides have notes" +
          (f"  (slides {notes_slides})" if notes_slides else ""))

    # ── Hyperlinks ─────────────────────────────────────────────────────────────
    links = []
    for i, slide in enumerate(prs.slides, 1):
        for rel in slide.part.rels.values():
            if 'hyperlink' in rel.reltype.lower():
                links.append((i, rel.target_ref, rel.is_external))

    section_header(f"HYPERLINKS  ({len(links)})")
    if links:
        for slide_n, target, external in links:
            kind = "external" if external else "internal"
            print(f"  slide {slide_n:>2}  {kind:<10}  {target}")
    else:
        print("  (none)")

    # ── Embedded OLE objects ───────────────────────────────────────────────────
    ole_files = [f for f in zf.namelist() if '/embeddings/' in f]
    if ole_files:
        section_header(f"EMBEDDED OBJECTS / OLE  ({len(ole_files)})")
        for f in ole_files:
            print(f"  {Path(f).name}  ({fmt_size(zf.getinfo(f).file_size)})")

    # ── Charts ─────────────────────────────────────────────────────────────────
    chart_files = [f for f in zf.namelist() if f.startswith('ppt/charts/chart')]
    if chart_files:
        section_header(f"CHARTS  ({len(chart_files)})")
        print(f"  {len(chart_files)} chart(s) — verify data labels render correctly on event system")

    # ── Preflight summary ──────────────────────────────────────────────────────
    issues   = []
    warnings = []

    linked_v = [m for m in videos if not m['embedded']]
    linked_a = [m for m in audios if not m['embedded']]

    if linked_v:
        issues.append(
            f"{len(linked_v)} linked video(s) — will break on a different machine: "
            f"{', '.join(m['name'] for m in linked_v)}"
        )
    if linked_a:
        issues.append(
            f"{len(linked_a)} linked audio file(s) — will break on a different machine: "
            f"{', '.join(m['name'] for m in linked_a)}"
        )
    if linked_imgs:
        issues.append(
            f"{len(linked_imgs)} linked image(s) — may not display on a different machine"
        )

    unsafe_vid = [m for m in videos if m['embedded'] and m['ext'] not in SAFE_VIDEO_FORMATS]
    if unsafe_vid:
        fmts = ', '.join(set(m['ext'] for m in unsafe_vid))
        warnings.append(f"Video format(s) {fmts} may need a codec — prefer .mp4/.mov")

    if not embedded_font_files and fonts:
        font_list = ', '.join(sorted(fonts)[:6])
        if len(fonts) > 6:
            font_list += f" (+{len(fonts)-6} more)"
        warnings.append(f"Fonts not embedded — must be installed on playback machine: {font_list}")

    if anim_slides:
        warnings.append(
            f"{len(anim_slides)} slide(s) have animations "
            f"(slides {[s for s,_ in anim_slides]}) — test playback on event system"
        )
    if auto_advance:
        warnings.append(f"Auto-advance timing on slides {auto_advance} — verify on event system")
    if ole_files:
        warnings.append(f"{len(ole_files)} OLE embedded object(s) — requires matching software to render")
    if hidden_slides:
        warnings.append(f"{len(hidden_slides)} hidden slide(s): {hidden_slides}")
    if show_props['mode'] == 'kiosk (auto-advance, no menu)':
        warnings.append("Kiosk mode — presenter cannot skip slides manually")

    ext_links = [l for l in links if l[2]]
    if ext_links:
        warnings.append(
            f"{len(ext_links)} external hyperlink(s) — confirm internet access on event system"
        )

    if not ratio_ok:
        warnings.append(
            f"Aspect ratio mismatch — deck is {aspect_label(w_emu, h_emu)}, "
            f"display is {disp_w}×{disp_h} ({aspect_label(disp_w, disp_h)}) — "
            f"will be {consequence}"
        )

    print(f"\n{'━'*60}")
    print("PREFLIGHT SUMMARY")
    print("─" * 60)
    if not issues and not warnings:
        print("  ✓  All checks passed — good to go")
    else:
        for item in issues:
            print(f"  ✗  FAIL   {item}")
        for item in warnings:
            print(f"  ⚠  WARN   {item}")
    print(f"{'━'*60}\n")

    zf.close()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description="AV preflight check for PowerPoint presentations.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument('file', help='Path to the .pptx file')
    parser.add_argument(
        '--display',
        metavar='WxH',
        type=parse_display,
        default=(1920, 1080),
        help='Target display resolution (default: 1920x1080)',
    )
    args = parser.parse_args()
    analyze(args.file, display_wh=args.display)
