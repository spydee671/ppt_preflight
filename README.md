# ppt_preflight

A command-line AV preflight tool for PowerPoint presentations. Analyses a `.pptx` file and produces a structured report covering everything that matters before taking a deck to a live event.

## Installation

```bash
pip install python-pptx
pip install Pillow        # optional — required for image resolution checks
```

Install [FFmpeg](https://ffmpeg.org/download.html) for full video codec detection (the `ffprobe` command must be on your PATH). Without it, only fast-start MP4/MOV files are probed; all other metadata is skipped.

Clone or download `av_preflight.py` — no other Python dependencies required.

## Usage

```bash
python av_preflight.py presentation.pptx
python av_preflight.py presentation.pptx --display 3840x2160
python av_preflight.py presentation.pptx --verbose
```

| Flag | Default | Description |
|---|---|---|
| `--display WxH` | `1920x1080` | Target display resolution — used for aspect ratio check and flagging oversized video |
| `--verbose` / `-v` | off | Show full per-image resolution detail (default: summary only) |

## What it checks

| Section | Details |
|---|---|
| **Deck overview** | Slide count, hidden slides, dimensions, aspect ratio vs. target display (letterbox/pillarbox warning on mismatch), file size, show mode (presenter / kiosk / window), loop setting |
| **Document properties** | Author, last-saved-by, created/modified timestamps, revision count, PowerPoint version, OS the deck was saved on |
| **Fonts** | Every explicit font name with slide locations (range-compressed); theme fonts from master; whether any fonts are embedded |
| **Video** | Table per video — embedded vs. linked, **codec** (H.264/H.265/ProRes/etc.), **autoplay**, **loop**, **mute**, file size; detail line with resolution, frame rate, duration, audio codec/channels; flags for interlaced, HDR, no audio track, resolution exceeding target display |
| **Audio** | Embedded vs. linked, format, file size |
| **Images** | Format breakdown, total embedded image size, linked images flagged |
| **Image resolution** | Per-image PPI check (requires Pillow): flags images below 96 PPI (blurry on screen) and above 300 PPI (wasteful); compact summary by default, full per-image table with `--verbose` |
| **Transitions** | Slides with actual transitions (none/default suppressed); transition type, duration, advance mode (click vs. auto-timer) |
| **Animations** | Effect count per animated slide, grouped by duration; indefinite/slow effects flagged; master/layout animations flagged separately |
| **Speaker notes** | Count and slide list of slides with notes |
| **Hyperlinks** | All hyperlinks with slide location and internal/external classification |
| **Embedded objects** | OLE object count and total size (e.g. embedded Excel/Word); reminder that matching software must be installed on the playback machine |
| **Charts** | Count of chart objects |

### Preflight summary

The report ends with a prioritised summary — all items reference **slide numbers** so you can jump straight to the problem:

- `✗ FAIL` — hard problems that will break the deck on a different machine (linked media, ProRes/WMV video that won't play on Windows)
- `⚠ WARN` — items to verify on the event system (fonts, codec issues, animations, auto-advance, interlaced/HDR video, no-audio tracks, hidden slides, external links)

## Sample output

```
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  AV PREFLIGHT — show_deck.pptx
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

DECK OVERVIEW
────────────────────────────────────────────────────────────
  Slides          5  (1 hidden: [5])
  Dimensions      10.00" × 5.62"  (960 × 540 px @ 96 dpi)
  Aspect ratio    16:9
  Target display  1920×1080  (16:9)
  Ratio match     ✓ matches
  File size       39.3 KB
  Show mode       standard

DOCUMENT PROPERTIES
────────────────────────────────────────────────────────────
  Last saved by   Jane Smith
  Created         2026-05-01  14:22 UTC
  Modified        2026-05-28  09:45 UTC
  Revision        12
  Application     Microsoft Office PowerPoint  (Office 2016 / 2019 / 2021 / 365)  [Windows]
  Format          Widescreen

FONTS  (8 explicit)
────────────────────────────────────────────────────────────
  Arial                                    slides [2]
  Calibri                                  slides [1]
  ...
  ⚠  No embedded fonts — all fonts must be installed on the playback machine

VIDEO  (2)
────────────────────────────────────────────────────────────
  Slide  File                            Source     Codec           Autoplay  Loop   Mute   Size
  ─────  ──────────────────────────────  ─────────  ──────────────  ────────  ─────  ─────  ──────────
      3  event_intro.mp4                 embedded   H.264           YES       no     no     136.7 MB
         ↳ 1920×1080  29.97 fps  2:35  AAC stereo
      3  bg_ambient.mp4                  embedded   H.265 ⚠         NO ⚠      yes    yes    60.0 MB
         ↳ 3840×2160 ⚠ (display is 1920×1080)  23.98 fps  1:00  no audio ⚠

  ⚠  1 video(s) not set to autoplay — will require a manual click or trigger to start

IMAGE RESOLUTION  (target 192 PPI, range 96–300 PPI)
────────────────────────────────────────────────────────────
  513 image(s) checked — 342 issue(s)  (171 ok)
  ⚠  335 oversized  (>300 PPI)  on slides [2, 5, 10, 29, 34–35, ...]
       worst: 56.4× — image433.png  slide 60  (3023×3000 px → ~54×53 px ideal)
  ⚠    7 blurry     (<96 PPI)   on slides [43, 46, 55, 57–58, 85, 94]

  Add --verbose for per-image detail

TRANSITIONS & ANIMATIONS
────────────────────────────────────────────────────────────
  Transitions  (3/5 slides)
  slide   2  fade              0.7s          click
  slide   3  push              0.5s          click
  slide   4  dissolve          0.6s          click

  Animations  (1/5 slides)
  slide   4    1 effect(s)   1×0.5s

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PREFLIGHT SUMMARY
────────────────────────────────────────────────────────────
  ✗  FAIL   1 ProRes video(s) on slides [7] — Apple codec — will NOT play on Windows without QuickTime or a codec pack
  ✗  FAIL   3 linked image(s) on slides [63] — will not display on a different machine
  ⚠  WARN   1 video(s) not set to autoplay on slides [3] — require manual click to start
  ⚠  WARN   1 H.265 video(s) on slides [3] — hardware decode required — may fail on older event systems
  ⚠  WARN   1 video(s) have no audio track on slides [3] — confirm this is intentional
  ⚠  WARN   1 video(s) exceed target display resolution (1920×1080) on slides [3] — largest is 3840×2160 (slide 3)
  ⚠  WARN   Fonts not embedded — must be installed on playback machine: Arial, Calibri, ... (+N more)
  ⚠  WARN   1 slide(s) have animations [4] — test playback on event system
  ⚠  WARN   1 hidden slide(s): [5]
  ⚠  WARN   1 external hyperlink(s) on slides [4] — confirm internet access on event system
  ⚠  WARN   335 oversized image(s) on slides [...] — largest 56.4× over (slide 60) — downsample to reduce file size
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

## Testing

A synthetic test deck can be generated to exercise all checks:

```bash
python make_test_pptx.py   # writes test_deck.pptx
python av_preflight.py test_deck.pptx
python av_preflight.py test_deck.pptx --verbose
```

The generated deck includes: 8 fonts, two linked videos (one autoplay/unmuted, one click-to-play/looping/muted), linked audio, two embedded images (one blurry, one oversized), three transition types, a click-triggered animation, an external hyperlink, speaker notes, and a hidden slide.

## Requirements

| Dependency | Required | Purpose |
|---|---|---|
| Python 3.8+ | ✓ | — |
| [python-pptx](https://python-pptx.readthedocs.io) 0.6+ | ✓ | PPTX parsing |
| [Pillow](https://pillow.readthedocs.io) | optional | Image resolution checks |
| [FFmpeg](https://ffmpeg.org) (`ffprobe`) | optional | Video codec, resolution, frame rate, duration, audio track detection |

## Roadmap

### SmartArt detection
Flag slides containing SmartArt (`dgm:relIds` elements). SmartArt can render differently across Office versions and may not display correctly in some presentation environments.

### Per-slide thumbnail sheet
Render all slides to a contact-sheet image using LibreOffice + `pdftoppm` for a fast visual scan during tech check — one image showing the whole deck at a glance.

## Licence

MIT
