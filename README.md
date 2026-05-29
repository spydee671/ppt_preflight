# ppt_preflight

A command-line AV preflight tool for PowerPoint presentations. Analyses a `.pptx` file and produces a structured report covering everything that matters before taking a deck to a live event.

## Installation

```bash
pip install python-pptx
```

Clone or download `av_preflight.py` — no other dependencies required.

## Usage

```bash
python av_preflight.py presentation.pptx
python av_preflight.py presentation.pptx --display 3840x2160
```

The `--display` flag sets the target display resolution for the aspect ratio check. Defaults to `1920x1080`.

## What it checks

| Section | Details |
|---|---|
| **Deck overview** | Slide count, hidden slides, dimensions, aspect ratio, file size, show mode (presenter / kiosk / window), loop setting, sections; aspect ratio compared against target display (default 1920×1080) with letterbox/pillarbox warning on mismatch |
| **Fonts** | Every explicit font name with per-slide locations; theme heading/body fonts from slide master; whether fonts are embedded in the file |
| **Video** | Table per video — embedded vs. linked, format/extension, **autoplay**, **loop**, **mute**, file size; codec warning for less-common formats; non-autoplay videos flagged in summary |
| **Audio** | Same as video |
| **Images** | Format breakdown and total embedded image size; linked images flagged |
| **Image resolution** | Per-image PPI check (requires Pillow): flags images below 96 PPI (blurry on screen) and above 300 PPI (wasteful — print quality, no benefit on screen); oversized images show the pixel-count ratio and an estimated ideal size to resize to |
| **Transitions** | Per-slide transition type; advance mode (click vs. auto-timer); auto-advance slides highlighted |
| **Animations** | Animation effect count per animated slide |
| **Speaker notes** | Count of slides that have notes |
| **Hyperlinks** | All hyperlinks with slide location and internal/external classification |
| **Embedded objects** | OLE objects (e.g. embedded Excel, Word) flagged with file size |
| **Charts** | Count of chart objects — reminder to verify data labels render correctly |

### Preflight summary

The report ends with a prioritised summary:

- `✗ FAIL` — hard problems that will break the deck on a different machine (linked media)
- `⚠ WARN` — items that need verifying on the event system (fonts, animations, auto-advance, kiosk mode, external links)

## Sample output

```
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  AV PREFLIGHT — show_deck.pptx
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

DECK OVERVIEW
────────────────────────────────────────────────────────────
  Slides          5  (1 hidden: slides [5])
  Dimensions      10.00" × 5.62"  (960 × 540 px @ 96 dpi)
  Aspect ratio    16:9
  File size       38.2 KB
  Show mode       standard

FONTS  (8 explicit)
────────────────────────────────────────────────────────────
  Arial                                    slides [2]
  Calibri                                  slides [1]
  Courier New                              slides [5]
  Georgia                                  slides [1]
  Gill Sans MT                             slides [2]
  Helvetica Neue                           slides [4]
  Impact                                   slides [4]
  Trebuchet MS                             slides [3]

  ⚠  No embedded fonts — all fonts must be installed on the playback machine

VIDEO  (2)
────────────────────────────────────────────────────────────
  Slide  File                            Source     Autoplay  Loop   Mute   Size
  ─────  ──────────────────────────────  ─────────  ────────  ─────  ─────  ──────────
      3  event_intro.mp4                 linked ⚠   YES       no     no     —
         ↳ C:/Videos/event_intro.mp4
      3  bg_ambient.mp4                  linked ⚠   NO ⚠      yes    yes    —
         ↳ C:/Videos/bg_ambient.mp4

  ⚠  1 video(s) not set to autoplay — will require a manual click or trigger to start

AUDIO  (1)
────────────────────────────────────────────────────────────
  slide  3  LINKED ⚠      .mp3    background.mp3
           ↳ linked path: C:/Music/background.mp3

TRANSITIONS & ANIMATIONS
────────────────────────────────────────────────────────────
  Slides with transitions    3/5
  slide  2  fade                  advance on click
  slide  3  push                  auto after 8000 ms
  slide  4  dissolve              advance on click

  Slides with animations     1/5
  slide  4  1 animation effect(s)

  ⚠  Auto-advancing slides: [3]  — verify timing on event system

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PREFLIGHT SUMMARY
────────────────────────────────────────────────────────────
  ✗  FAIL   2 linked video(s) — will break on a different machine: event_intro.mp4, bg_ambient.mp4
  ✗  FAIL   1 linked audio file(s) — will break on a different machine: background.mp3
  ⚠  WARN   1 video(s) not set to autoplay — require manual click to start: bg_ambient.mp4
  ⚠  WARN   Fonts not embedded — must be installed on playback machine: Arial, Calibri ...
  ⚠  WARN   1 slide(s) have animations (slides [4]) — test playback on event system
  ⚠  WARN   1 hidden slide(s): [5]
  ⚠  WARN   1 external hyperlink(s) — confirm internet access on event system
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

## Testing

A synthetic test deck can be generated to exercise all checks:

```bash
pip install python-pptx
python make_test_pptx.py   # writes test_deck.pptx
python av_preflight.py test_deck.pptx
```

The generated deck includes all the features the tool checks: 8 fonts, two linked videos (one autoplay/unmuted, one click-to-play/looping/muted), linked audio, two embedded images (one blurry, one oversized), three transition types, a click-triggered animation, an external hyperlink, speaker notes, and a hidden slide.

## Roadmap

### ~~Video autoplay detection~~ ✓ done
Implemented — detects whether each video shape autoplays on slide entry (`delay="0"` in the outer mainSeq timing block) or requires a manual click (`delay="indefinite"`). Non-autoplay videos are flagged with `NO ⚠` in the video table and summarised in the preflight report.

### ~~Video loop / mute settings~~ ✓ done
Implemented — reads per-video `repeatCount="indefinite"` (loop) and `setVolume(0)` command (muted) from the slide's `p:timing` XML. Displayed as `yes`/`no` columns in the video table alongside autoplay status.

### ~~Image resolution check~~ ✓ done
Implemented — requires `pip install Pillow`. Flags images below 96 PPI (blurry) and above 300 PPI (oversized), with pixel-count ratio and suggested resize dimensions for oversized images.

### Video codec / container report
Pipe embedded video files through `ffprobe` (part of [ffmpeg](https://ffmpeg.org)) to report codec, resolution, frame rate, and bitrate. Useful for confirming H.264/AAC in an MP4 container — the most broadly supported format on event playback systems.

### ~~Aspect ratio vs display mismatch~~ ✓ done
Implemented — use `--display WxH` to override the default 1920×1080 target.

### SmartArt detection
Flag slides containing SmartArt (`dgm:relIds` elements). SmartArt can render differently across Office versions and may not display at all in some presentation environments. Better to know in advance and have a rasterised fallback ready.

### Per-slide thumbnail sheet
Render all slides to a contact-sheet image using LibreOffice + `pdftoppm` for a fast visual scan during tech check — one image showing the whole deck at a glance.

## Requirements

- Python 3.8+
- [python-pptx](https://python-pptx.readthedocs.io) 0.6+
- [Pillow](https://pillow.readthedocs.io) *(optional — required for image resolution checks)*

## Licence

MIT
