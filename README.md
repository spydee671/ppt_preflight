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
```

## What it checks

| Section | Details |
|---|---|
| **Deck overview** | Slide count, hidden slides, dimensions, aspect ratio, file size, show mode (presenter / kiosk / window), loop setting, sections |
| **Fonts** | Every explicit font name with per-slide locations; theme heading/body fonts from slide master; whether fonts are embedded in the file |
| **Video** | Each video file — embedded vs. linked, format/extension, file size; codec warning for less-common formats |
| **Audio** | Same as video |
| **Images** | Format breakdown and total embedded image size; linked images flagged |
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

VIDEO  (1)
────────────────────────────────────────────────────────────
  slide  3  LINKED ⚠      .mp4    event_intro.mp4
           ↳ linked path: C:/Users/Presenter/Videos/event_intro.mp4

AUDIO  (1)
────────────────────────────────────────────────────────────
  slide  3  LINKED ⚠      .mp3    background.mp3
           ↳ linked path: C:/Users/Presenter/Music/background.mp3

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
  ✗  FAIL   1 linked video(s) — will break on a different machine: event_intro.mp4
  ✗  FAIL   1 linked audio file(s) — will break on a different machine: background.mp3
  ⚠  WARN   Fonts not embedded — must be installed on playback machine: Arial, Calibri ...
  ⚠  WARN   1 slide(s) have animations (slides [4]) — test playback on event system
  ⚠  WARN   Auto-advance timing on slides [3] — verify on event system
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

The generated deck includes all the features the tool checks: 8 fonts, a linked video, linked audio, an embedded image, three transition types, a click-triggered animation, an external hyperlink, speaker notes, and a hidden slide.

## Roadmap

### Video autoplay detection
Determine whether a video is set to play automatically on slide entry or only on click. Autoplay videos have a `delay="0"` trigger tied to the slide timeline in the `p:timing` XML, rather than a click trigger. Important for unattended kiosk decks and timed sequences.

### Video loop / mute settings
Read per-video playback flags (loop, mute, rewind after playing) from the `p14:media` extension element on the video shape's `nvPr`. Useful for checking background loop videos common in conference staging.

### Image resolution check
Use [Pillow](https://pillow.readthedocs.io) to read each embedded image's pixel dimensions, then divide by the shape's physical size in inches to compute effective PPI. Flag anything below ~96 PPI as potentially blurry on a large screen. Requires `pip install Pillow`.

### Video codec / container report
Pipe embedded video files through `ffprobe` (part of [ffmpeg](https://ffmpeg.org)) to report codec, resolution, frame rate, and bitrate. Useful for confirming H.264/AAC in an MP4 container — the most broadly supported format on event playback systems.

### Aspect ratio vs display mismatch
Accept a `--display WxH` argument (e.g. `--display 1920x1080`) and warn when the deck's aspect ratio doesn't match the target display. A 4:3 deck on a 16:9 screen (or vice versa) will show black bars or stretching unless the AV operator scales it correctly.

### SmartArt detection
Flag slides containing SmartArt (`dgm:relIds` elements). SmartArt can render differently across Office versions and may not display at all in some presentation environments. Better to know in advance and have a rasterised fallback ready.

### Per-slide thumbnail sheet
Render all slides to a contact-sheet image using LibreOffice + `pdftoppm` for a fast visual scan during tech check — one image showing the whole deck at a glance.

## Requirements

- Python 3.8+
- [python-pptx](https://python-pptx.readthedocs.io) 0.6+

## Licence

MIT
