# Themes and languages (v5)

## Themes

Three modes: **System**, **Light**, **Dark** (toolbar cycle button, combos in main window
and settings, tray submenu — all stay in sync).

* *System* reads `HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme`
  and re-applies instantly when Windows broadcasts `WM_SETTINGCHANGE` with
  `ImmersiveColorSet` (switching the OS theme flips the app live).
* Dark mode: immersive dark title bar (`DwmSetWindowAttribute`,
  `DWMWA_USE_IMMERSIVE_DARK_MODE`), dark common controls (`DarkMode_Explorer` /
  `DarkMode_ItemsView` where supported), plus fully custom-drawn Fluent-style buttons,
  switches and status chips with the dark palette.
* Light mode mirrors the same geometry with the light palette.
* Per-Monitor V2 DPI: layouts, fonts and glyphs rescale on monitor change.

Palette (hex): light bg `#F3F3F3`, surface `#FBFBFB`, text `#1B1B1B`, accent `#0F6CBD`;
dark bg `#202020`, surface `#2B2B2B`, list `#1B1B1B`, text `#FFFFFF`, accent `#0078D4`.
Contrast ratios meet WCAG AA for text and AA-large for accent fills.

## Languages

30 languages, default = **system UI language** (`GetUserDefaultUILanguage`, with
script/region awareness: `zh-Hans` vs `zh-Hant`, `pt`, `nb`…), override in the Language
combo or `/lang=<tag>`; the choice persists.

`en, ru, es, zh-Hans, zh-Hant, ar, hi, pt, fr, de, ja, ko, it, tr, vi, id, th, nl, pl, uk,
ro, cs, sv, hu, el, he, fa, da, fi, nb`

* Right-to-left layouts for **ar, he, fa** (mirrored window layout).
* Dates/times/sizes use the user's locale formats; numbers in CLI output are invariant.
* Integrity is enforced by CI: every table must contain all 145 strings with placeholder
  parity against English (`tools/run_native_tests.sh`).
