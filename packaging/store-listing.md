# Microsoft Store listing content (draft)

## Title
Print Spooler Queue Cleanup

## Short description (≤ 150 chars, EN)
View, pause, purge and repair stuck Windows print queues safely — with backups, tray
counter, dark mode and 30 languages. One tiny native exe.

## Description (EN)
When a print job wedges, the whole queue stalls: the service says "Running", nothing
prints, and cancelling does nothing. Print Spooler Queue Cleanup shows every job on every
printer, lets you remove just the bad ones — or clear everything through a careful,
reversible path: optional backup of spool files, documented API purge first, and only then
a guarded Print Spooler restart with a surgical file sweep that touches only *.SPL/*.SHD.

Features
• Live queue view: printers grouped, document, owner, size, status, submitted time, auto-refresh
• Selective delete, clear-all with confirmation, pause/resume queues, restart the service
• Automatic spool-file backups before bulk purge and one-click restore
• Tray guardian with a live job counter and one-click purge; double-click opens the app
• Light / Dark / Follow-system themes with instant reaction to Windows theme changes
• 30 interface languages, right-to-left layouts for Arabic, Hebrew and Persian
• Built-in installer (per-user or all-users), portable mode, silent CLI for IT automation
• Self-update from GitHub Releases with in-place replace
• Zero dependencies: no .NET, no VC++ runtime, no PowerShell — pure native Win32

Administrator rights are requested only for operations that truly need them (other users'
jobs, service restart, all-users install); everything else runs as a standard user.

License: PolyForm Noncommercial 1.0.0 (free for non-commercial use).
Author: Mikhail Deynekin — https://Deynekin.com

## Description (RU, сокращённо)
Когда задание печати зависает, встаёт вся очередь: служба «выполняется», ничего не
печатается. Утилита показывает все задания на всех принтерах, позволяет удалить только
проблемные или очистить всё безопасно: резервная копия файлов спула, штатный purge API и
лишь затем аккуратный перезапуск службы с удалением только *.SPL/*.SHD. Трей со счётчиком
заданий, светлая/тёмная тема, 30 языков, встроенный установщик и портативный режим,
автообновление. Без .NET и сторонних рантаймов. Права администратора запрашиваются только
для операций, которые действительно их требуют.

## Keywords
print spooler, print queue, stuck print job, clear print queue, spooler repair, printer
maintenance, purge print jobs

## Category
Productivity ▸ Utilities

## Permissions justification note (certification)
Full-trust desktop utility for print-queue maintenance. Elevation (UAC) is requested only
for explicit admin operations: cross-user job deletion, Print Spooler service restart,
all-users installation. No network access except HTTPS to api.github.com / github.com for
the optional update check. No personal data is collected or transmitted.
