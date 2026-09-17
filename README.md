# netserverenum

CSC VB6 domain machine enumerator (project Project1). Enter a domain name and `&NetServerEnum` calls NetAPI32 `NetServerEnum` at info level 100 with `SV_TYPE_NT`, then fills a ListView with the active NT computer names returned for that domain.

**Source last updated:** 2026-08-27 · **Language:** VB6 · **Target:** VB6 Win32 · **Output:** WinForms exe

_Note: original OneDrive LastWriteTime values were wiped to 2026-08-27 by a zip transfer; date above uses best available evidence (headers/copyright where helpful)._

## Solution structure

| Project | Language | Type | Purpose |
|---------|----------|------|---------|
| `NetServerEnum` (`Project1.vbp`) | VB6 | WinForms exe | Domain NetServerEnum ListView of NT hosts |

## How to open

Open the `.vbp` in Visual Basic 6.0 IDE:
- `Project1.vbp`

## Requirements

- Visual Basic 6.0 IDE
- Registered OCX/DLL dependencies referenced by the `.vbp` (may need to be installed separately):
  - `MSCOMCTL.OCX`

## Attribution and provenance

Working copy from Dave Robinson's OneDrive Historical Dev folder `VB/Old/netserverenum`.
Company names in project files: CSC.

## License

MIT (c) 2026 VaderConsulting for Dave Robinson's code. See `LICENSE`.
