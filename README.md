# MCG Create Dummy Detail/Section

Inventor 2023 add-in for placing **dummy section markers** and **detail boundaries** on drawings without generating real Section/Detail views — keeps exported AutoCAD DWG files clean.

---

## Features

| Task | Status | Description |
|------|--------|-------------|
| 1. Smart Replace Model | 🟡 Pending | Shelved — Inventor proxy limitations |
| 2. Smart Rotate View | 🟡 Pending | Shelved — Inventor proxy limitations |
| 3. **Create Dummy Section** | ✅ Active | Section markers (4 directional symbols) with auto-stub + dimension |
| 4. **Create Dummy Detail** | ✅ Active | Detail boundary (circle/rectangle) + connection line + ISOCPEUR label |

---

## Project Structure

```
InventorDrawingPlugin/
├── Core/                             # MCG portable ribbon SDK (shared with other MCG addins)
│   ├── MCGRibbonManager.cs           # Ribbon tab + panels + dockable window manager
│   ├── DockableWindowHost.cs         # WPF embedding into Inventor DockableWindow
│   ├── IToolDescriptor.cs            # Tool descriptor interface
│   ├── IDockablePanelDescriptor.cs   # Palette descriptor interface
│   ├── PictureDispConverter.cs       # Bitmap ↔ IPictureDisp helper
│   ├── PanelLocation.cs              # Model | Drawing | Utility
│   └── RibbonContext.cs              # Part | Assembly | Drawing flags
├── Services/
│   ├── DummyToolsDescriptor.cs       # IToolDescriptor for "Dummy Tools" button
│   ├── DummyToolsPanelDescriptor.cs  # IDockablePanelDescriptor for palette
│   ├── DummySectionService.cs        # Task 3 logic
│   ├── DummyDetailService.cs         # Task 4 logic
│   ├── SmartReplaceService.cs        # Task 1 (shelved, kept for reference)
│   ├── SmartRotateService.cs         # Task 2 (shelved, kept for reference)
│   └── SketchConstraintService.cs    # Sketch constraint helpers
├── Views/
│   ├── MainPalette.xaml              # WPF palette UI
│   └── MainPalette.xaml.cs
├── Resources/
│   ├── icondata16.png                # Ribbon button icon (16×16)
│   └── icondata32.png                # Ribbon button icon (32×32)
├── StandardAddInServer.cs            # Add-in entry point (JIT-safe pattern)
├── MCG_InventorCreateDummyDetailSection.addin   # Inventor addin descriptor
├── Install_AutoLoadInventorAddin.bat            # Installer script
├── Symbol Sections.idw                           # Section symbol library (4 directional symbols)
├── Instructions.html                             # User-facing usage guide
└── InventorPlugin.csproj
```

---

## Build

```bash
dotnet build
```

Output: `bin/Debug/MCG_InventorCreateDummyDetailSection.dll` + `MCG_InventorCreateDummyDetailSection.addin`

---

## Deployment Layout

The `.bat` script expects subfolder layout per addin. Place the entire addin in its own `MCG_<AddinName>` subfolder under `C:\CustomTools\Inventor\`:

```
C:\CustomTools\Inventor\
├── MCG_InventorCreateDummyDetailSection\        ← addin subfolder
│   ├── MCG_InventorCreateDummyDetailSection.dll
│   ├── MCG_InventorCreateDummyDetailSection.addin
│   └── Symbol Sections.idw                       ← stays here, addin finds via fallback
└── Install_AutoLoadInventorAddin.bat             ← run once
```

Run `Install_AutoLoadInventorAddin.bat` → copies `.dll + .addin` to `%APPDATA%\Autodesk\Inventor 2023\Addins\MCG_InventorCreateDummyDetailSection\`.

`Symbol Sections.idw` is **not** copied by the .bat — the addin reads it from the source location (`C:\CustomTools\Inventor\MCG_InventorCreateDummyDetailSection\`) when needed.

---

## Ribbon Integration

Uses the `MCG.Inventor.Ribbon` SDK (Core folder). Tool appears at:

```
Drawing ribbon
└── MCG TOOLS tab               (shared with other MCG addins)
    └── Drawing panel
        └── [Dummy Tools]       click → toggle palette dock
```

The MCG TOOLS tab is created by whichever MCG addin loads first; subsequent addins reuse it.

---

## Symbol Library

Task 3 (Dummy Section) requires 4 SketchedSymbol definitions:
`SECTION_L-R`, `SECTION_R-L`, `SECTION_U-D`, `SECTION_D-U`

If missing in the active drawing, the addin opens `Symbol Sections.idw` (visible) and prompts the user to copy symbols via Inventor UI (right-click → Copy/Paste in Browser → Sketched Symbols).

Manual copy is required because the Inventor API's `CopyContentsTo` does not transfer solid hatch fill on arrow heads.

---

## Add-in Architecture

- **Entry point**: `StandardAddInServer.cs` uses JIT-safe pattern (`Activate` → `ActivateInternal` with `[MethodImpl(NoInlining)]`) so dependency-loading exceptions don't propagate to Inventor and break Vault/Content Center.
- **Logger**: `C:\CustomTools\Inventor\logs\MCG_InventorCreateDummyDetailSection.log` — auto-deleted on successful load. Presence indicates an error.
- **GUID**: `{9D0B4403-518D-40C6-98AC-CD2FF878B309}` — must be unique across all MCG addins on a machine.

---

## Documentation

- **End users**: open `Instructions.html` for installation + usage guide.
- **Developers**: see `Core/MCGRibbonReadme.md` for the portable ribbon SDK conventions.
