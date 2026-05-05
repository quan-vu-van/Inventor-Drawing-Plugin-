# Core/ — MCG Inventor Ribbon SDK (portable)

Folder này là **portable SDK** chứa toàn bộ logic tạo tab "MCG TOOLS" + panels (Model/Drawing/Utility) + dockable window. Copy nguyên cả folder vào project Inventor addin mới để tool của addin đó hiện lên cùng tab "MCG TOOLS" dùng chung với các addin MCG khác.

Namespace: `MCG.Inventor.Ribbon` — không tham chiếu addin cụ thể nào.

---

## 1. Nguyên tắc share tab giữa nhiều addin

Inventor cho phép nhiều addin cùng add vào 1 ribbon tab khi dùng chung `InternalName`. `MCGRibbonManager` dùng pattern `try { tabs[id] } catch { tabs.Add(...) }` → addin load trước tạo tab, addin load sau tái sử dụng và chỉ thêm button của mình vào.

Mapping panel ↔ ribbon (cố định, mọi addin phải tuân theo):

| Panel | Xuất hiện trong ribbon |
|-------|------------------------|
| `Model` | Part + Assembly |
| `Drawing` | Drawing |
| `Utility` | Part + Assembly + Drawing |

---

## 2. Sao chép vào project mới — checklist 5 bước

### Bước 1. Copy folder `Core/` nguyên vẹn

Paste vào project mới. **Giữ nguyên** tên folder `Core` và namespace `MCG.Inventor.Ribbon`. Cùng đem theo file `MCGRibbonReadme.md` này.

### Bước 2. Cấu hình `.csproj`

```xml
<PropertyGroup>
  <TargetFramework>net48</TargetFramework>
  <UseWindowsForms>true</UseWindowsForms>     <!-- Timer, NativeWindow -->
  <UseWPF>true</UseWPF>                       <!-- HwndSource, UserControl -->
  <PlatformTarget>x64</PlatformTarget>
  <RegisterForComInterop Condition="'$(MSBuildRuntimeType)'=='Full'">true</RegisterForComInterop>
</PropertyGroup>

<ItemGroup>
  <Reference Include="Autodesk.Inventor.Interop">
    <HintPath>C:\Program Files\Autodesk\Inventor 2023\Bin\Public Assemblies\Autodesk.Inventor.Interop.dll</HintPath>
    <EmbedInteropTypes>False</EmbedInteropTypes>
    <Private>False</Private>
  </Reference>
  <Reference Include="stdole">
    <HintPath>C:\Program Files\Autodesk\Inventor 2023\Bin\stdole.dll</HintPath>
    <EmbedInteropTypes>False</EmbedInteropTypes>
    <Private>False</Private>
  </Reference>
</ItemGroup>
```

### Bước 3. Implement `IToolDescriptor` cho mỗi tool

```csharp
using System.Drawing;
using System.Reflection;
using Inventor;
using MCG.Inventor.Ribbon;

internal class MyToolDescriptor : IToolDescriptor
{
    public string Id          => "id.Button.MyAddin.MyTool";   // ← unique toàn hệ thống
    public string DisplayName => "My\nTool";
    public string Tooltip     => "My Tool";
    public string Description => "Làm việc XYZ";

    public Bitmap Icon16 => LoadIcon("MyTool_16.png");
    public Bitmap Icon32 => LoadIcon("MyTool_32.png");

    public PanelLocation Panel    => PanelLocation.Utility;     // Model | Drawing | Utility
    public RibbonContext Contexts => RibbonContext.Part | RibbonContext.Assembly;

    public ButtonDisplayEnum ButtonDisplay => ButtonDisplayEnum.kAlwaysDisplayText;
    public CommandTypesEnum  CommandType   => CommandTypesEnum.kNonShapeEditCmdType;
    public bool              UseLargeIcon  => true;

    public void OnExecute(NameValueMap context)
    {
        // Logic khi click button (nếu không có palette).
        // Nếu có palette, MCGRibbonManager auto-toggle, không cần xử lý ở đây.
    }

    public IDockablePanelDescriptor DockablePanel => null;   // hoặc provide nếu có palette

    // ⚠ QUAN TRỌNG: lấy root namespace từ TYPE, không dùng asm.GetName().Name
    // Lý do: AssemblyName có thể khác RootNamespace (ví dụ rename assembly
    // mà giữ nguyên RootNamespace). Embedded resource luôn dùng RootNamespace.
    private static Bitmap LoadIcon(string fileName)
    {
        var asm = Assembly.GetExecutingAssembly();
        string rootNs = typeof(MyToolDescriptor).Namespace.Split('.')[0];
        string resourceName = $"{rootNs}.Resources.MyTool.{fileName}";
        return PictureDispConverter.LoadBitmapFromResource(asm, resourceName);
    }
}
```

### Bước 4. Entry point addin — pattern JIT-safe (BẮT BUỘC)

```csharp
using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Inventor;
using MCG.Inventor.Ribbon;

[Guid("XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX")]   // ← GUID DUY NHẤT, không trùng addin khác
[ClassInterface(ClassInterfaceType.None)]
[ComVisible(true)]
public class MyAddin : ApplicationAddInServer
{
    private const string ADDIN_GUID = "{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}";
    private Inventor.Application _app;
    private MCGRibbonManager _ribbon;

    public void Activate(ApplicationAddInSite site, bool firstTime)
    {
        try { ActivateInternal(site, firstTime); }
        catch (Exception ex) { /* log to file, KHÔNG re-throw */ }
    }

    // ⚠ NoInlining BẮT BUỘC. JIT compile method body khi method được GỌI;
    // nếu thiếu DLL phụ thuộc, JIT exception văng ra TRƯỚC try/catch ở cùng method.
    // Tách sang method khác + NoInlining → exception propagate UP vào try/catch
    // của Activate, không thoát ra Inventor (gây mất Vault/Content Centre).
    [MethodImpl(MethodImplOptions.NoInlining)]
    private void ActivateInternal(ApplicationAddInSite site, bool firstTime)
    {
        _app = site.Application;
        _ribbon = new MCGRibbonManager(_app, ADDIN_GUID);
        _ribbon.RegisterTool(new MyToolDescriptor());
        _ribbon.RegisterTool(new AnotherToolDescriptor());
        _ribbon.Build(firstTime);   // ← truyền firstTime vào Build
    }

    public void Deactivate()
    {
        try { _ribbon?.Cleanup(); } catch { }
        _ribbon = null;
        _app = null;
    }

    public void ExecuteCommand(int id) { }
    public object Automation => null;
}
```

### Bước 5. (Optional) Tool có palette — implement `IDockablePanelDescriptor`

```csharp
internal class MyPanelDescriptor : IDockablePanelDescriptor
{
    private MyWpfPanel _panel;

    public string Id                            => "MyAddin.MyTool.DockableWindow";  // ← unique global
    public string Title                         => "My Tool";
    public DockingStateEnum DefaultDockingState => DockingStateEnum.kDockRight;
    public int MinWidth  => 220;
    public int MinHeight => 400;

    public UserControl CreateContent()
    {
        _panel = new MyWpfPanel();
        return _panel;
    }

    public void OnContentEmbedded(UserControl content)
    {
        // Wire controllers với WPF content tại đây.
    }
}
```

Trỏ `IToolDescriptor.DockablePanel` về instance của descriptor này.

---

## 3. API reference

### `MCGRibbonManager`

| Method | Khi gọi |
|---|---|
| `new MCGRibbonManager(app, addinGuid)` | Mỗi addin 1 instance, dùng `addinGuid` riêng |
| `RegisterTool(IToolDescriptor)` | Trước `Build()`, mỗi tool 1 lần |
| `Build(bool firstTime)` | Sau khi `RegisterTool` xong. **Truyền `firstTime` từ `Activate`**. Khi `firstTime=false`, RibbonTab/Panel **không** được tạo lại (Inventor đã có cached layout) |
| `Cleanup()` | Gọi từ `Deactivate()` |

### `PictureDispConverter`

| Method | Mục đích |
|---|---|
| `ToIPictureDisp(Bitmap)` | Convert `Bitmap` → `stdole.IPictureDisp` cho Inventor button icon |
| `LoadBitmapFromResource(Assembly, fullResourceName)` | Load `Bitmap` từ embedded resource |

---

## 4. Quy ước ID — phải tuân thủ để share tab đúng

| Loại | Convention | Có được tự đổi? |
|---|---|---|
| Tab ID | `id.Tab.MCGTools` (hard-coded trong `MCGRibbonManager`) | **KHÔNG** |
| Panel ID | `id.Panel.MCGTools.{Model\|Drawing\|Utility}` (hard-coded) | **KHÔNG** |
| Tool button ID | `id.Button.<AddinShort>.<ToolName>` | Có — phải unique |
| DockableWindow ID | `<AddinShort>.<ToolName>.DockableWindow` | Có — phải unique |
| Addin GUID | `XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX` (PowerShell `[guid]::NewGuid()`) | Có — bắt buộc unique |

---

## 5. Checklist tránh xung đột giữa các addin MCG

| Điểm cần check | Lý do |
|---|---|
| **Addin GUID khác nhau giữa các project** | Đụng GUID → Inventor lookup CLSID nhầm class → ribbon manager corrupt → tab Vault/Content Centre biến mất |
| **HKCU registry không còn entry stale** | RegAsm `/codebase` các build cũ ghi `HKCU\SOFTWARE\Classes\CLSID\{GUID}` với class name có thể đã đổi. Khi rename class hoặc đổi GUID, **xoá entry cũ** bằng `reg delete "HKCU\SOFTWARE\Classes\CLSID\{old-guid}" /f` |
| **`<SoftwareVersionGreaterThan>26</...>` trong .addin** | Tag chuẩn Inventor SDK. **KHÔNG** dùng `<SupportedSoftwareVersionGreaterThan>` (sai tag) hoặc value `26..` (malformed) |
| **`<Assembly>` dùng relative path** | `<Assembly>MyAddin.dll</Assembly>` cùng folder với `.addin` — chuẩn pattern Inventor 2023 |
| **Tool button ID + DockableWindow ID unique** | Đụng ID → button thứ 2 không add vào panel |

---

## 6. Layout deploy phía client

Mỗi addin một subfolder trong `%APPDATA%\Autodesk\Inventor 2023\Addins\`:

```
%APPDATA%\Autodesk\Inventor 2023\Addins\
├── MCG_InventorSymbolHandler\
│   ├── MCG_InventorSymbolHandler.addin
│   └── MCG_InventorSymbolHandler.dll
├── MCG_InventorPickDataPlugin\
│   ├── MCG_InventorPickDataPlugin.addin
│   └── MCG_InventorPickDataPlugin.dll
└── MyNewAddin\
    ├── MyNewAddin.addin
    └── MyNewAddin.dll
```

Inventor scan recursive → tìm tất cả `.addin` → mỗi addin Activate riêng → tất cả cùng register button vào tab "MCG TOOLS" → user thấy 1 tab thống nhất.

Mẫu `.addin` file (sửa theo addin của bạn):

```xml
<?xml version="1.0" encoding="utf-8"?>
<Addin Type="Standard">
  <ClassId>{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}</ClassId>
  <ClientId>{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}</ClientId>
  <DisplayName>MyAddin</DisplayName>
  <Description>Mô tả ngắn</Description>
  <Assembly>MyAddin.dll</Assembly>
  <LoadOnStartUp>1</LoadOnStartUp>
  <SoftwareVersionGreaterThan>26</SoftwareVersionGreaterThan>
</Addin>
```

---

## 7. (Khuyến nghị) File logger debug

Inventor không hiện console — exception trong `Activate` câm lặng. Pattern khuyến nghị:

```csharp
private const string LOG_DIR  = @"C:\CustomTools\Inventor\logs";
private const string LOG_FILE = @"C:\CustomTools\Inventor\logs\MyAddin.log";

public void Activate(ApplicationAddInSite site, bool firstTime)
{
    try
    {
        ActivateInternal(site, firstTime);
        DeleteLogFile();   // ← success → xoá log để ngầm báo "OK"
    }
    catch (Exception ex) { LogToFile("Activate", ex); }
}

private static void LogToFile(string phase, Exception ex)
{
    try
    {
        if (!Directory.Exists(LOG_DIR)) Directory.CreateDirectory(LOG_DIR);
        System.IO.File.AppendAllText(LOG_FILE,
            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {phase} ERROR: {ex.Message}\n{ex.StackTrace}\n\n");
    }
    catch { /* nuốt I/O errors */ }
}

private static void DeleteLogFile()
{
    try { if (System.IO.File.Exists(LOG_FILE)) System.IO.File.Delete(LOG_FILE); }
    catch { }
}
```

**Quy ước:** file log tồn tại = addin có lỗi. File không tồn tại = load sạch.

---

## 8. Quy trình tạo addin mới — checklist nhanh

```
[ ] 1. Tạo project net48 + UseWPF + UseWindowsForms + x64
[ ] 2. Copy Core/ folder + giữ nguyên namespace MCG.Inventor.Ribbon
[ ] 3. Sinh GUID mới (PowerShell: [guid]::NewGuid()) — không copy từ project khác
[ ] 4. Tạo class entry point [Guid(...)] + ApplicationAddInServer
[ ] 5. Pattern JIT-safe: Activate → ActivateInternal [MethodImpl(NoInlining)]
[ ] 6. Tạo .addin file với GUID khớp, <SoftwareVersionGreaterThan>26</...>, relative path
[ ] 7. Mỗi tool: implement IToolDescriptor; load icon dùng typeof(...).Namespace
[ ] 8. Tool có palette: implement IDockablePanelDescriptor
[ ] 9. Build → output deploy vào subfolder Addins\<AddinName>\
[ ] 10. Verify: GUID unique global; HKCU không còn entry stale từ build cũ
[ ] 11. Mở Inventor → check tab "MCG TOOLS" có button mới + Vault/Content Centre vẫn OK
```

---

## 9. Update SDK — backwards compatibility

SDK update theo nguyên tắc:
- Chỉ ADD method/property mới vào interface
- KHÔNG remove hoặc thay signature của method cũ
- Khi cần breaking change: bump namespace (vd `MCG.Inventor.Ribbon.V2`) và giữ cả 2 trong giai đoạn migrate

Khi bạn fix bug trong `Core/`, hãy sync bản mới sang **mọi addin** đang dùng để tránh code drift. Một bug fix chỉ ở 1 addin là vô nghĩa khi addin khác vẫn cùng load chung tab.

---

## 10. Files trong folder

| File | Mục đích |
|---|---|
| `MCGRibbonManager.cs` | Tạo/cleanup tab + panels + buttons + dock hosts |
| `DockableWindowHost.cs` | Embed WPF UserControl vào Inventor DockableWindow (xử lý quirks: HWND=0 retry, force-hide 2s, WM_SIZE relay) |
| `IToolDescriptor.cs` | Interface mô tả 1 tool (button + panel + palette descriptor) |
| `IDockablePanelDescriptor.cs` | Interface mô tả palette WPF (id, title, dock state, content factory) |
| `PanelLocation.cs` | Enum vị trí panel: `Model`, `Drawing`, `Utility` |
| `RibbonContext.cs` | Enum `[Flags]` ribbon: `Part`, `Assembly`, `Drawing`, `Model3D`, `All` |
| `PictureDispConverter.cs` | Convert `Bitmap` ↔ `stdole.IPictureDisp` + load embedded resource |
| `MCGRibbonReadme.md` | File này — hướng dẫn sử dụng SDK |
