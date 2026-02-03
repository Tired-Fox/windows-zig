const std = @import("std");
const windows = @import("windows");
const win32 = windows.win32;

const hresultToError = windows.core.hresultToError;

const CHOOSECOLORA = win32.ui.controls.dialogs.CHOOSECOLORA;
const CHOOSEFONTA = win32.ui.controls.dialogs.CHOOSEFONTA;
const COMDLG_FILTERSPEC = win32.ui.shell.common.COMDLG_FILTERSPEC;
const FILEOPENDIALOGOPTIONS = win32.ui.shell.FILEOPENDIALOGOPTIONS;
const IFileDialog = win32.ui.shell.IFileDialog;
const IFileOpenDialog = win32.ui.shell.IFileOpenDialog;
const IFileSaveDialog = win32.ui.shell.IFileSaveDialog;
const IShellItemArray = win32.ui.shell.IShellItemArray;
const IModalWindow = win32.ui.shell.IModalWindow;
const IShellItem = win32.ui.shell.IShellItem;
const IUnknown = windows.IUnknown;
const MESSAGEBOX_STYLE = win32.ui.windows_and_messaging.MESSAGEBOX_STYLE;
const MESSAGEBOX_RESULT = win32.ui.windows_and_messaging.MESSAGEBOX_RESULT;
const COINIT_APARTMENTTHREADED = win32.system.com.COINIT_APARTMENTTHREADED;
const CLSCTX_ALL = win32.system.com.CLSCTX_ALL;
const CLSID_FileOpenDialog = win32.ui.shell.CLSID_FileOpenDialog;
const CLSID_FileSaveDialog = win32.ui.shell.CLSID_FileSaveDialog;
const IID_IFileDialog = win32.ui.shell.IID_IFileDialog;
const IID_IFileOpenDialog = win32.ui.shell.IID_IFileOpenDialog;
const IID_IFileSaveDialog = win32.ui.shell.IID_IFileSaveDialog;
const IID_IShellItem = win32.ui.shell.IID_IShellItem;
const CoCreateInstance = win32.system.com.CoCreateInstance;
const CoInitializeEx = win32.system.com.CoInitializeEx;
const CoUninitialize = win32.system.com.CoUninitialize;
const ChooseColorA = win32.ui.controls.dialogs.ChooseColorA;
const ChooseFontA = win32.ui.controls.dialogs.ChooseFontA;
const MessageBoxA = win32.ui.windows_and_messaging.MessageBoxA;

const SFGAO_FILESYSTEM: u32 = 0x40000000;

pub fn main() !void {
    var arena = std.heap.ArenaAllocator.init(std.heap.page_allocator);
    defer arena.deinit();
    const allocator = arena.allocator();

    if (CoInitializeEx(null, COINIT_APARTMENTTHREADED) != 0) {
        return error.CoInitializeFailure;
    }
    defer CoUninitialize();

    messageBox();
    pickColor();
    pickFont();
    try pickFolders(allocator);
    try pickFile(allocator);
    try saveFile(allocator);
}

fn messageBox() void {
    std.debug.print("[Message Box]\n", .{});

    const title: [:0]const u8 = "Example Message box\x00";
    const content: [:0]const u8 = "This is an example message box\x00";

    const result = MessageBoxA(null, content.ptr, title.ptr, .{ .ICONASTERISK = 1, .OKCANCEL = 1 });
    std.debug.print("  {s}\n", .{@tagName(result)});
}

fn rgb(r: u8, g: u8, b: u8) u32 {
    // Equivalent to the RGB() macro: 0x00BBGGRR (little-endian)
    return (@as(u32, r)) |
        (@as(u32, g) << 8) |
        (@as(u32, b) << 16);
}

fn pickColor() void {
    std.debug.print("[Pick Color]\n", .{});

    const start_color = rgb(0x33, 0x99, 0xFF);
    var custom_colors: [16]u32 = [2]u32{ rgb(74, 103, 65), rgb(135, 206, 235) } ++ ([_]u32{0x00FFFFFF} ** 14);

    var color = CHOOSECOLORA{
        .lStructSize = @sizeOf(CHOOSECOLORA),
        .hwndOwner = null,
        .hInstance = null,
        .rgbResult = start_color,
        .lpCustColors = @ptrCast((&custom_colors).ptr),
        //       CC_FULLOPEN | CC_RGBINIT
        .Flags = 0x00000002 | 0x00000001,
        .lCustData = 0,
        .lpfnHook = null,
        .lpTemplateName = null,
    };

    if (ChooseColorA(&color) == win32.zig.TRUE) {
        const result = color.rgbResult;
        const r: u8 = @intCast(result & 0xFF);
        const g: u8 = @intCast((result >> 8) & 0xFF);
        const b: u8 = @intCast((result >> 16) & 0xFF);
        std.debug.print("  rgb({d}, {d}, {d})\n", .{ r, g, b });
    } else {
        std.debug.print("  NO COLOR\n", .{});
    }
}

fn getDpiY(hwnd: ?win32.foundation.HWND) i32 {
    const hdc = win32.graphics.gdi.GetDC(hwnd);
    defer _ = win32.graphics.gdi.ReleaseDC(hwnd, hdc);
    return win32.graphics.gdi.GetDeviceCaps(hdc, win32.graphics.gdi.LOGPIXELSY);
}

fn ptToLfHeight(pt: i32, dpiY: i32) i32 {
    // lfHeight in logical units; negative -> character height mapping.
    // height = -round(pt * dpi / 72)
    return -@as(i32, @intCast(@divTrunc(pt * dpiY + 36, 72))); // simple rounding
}

fn pickFont() void {
    std.debug.print("[Pick Font]\n", .{});

    var lf: win32.graphics.gdi.LOGFONTA = undefined;
    lf.lfCharSet = win32.graphics.gdi.DEFAULT_CHARSET;
    lf.lfHeight = ptToLfHeight(12, getDpiY(null));

    var cf = CHOOSEFONTA{
        .lStructSize = @sizeOf(CHOOSEFONTA),
        .hwndOwner = null,
        .hDC = null,
        .lpLogFont = &lf,
        .iPointSize = 0,
        .Flags = .{ .SCREENFONTS = 1, .EFFECTS = 1, .INITTOLOGFONTSTRUCT = 1 },
        .rgbColors = 0,
        .lCustData = 0,
        .lpfnHook = null,
        .lpTemplateName = null,
        .hInstance = null,
        .lpszStyle = null,
        .nFontType = .{},
        .___MISSING_ALIGNMENT__ = 0,
        .nSizeMin = 0,
        .nSizeMax = 0,
    };

    if (ChooseFontA(&cf) == win32.zig.TRUE) {
        const pt10 = cf.iPointSize; // thenths of a point
        const color = cf.rgbColors;

        // This is how you would create an HFONT you can select into an HDC
        const hfont = win32.graphics.gdi.CreateFontIndirectA(&lf);
        defer _ = win32.graphics.gdi.DeleteObject(hfont);

        std.debug.print("Picked ~{d}.{d}pt, color=0x{X:0>6}\n", .{ @divTrunc(pt10, 10), @mod(pt10, 10), color & 0x00FFFFFF });
        // lf now has face name, weight, italic, charset, etc.
        std.debug.print("{s}\n", .{std.mem.sliceTo(&lf.lfFaceName, 0)});
    } else {
        std.debug.print("  NO FONT\n", .{});
    }
}

fn pickFolders(allocator: std.mem.Allocator) !void {
    std.debug.print("[Pick Folders]\n", .{});
    var file_open_dialog: *IFileOpenDialog = undefined;
    //                                                           v Must use this iid to get access to IFileOpenDialog::GetResults
    if (CoCreateInstance(CLSID_FileOpenDialog, null, CLSCTX_ALL, IID_IFileOpenDialog, @ptrCast(&file_open_dialog)) != 0) {
        return error.CoCreateInstanceFailure;
    }
    defer _ = IUnknown.Release(@ptrCast(file_open_dialog));

    var file_dialog: *IFileDialog = @ptrCast(file_open_dialog);

    var options: FILEOPENDIALOGOPTIONS = undefined;
    var hresult = file_dialog.GetOptions(&options);
    if (hresult != 0) try hresultToError(hresult);

    // Make the picker pick multiple folders
    options.PICKFOLDERS = 1;
    options.DONTADDTORECENT = 1;
    options.ALLOWMULTISELECT = 1;
    options.PATHMUSTEXIST = 1;

    hresult = file_dialog.SetOptions(options);
    if (hresult != 0) try hresultToError(hresult);

    var iid: ?*IShellItem = undefined;
    _ = win32.ui.shell.SHGetKnownFolderItem(&win32.ui.shell.FOLDERID_Documents, win32.ui.shell.KF_FLAG_DEFAULT, null, IID_IShellItem, @ptrCast(&iid));
    _ = file_dialog.SetFolder(iid);
    _ = IUnknown.Release(@ptrCast(iid));

    const title = std.unicode.utf8ToUtf16LeStringLiteral("Open Folders");
    _ = file_dialog.SetTitle(title.ptr);

    var modal: *IModalWindow = @as(*IModalWindow, @ptrCast(file_open_dialog));
    switch (@as(u32, @bitCast(modal.Show(null)))) {
        0 => {},
        0x800704C7 => return error.UserCancelled,
        else => return error.DialogShowFailure,
    }

    var shell_items: ?*IShellItemArray = null;
    hresult = file_open_dialog.GetResults(&shell_items);
    if (hresult != 0) try hresultToError(hresult);

    if (shell_items) |items| {
        defer _ = IUnknown.Release(@ptrCast(items));

        var len: u32 = 0;
        _ = items.GetCount(&len);

        for (0..len) |i| {
            var item: ?*IShellItem = undefined;
            hresult = items.GetItemAt(@intCast(i), &item);
            if (hresult != 0) try hresultToError(hresult);
            defer _ = IUnknown.Release(@ptrCast(item.?));

            var attrs: u32 = 0;
            hresult = item.?.GetAttributes(SFGAO_FILESYSTEM, &attrs);
            if (hresult != 0) try hresultToError(hresult);
            // If filesystem isn't in the attributes skip the item
            if (attrs & SFGAO_FILESYSTEM == 0) continue;

            var name: ?[*:0]u16 = undefined;
            hresult = item.?.GetDisplayName(.DESKTOPABSOLUTEPARSING, &name);
            if (hresult != 0) try hresultToError(hresult);

            if (name) |n| {
                defer win32.system.com.CoTaskMemFree(@ptrCast(name));
                const utf8Name = try std.unicode.utf16LeToUtf8Alloc(allocator, std.mem.sliceTo(n, 0));
                defer allocator.free(utf8Name);

                std.debug.print("  {s}\n", .{utf8Name});
            }
        }
    } else {
        std.debug.print(" NO RESULTS", .{});
    }
}

fn pickFile(allocator: std.mem.Allocator) !void {
    std.debug.print("[Pick File]\n", .{});
    var file_open_dialog: *IFileDialog = undefined;
    //                                                           v Can use this base interface as IFileOpenDialog::GetResults isn't needed
    if (CoCreateInstance(CLSID_FileOpenDialog, null, CLSCTX_ALL, IID_IFileDialog, @ptrCast(&file_open_dialog)) != 0) {
        return error.CoCreateInstanceFailure;
    }
    defer _ = IUnknown.Release(@ptrCast(file_open_dialog));

    var options: FILEOPENDIALOGOPTIONS = undefined;
    var hresult = file_open_dialog.GetOptions(&options);
    if (hresult != 0) try hresultToError(hresult);

    options.DONTADDTORECENT = 1;
    options.FILEMUSTEXIST = 1;

    hresult = file_open_dialog.SetOptions(options);
    if (hresult != 0) try hresultToError(hresult);

    var modal: *IModalWindow = @as(*IModalWindow, @ptrCast(file_open_dialog));
    switch (@as(u32, @bitCast(modal.Show(null)))) {
        0 => {},
        0x800704C7 => return error.UserCancelled,
        else => return error.DialogShowFailure,
    }

    var shell_items: ?*IShellItem = null;
    hresult = file_open_dialog.GetResult(&shell_items);
    if (hresult != 0) try hresultToError(hresult);

    if (shell_items) |item| {
        defer _ = IUnknown.Release(@ptrCast(item));
        var name: ?[*:0]u16 = undefined;
        hresult = item.GetDisplayName(.DESKTOPABSOLUTEPARSING, &name);
        if (hresult != 0) try hresultToError(hresult);

        if (name) |n| {
            defer win32.system.com.CoTaskMemFree(@ptrCast(name));
            const utf8Name = try std.unicode.utf16LeToUtf8Alloc(allocator, std.mem.sliceTo(n, 0));
            defer allocator.free(utf8Name);

            std.debug.print("  {s}\n", .{utf8Name});
        }
    } else {
        std.debug.print("  NO RESULT\n", .{});
    }
}

fn saveFile(allocator: std.mem.Allocator) !void {
    std.debug.print("[Save File]\n", .{});
    var file_save_dialog: *IFileSaveDialog = undefined;
    if (CoCreateInstance(CLSID_FileSaveDialog, null, CLSCTX_ALL, IID_IFileSaveDialog, @ptrCast(&file_save_dialog)) != 0) {
        return error.CoCreateInstanceFailure;
    }
    defer _ = IUnknown.Release(@ptrCast(file_save_dialog));

    const file_dialog: *IFileDialog = @ptrCast(file_save_dialog);

    const filename = std.unicode.utf8ToUtf16LeStringLiteral("example.txt");
    _ = file_dialog.SetFileName(filename.ptr);

    const filters: [2]COMDLG_FILTERSPEC = [_]COMDLG_FILTERSPEC{ .{
        .pszName = std.unicode.utf8ToUtf16LeStringLiteral("Text Documents"),
        .pszSpec = std.unicode.utf8ToUtf16LeStringLiteral("*.txt"),
    }, .{
        .pszName = std.unicode.utf8ToUtf16LeStringLiteral("All Files"),
        .pszSpec = std.unicode.utf8ToUtf16LeStringLiteral("*.*"),
    } };
    _ = file_dialog.SetFileTypes(2, &filters);
    _ = file_dialog.SetFileTypeIndex(1);

    var iid: ?*IShellItem = undefined;
    _ = win32.ui.shell.SHGetKnownFolderItem(&win32.ui.shell.FOLDERID_Documents, win32.ui.shell.KF_FLAG_DEFAULT, null, IID_IShellItem, @ptrCast(&iid));
    _ = file_dialog.SetFolder(iid);
    _ = IUnknown.Release(@ptrCast(iid));

    var modal: *IModalWindow = @as(*IModalWindow, @ptrCast(file_save_dialog));
    switch (@as(u32, @bitCast(modal.Show(null)))) {
        0 => {},
        0x800704C7 => return error.UserCancelled,
        else => return error.DialogShowFailure,
    }

    var shell_items: ?*IShellItem = null;
    var hresult = file_dialog.GetResult(&shell_items);
    if (hresult != 0) try hresultToError(hresult);

    if (shell_items) |item| {
        defer _ = IUnknown.Release(@ptrCast(item));
        var name: ?[*:0]u16 = undefined;
        hresult = item.GetDisplayName(.DESKTOPABSOLUTEPARSING, &name);
        if (hresult != 0) try hresultToError(hresult);

        if (name) |n| {
            defer win32.system.com.CoTaskMemFree(@ptrCast(name));
            const utf8Name = try std.unicode.utf16LeToUtf8Alloc(allocator, std.mem.sliceTo(n, 0));
            defer allocator.free(utf8Name);

            std.debug.print("  {s}\n", .{utf8Name});
        }
    } else {
        std.debug.print("  NO RESULT\n", .{});
    }
}

fn writeWinError(writer: *std.io.Writer, error_code: u32) !void {
    try writer.print("{} (", .{error_code});
    var buf: [300]u8 = undefined;
    const len = win32.system.diagnostics.debug.FormatMessageA(
        .{ .FROM_SYSTEM = 1, .IGNORE_INSERTS = 1 },
        null,
        error_code,
        0,
        @ptrCast(&buf),
        buf.len,
        null,
    );
    if (len == 0) {
        try writer.writeAll("unknown error");
    }
    const msg = std.mem.trimRight(u8, buf[0..len], "\r\n");
    try writer.writeAll(msg);
    if (len + 1 >= buf.len) {
        try writer.writeAll("...");
    }
    try writer.writeAll(")");
}
