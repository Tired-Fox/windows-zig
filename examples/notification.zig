const std = @import("std");
const windows = @import("windows");
const win32 = windows.win32;

const XmlDocument = windows.Data.Xml.Dom.XmlDocument;
const XmlElement = windows.Data.Xml.Dom.XmlElement;
const IXmlNode = windows.Data.Xml.Dom.IXmlNode;
const IInspectable = windows.Foundation.IInspectable;
const HSTRING = windows.HSTRING;
const IUnknown = windows.IUnknown;

const ToastNotificationManager = windows.UI.Notifications.ToastNotificationManager;
const ToastNotification = windows.UI.Notifications.ToastNotification;
const NotificationData = windows.UI.Notifications.NotificationData;

const TypedEventHandler = windows.Foundation.TypedEventHandler;
const ToastDismissedEventArgs = windows.UI.Notifications.ToastDismissedEventArgs;
const ToastActivatedEventArgs = windows.UI.Notifications.ToastActivatedEventArgs;
const ToastFailedEventArgs = windows.UI.Notifications.ToastFailedEventArgs;

const L = std.unicode.utf8ToUtf16LeStringLiteral;

pub fn WindowsCreateString(string: [:0]const u16) !?HSTRING {
    var result: ?HSTRING = undefined;
    if (win32.system.win_rt.WindowsCreateString(string.ptr, @intCast(string.len), &result) != 0) {
        return error.E_OUTOFMEMORY;
    }
    return result;
}

pub fn WindowsDeleteString(string: ?HSTRING) void {
    _ = win32.system.win_rt.WindowsDeleteString(string);
}

pub fn WindowsGetString(string: ?HSTRING) ?[]const u16 {
    var len: u32 = 0;
    const buffer = win32.system.win_rt.WindowsGetStringRawBuffer(string, &len);
    if (buffer) |buf| {
        return buf[0..@as(usize, @intCast(len))];
    }
    return null;
}

fn dismissNotification(_: ?*anyopaque, sender: *ToastNotification, args: *ToastDismissedEventArgs) void {
    _ = sender;
    std.debug.print("{any}\n", .{args.getReason() catch return});
    wait.store(false, .release);
}

fn activatedNotification(_: ?*anyopaque, sender: *ToastNotification, args: *IInspectable) void {
    _ = sender;

    const event_args: *ToastActivatedEventArgs = @ptrCast(@alignCast(args));

    const ea = event_args.getArguments() catch return;
    const arguments = std.unicode.utf16LeToUtf8Alloc(std.heap.smp_allocator, WindowsGetString(ea).?) catch return;
    defer std.heap.smp_allocator.free(arguments);

    std.debug.print("Activated: {s}\n", .{arguments});
    wait.store(false, .release);
}

fn failedNotification(_: ?*anyopaque, sender: *ToastNotification, args: *ToastFailedEventArgs) void {
    _ = sender;
    const result = args.getErrorCode() catch return;
    std.debug.print("[0x{X}] Toast Failure", .{result.Value});
    wait.store(false, .release);
}

fn relative_file_uri(allocator: std.mem.Allocator, path: []const u8) ![:0]const u16 {
    const file_path = try std.fs.cwd().realpathAlloc(allocator, path);
    defer allocator.free(file_path);

    const uriUtf8 = try std.fmt.allocPrint(allocator, "file:///{s}", .{file_path});
    defer allocator.free(uriUtf8);

    return try std.unicode.utf8ToUtf16LeAllocZ(allocator, uriUtf8);
}

var wait = std.atomic.Value(bool).init(true);

pub fn main() !void {
    @setEvalBranchQuota(10_000);
    const powershell_app_id: [:0]const u16 = L("{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\\WindowsPowerShell\\v1.0\\powershell.exe");

    const POWERSHELL = try WindowsCreateString(powershell_app_id);
    defer WindowsDeleteString(POWERSHELL);

    const xml_document = try XmlDocument.init();
    defer xml_document.deinit();

    {
        const toast_tag = try WindowsCreateString(L("toast"));
        defer WindowsDeleteString(toast_tag);

        const toastElement = try xml_document.CreateElement(toast_tag.?);
        defer toastElement.deinit();
        _ = try xml_document.AppendChild(@ptrCast(toastElement));

        {
            const visual_tag = try WindowsCreateString(L("visual"));
            defer WindowsDeleteString(visual_tag);

            const visualElement = try xml_document.CreateElement(visual_tag.?);
            defer visualElement.deinit();
            _ = try toastElement.AppendChild(@ptrCast(visualElement));

            {
                const binding_tag = try WindowsCreateString(L("binding"));
                defer WindowsDeleteString(binding_tag);

                const bindingElement = try xml_document.CreateElement(binding_tag.?);
                defer bindingElement.deinit();
                _ = try visualElement.AppendChild(try bindingElement.cast(IXmlNode));

                {
                    {
                        const aname = try WindowsCreateString(L("template"));
                        defer WindowsDeleteString(aname);
                        const avalue = try WindowsCreateString(L("ToastGeneric"));
                        defer WindowsDeleteString(avalue);

                        try bindingElement.SetAttribute(aname.?, avalue.?);
                    }

                    const text_tag = try WindowsCreateString(L("text"));
                    defer WindowsDeleteString(text_tag);

                    const titleElement = try xml_document.CreateElement(text_tag.?);
                    defer titleElement.deinit();
                    _ = try bindingElement.AppendChild(@ptrCast(titleElement));

                    {
                        const aname = try WindowsCreateString(L("id"));
                        defer WindowsDeleteString(aname);
                        const avalue = try WindowsCreateString(L("1"));
                        defer WindowsDeleteString(avalue);

                        try titleElement.SetAttribute(aname.?, avalue.?);
                    }
                    {
                        const aname = try WindowsCreateString(L("hint-style"));
                        defer WindowsDeleteString(aname);
                        const avalue = try WindowsCreateString(L("title"));
                        defer WindowsDeleteString(avalue);

                        try titleElement.SetAttribute(aname.?, avalue.?);
                    }

                    const title_text = try WindowsCreateString(L("{NotificationTitle}"));
                    defer WindowsDeleteString(title_text);

                    const titleText = try xml_document.CreateTextNode(title_text.?);
                    defer titleText.deinit();
                    _ = try titleElement.AppendChild(@ptrCast(titleText));
                }
            }
        }
    }

    // Above is the same as just parsing the xml
    //
    // <toast>
    //     <visual>
    //         <binding template="ToastGeneric">
    //           <text id="1" hint-style="title">Zig Windows Runtime</text>
    //         </binding>
    //     </visual>
    // </toast>

    {
        const xml = try xml_document.GetXml();
        const built_xml = try std.unicode.utf16LeToUtf8Alloc(std.heap.smp_allocator, WindowsGetString(xml).?);
        defer std.heap.smp_allocator.free(built_xml);
        std.debug.print("[XML]\n{s}\n", .{built_xml});
    }

    const notification = try ToastNotification.CreateToastNotification(xml_document);
    defer notification.deinit();

    var data = try NotificationData.init();
    defer data.deinit();
    try notification.putData(data);

    {
        const h_key = try WindowsCreateString(L("NotificationTitle"));
        defer _ = WindowsDeleteString(h_key);

        const h_title = try WindowsCreateString(L("Zig Windows Runtime"));
        defer _ = WindowsDeleteString(h_title);

        const values = try data.getValues();
        _ = try values.Insert(h_key.?, h_title.?);
    }

    const dhandler = try TypedEventHandler(ToastNotification, ToastDismissedEventArgs).init(dismissNotification);
    const dhandle = try notification.addDismissed(dhandler);

    const ahandler = try TypedEventHandler(ToastNotification, IInspectable).init(activatedNotification);
    const ahandle = try notification.addActivated(ahandler);

    const fhandler = try TypedEventHandler(ToastNotification, ToastFailedEventArgs).init(failedNotification);
    const fhandle = try notification.addFailed(fhandler);

    var notifier = try ToastNotificationManager.CreateToastNotifierWithApplicationId(POWERSHELL.?);
    defer notifier.deinit();

    try notifier.Show(notification);

    while (wait.load(.acquire)) {
        std.Thread.sleep(std.time.ns_per_ms * 500);
    }

    std.debug.print("END\n", .{});
    try notification.removeDismissed(dhandle);
    try notification.removeActivated(ahandle);
    try notification.removeFailed(fhandle);
}
