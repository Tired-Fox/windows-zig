// ----- This code is automatically generated -----
pub const IXsltProcessor = extern struct {
    vtable: *const VTable,
    /// Must call `deinit` or `IUnknown.Release` on returned pointer
    pub fn cast(self: *@This(), AS: type) !*AS {
        var _r: ?*AS = undefined;
        try IUnknown.QueryInterface(@ptrCast(self), &AS.IID, @ptrCast(&_r));
        return _r.?;
    }
    pub fn deinit(self: *@This()) void {
        _ = IUnknown.Release(@ptrCast(self));
    }
    pub fn TransformToString(self: *@This(), inputNode: *IXmlNode) core.HResult!?HSTRING {
        var _r: ?HSTRING = undefined;
        const _c = self.vtable.TransformToString(@ptrCast(self), inputNode, &_r);
        try core.hresultToError(_c);
        return _r;
    }
    pub const NAME: []const u8 = "Windows.Data.Xml.Xsl.IXsltProcessor";
    pub const RUNTIME_NAME: [:0]const u16 = @import("std").unicode.utf8ToUtf16LeStringLiteral(NAME);
    pub const GUID: []const u8 = "7b64703f-550c-48c6-a90f-93a5b964518f";
    pub const IID: Guid = Guid.initString(GUID);
    pub const SIGNATURE: []const u8 = core.Signature.interface(GUID);
    pub const VTable = extern struct {
        QueryInterface: *const fn(self: *anyopaque, riid: *const Guid, ppvObject: *?*anyopaque) callconv(.winapi) HRESULT,
        AddRef: *const fn(self: *anyopaque) callconv(.winapi) u32,
        Release: *const fn(self: *anyopaque,) callconv(.winapi) u32,
        GetIids: *const fn(self: *anyopaque, iidCount: *u32, iids: *[*]const Guid) callconv(.winapi) HRESULT,
        GetRuntimeClassName: *const fn(self: *anyopaque, className: *?HSTRING) callconv(.winapi) HRESULT,
        GetTrustLevel: *const fn(self: *anyopaque, trustLevel: *TrustLevel) callconv(.winapi) HRESULT,
        TransformToString: *const fn(self: *anyopaque, inputNode: *IXmlNode, _r: *?HSTRING) callconv(.winapi) HRESULT,
    };
};
pub const IXsltProcessor2 = extern struct {
    vtable: *const VTable,
    /// Must call `deinit` or `IUnknown.Release` on returned pointer
    pub fn cast(self: *@This(), AS: type) !*AS {
        var _r: ?*AS = undefined;
        try IUnknown.QueryInterface(@ptrCast(self), &AS.IID, @ptrCast(&_r));
        return _r.?;
    }
    pub fn deinit(self: *@This()) void {
        _ = IUnknown.Release(@ptrCast(self));
    }
    pub fn TransformToDocument(self: *@This(), inputNode: *IXmlNode) core.HResult!*XmlDocument {
        var _r: *XmlDocument = undefined;
        const _c = self.vtable.TransformToDocument(@ptrCast(self), inputNode, &_r);
        try core.hresultToError(_c);
        return _r;
    }
    pub const NAME: []const u8 = "Windows.Data.Xml.Xsl.IXsltProcessor2";
    pub const RUNTIME_NAME: [:0]const u16 = @import("std").unicode.utf8ToUtf16LeStringLiteral(NAME);
    pub const GUID: []const u8 = "8da45c56-97a5-44cb-a8be-27d86280c70a";
    pub const IID: Guid = Guid.initString(GUID);
    pub const SIGNATURE: []const u8 = core.Signature.interface(GUID);
    pub const VTable = extern struct {
        QueryInterface: *const fn(self: *anyopaque, riid: *const Guid, ppvObject: *?*anyopaque) callconv(.winapi) HRESULT,
        AddRef: *const fn(self: *anyopaque) callconv(.winapi) u32,
        Release: *const fn(self: *anyopaque,) callconv(.winapi) u32,
        GetIids: *const fn(self: *anyopaque, iidCount: *u32, iids: *[*]const Guid) callconv(.winapi) HRESULT,
        GetRuntimeClassName: *const fn(self: *anyopaque, className: *?HSTRING) callconv(.winapi) HRESULT,
        GetTrustLevel: *const fn(self: *anyopaque, trustLevel: *TrustLevel) callconv(.winapi) HRESULT,
        TransformToDocument: *const fn(self: *anyopaque, inputNode: *IXmlNode, _r: **XmlDocument) callconv(.winapi) HRESULT,
    };
};
pub const IXsltProcessorFactory = extern struct {
    vtable: *const VTable,
    /// Must call `deinit` or `IUnknown.Release` on returned pointer
    pub fn cast(self: *@This(), AS: type) !*AS {
        var _r: ?*AS = undefined;
        try IUnknown.QueryInterface(@ptrCast(self), &AS.IID, @ptrCast(&_r));
        return _r.?;
    }
    pub fn deinit(self: *@This()) void {
        _ = IUnknown.Release(@ptrCast(self));
    }
    pub fn CreateInstance(self: *@This(), document: *XmlDocument) core.HResult!*XsltProcessor {
        var _r: *XsltProcessor = undefined;
        const _c = self.vtable.CreateInstance(@ptrCast(self), document, &_r);
        try core.hresultToError(_c);
        return _r;
    }
    pub const NAME: []const u8 = "Windows.Data.Xml.Xsl.IXsltProcessorFactory";
    pub const RUNTIME_NAME: [:0]const u16 = @import("std").unicode.utf8ToUtf16LeStringLiteral(NAME);
    pub const GUID: []const u8 = "274146c0-9a51-4663-bf30-0ef742146f20";
    pub const IID: Guid = Guid.initString(GUID);
    pub const SIGNATURE: []const u8 = core.Signature.interface(GUID);
    pub const VTable = extern struct {
        QueryInterface: *const fn(self: *anyopaque, riid: *const Guid, ppvObject: *?*anyopaque) callconv(.winapi) HRESULT,
        AddRef: *const fn(self: *anyopaque) callconv(.winapi) u32,
        Release: *const fn(self: *anyopaque,) callconv(.winapi) u32,
        GetIids: *const fn(self: *anyopaque, iidCount: *u32, iids: *[*]const Guid) callconv(.winapi) HRESULT,
        GetRuntimeClassName: *const fn(self: *anyopaque, className: *?HSTRING) callconv(.winapi) HRESULT,
        GetTrustLevel: *const fn(self: *anyopaque, trustLevel: *TrustLevel) callconv(.winapi) HRESULT,
        CreateInstance: *const fn(self: *anyopaque, document: *XmlDocument, _r: **XsltProcessor) callconv(.winapi) HRESULT,
    };
};
pub const XsltProcessor = extern struct {
    vtable: *const IInspectable.VTable,
    /// Must call `deinit` or `IUnknown.Release` on returned pointer
    pub fn cast(self: *@This(), AS: type) !*AS {
        var _r: ?*AS = undefined;
        try IUnknown.QueryInterface(@ptrCast(self), &AS.IID, @ptrCast(&_r));
        return _r.?;
    }
    pub fn deinit(self: *@This()) void {
        _ = IUnknown.Release(@ptrCast(self));
    }
    pub fn TransformToString(self: *@This(), inputNode: *IXmlNode) core.HResult!?HSTRING {
        const this: *IXsltProcessor = @ptrCast(self);
        return try this.TransformToString(inputNode);
    }
    pub fn TransformToDocument(self: *@This(), inputNode: *IXmlNode) core.HResult!*XmlDocument {
        var this: ?*IXsltProcessor2 = undefined;
        defer _ = IUnknown.Release(@ptrCast(this));
        try IUnknown.QueryInterface(@ptrCast(self), &IXsltProcessor2.IID, @ptrCast(&this));
        return try this.?.TransformToDocument(inputNode);
    }
    pub fn CreateInstance(document: *XmlDocument) core.HResult!*XsltProcessor {
        const _f = try @This()._IXsltProcessorFactoryCache.get();
        return try _f.CreateInstance(document);
    }
    pub const NAME: []const u8 = "Windows.Data.Xml.Xsl.XsltProcessor";
    pub const RUNTIME_NAME: [:0]const u16 = @import("std").unicode.utf8ToUtf16LeStringLiteral(NAME);
    pub const GUID: []const u8 = IXsltProcessor.GUID;
    pub const IID: Guid = IXsltProcessor.IID;
    pub const SIGNATURE: []const u8 = core.Signature.class(NAME, IXsltProcessor.SIGNATURE);
    var _IXsltProcessorFactoryCache: FactoryCache(IXsltProcessorFactory, RUNTIME_NAME) = .{};
};
const IUnknown = @import("../../root.zig").IUnknown;
const Guid = @import("../../root.zig").Guid;
const HRESULT = @import("../../root.zig").HRESULT;
const core = @import("../../root.zig").core;
const IXmlNode = @import("./Dom.zig").IXmlNode;
const XmlDocument = @import("./Dom.zig").XmlDocument;
const IInspectable = @import("../../Foundation.zig").IInspectable;
const FactoryCache = @import("../../core.zig").FactoryCache;
const TrustLevel = @import("../../root.zig").TrustLevel;
const HSTRING = @import("../../root.zig").HSTRING;
