// ----- This code is automatically generated -----
pub const DragDropModifiers = packed struct(u32) {
    Shift: bool = false,
    Control: bool = false,
    Alt: bool = false,
    LeftButton: bool = false,
    MiddleButton: bool = false,
    RightButton: bool = false,
    _m: u26 = 0,
};
pub const Core = @import("./DragDrop/Core.zig");
