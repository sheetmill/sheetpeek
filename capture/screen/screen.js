// The main screen's size in points and its backing scale, tab-separated. Run with
// `osascript -l JavaScript`; reads AppKit's screen list only, sends Excel nothing.
ObjC.import("AppKit");
const screen = $.NSScreen.mainScreen;
const frame = screen.frame;
[frame.size.width, frame.size.height, screen.backingScaleFactor].join("\t");
