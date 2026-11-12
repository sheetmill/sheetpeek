// Excel's windows on screen, one tab-separated line each: CoreGraphics window number, left, top,
// width and height in points, title (empty for some). Run with `osascript -l JavaScript`; reads
// the window list only, sends Excel nothing.
ObjC.import("CoreGraphics");
const list = ObjC.castRefToObject($.CGWindowListCopyWindowInfo($.kCGWindowListOptionOnScreenOnly, 0));
const out = [];
for (let i = 0; i < list.count; i++) {
  const w = list.objectAtIndex(i);
  if (ObjC.unwrap(w.objectForKey("kCGWindowOwnerName")) !== "Microsoft Excel") continue;
  if (ObjC.unwrap(w.objectForKey("kCGWindowLayer")) !== 0) continue;
  const b = ObjC.deepUnwrap(w.objectForKey("kCGWindowBounds"));
  const title = ObjC.unwrap(w.objectForKey("kCGWindowName")) || "";
  out.push([ObjC.unwrap(w.objectForKey("kCGWindowNumber")), b.X, b.Y, b.Width, b.Height, title].join("\t"));
}
out.join("\n");
