# Icon Behavior - Excel Master Chart Creator

## Current Behavior

When you launch the Excel Master Chart Creator app, you will see **TWO icons** in the Dock:

1. **Your custom chart icon** (blue background with colorful bars) - This is the main app
2. **Python rocket icon** - This is the Python framework running the app

Both icons appear because:
- The `.app` bundle shows with your custom icon
- Python's tkinter framework shows its own icon

## Why This Happens

This is **normal behavior** for Python GUI applications on macOS that aren't fully bundled as standalone apps. The Python framework that runs tkinter creates its own Dock presence.

## Solutions Attempted

We tried several approaches from macOS development best practices:

1. ✅ **LSUIElement in Info.plist** - Hides the app completely (too restrictive)
2. ✅ **PyObjC setActivationPolicy** - Prevents app from staying active
3. ✅ **Custom .app bundle** - Works but shows both icons
4. ⚠️ **py2app standalone** - Would work but has code-signing issues on your system

## Recommendation

**Accept both icons** - Your app works perfectly. The presence of two icons is cosmetic and doesn't affect functionality.

### Alternative: Hide Python Icon Manually

If the Python icon bothers you:
1. Right-click the Python icon in the Dock
2. Select "Options" → "Remove from Dock"
3. The Python icon will disappear (but may reappear on next launch)

## Future Solution

For a fully standalone app with only one icon, you would need:
- Successful py2app build without code-signing errors
- Or distribution through Mac App Store (handles signing automatically)
- Or professional code-signing certificate

For now, your app is **fully functional** - just with two visual indicators in the Dock.
