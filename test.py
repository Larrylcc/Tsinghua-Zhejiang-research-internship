from matplotlib.font_manager import FontManager
fm = FontManager()
font_names = sorted([f.name for f in fm.ttflist])
for name in font_names:
    if 'Chinese' in name or 'SC' in name:
        print(name)