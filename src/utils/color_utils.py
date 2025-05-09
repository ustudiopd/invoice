from PyQt5.QtGui import QColor

def apply_tint(hex_rgb, tint):
    """hex_rgb: 'RRGGBB', tint: -1.0~1.0"""
    ch = [int(hex_rgb[0:2], 16), int(hex_rgb[2:4], 16), int(hex_rgb[4:6], 16)]
    out = [0, 0, 0]
    for i in range(3):
        c = ch[i]
        if tint < 0:
            nc = c * (1 + tint)
        else:
            nc = c * (1 - tint) + 255 * tint
        out[i] = max(0, min(int(round(nc)), 255))
    return QColor(out[0], out[1], out[2]) 