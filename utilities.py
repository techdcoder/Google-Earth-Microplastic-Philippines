def to_hex(number: int) -> str:
    out = hex(number)[2:]
    if len(out) == 1:
        out = '0' + out
    return out


def rgba(red: int, green: int, blue: int, alpha: int) -> str:
    return to_hex(red) + to_hex(green) + to_hex(blue) + to_hex(alpha)


def red_hue(value: int) -> str:
    return rgba(255, 255-value, 255-value, 150)
