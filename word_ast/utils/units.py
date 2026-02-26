TWIPS_PER_POINT = 20


def pt_to_half_points(pt: float | None) -> int | None:
    if pt is None:
        return None
    return int(round(pt * 2))


def half_points_to_pt(half_points: int | None) -> float | None:
    if half_points is None:
        return None
    return half_points / 2
