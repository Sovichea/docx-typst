"""Word face names for Typst weights that are not represented by w:b."""


def weight_face(family: str, weight: str | None) -> str:
    names = {"thin": "Thin", "extralight": "ExtraLight", "light": "Light",
             "medium": "Medium", "semibold": "SemiBold", "extrabold": "ExtraBold", "black": "Black"}
    face = names.get(weight or "")
    return family + " " + face if face and not family.endswith(" " + face) else family
