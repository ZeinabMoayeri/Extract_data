import arabic_reshaper
from bidi.algorithm import get_display
import unicodedata

def correct_persian_text(text):
    if isinstance(text, str):
        reshaped_text = arabic_reshaper.reshape(text)
        bidi_text = get_display(reshaped_text)
        normalized = unicodedata.normalize('NFKC', bidi_text)
        return normalized
    return text