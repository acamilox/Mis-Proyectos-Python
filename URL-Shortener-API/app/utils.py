import secrets
import string

from .config import settings


def generate_short_code(length: int = None) -> str:
    if length is None:
        length = settings.SHORT_CODE_LENGTH
    alphabet = string.ascii_letters + string.digits
    return "".join(secrets.choice(alphabet) for _ in range(length))
