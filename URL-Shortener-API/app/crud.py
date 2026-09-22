from sqlalchemy.orm import Session

from . import models


def create_short_url(db: Session, original_url: str, short_code: str):
    db_url = models.ShortURL(
        original_url=str(original_url),
        short_code=short_code,
    )
    db.add(db_url)
    db.commit()
    db.refresh(db_url)
    return db_url


def get_url_by_code(db: Session, short_code: str):
    return db.query(models.ShortURL).filter(
        models.ShortURL.short_code == short_code
    ).first()


def increment_visits(db: Session, url: models.ShortURL):
    url.visits += 1
    db.commit()
    db.refresh(url)
    return url
