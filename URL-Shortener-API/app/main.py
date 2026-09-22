from fastapi import Depends, FastAPI, HTTPException, status
from fastapi.responses import RedirectResponse
from sqlalchemy.orm import Session

from . import crud, models, schemas, utils
from .auth import AuthRateLimitMiddleware
from .config import settings
from .database import engine, get_db

models.Base.metadata.create_all(bind=engine)

app = FastAPI(title="URL Shortener API")
app.add_middleware(AuthRateLimitMiddleware)


@app.get("/")
def root():
    return {"message": "URL Shortener API funcionando", "docs": "/docs"}


@app.post(
    "/shorten",
    response_model=schemas.URLResponse,
    status_code=status.HTTP_201_CREATED,
)
def shorten_url(data: schemas.URLCreate, db: Session = Depends(get_db)):
    short_code = utils.generate_short_code()
    while crud.get_url_by_code(db, short_code):
        short_code = utils.generate_short_code()

    url_entry = crud.create_short_url(db, str(data.url), short_code)
    return schemas.URLResponse(
        original_url=url_entry.original_url,
        short_code=url_entry.short_code,
        short_url=f"{settings.BASE_URL}/r/{url_entry.short_code}",
        visits=url_entry.visits,
        created_at=url_entry.created_at,
    )


@app.get("/r/{short_code}")
def redirect_to_url(short_code: str, db: Session = Depends(get_db)):
    url_entry = crud.get_url_by_code(db, short_code)
    if not url_entry:
        raise HTTPException(status_code=404, detail="Short URL not found")
    crud.increment_visits(db, url_entry)
    return RedirectResponse(url=url_entry.original_url, status_code=301)


@app.get("/stats/{short_code}", response_model=schemas.URLStats)
def get_url_stats(short_code: str, db: Session = Depends(get_db)):
    url_entry = crud.get_url_by_code(db, short_code)
    if not url_entry:
        raise HTTPException(status_code=404, detail="Short URL not found")
    return schemas.URLStats(
        original_url=url_entry.original_url,
        short_code=url_entry.short_code,
        visits=url_entry.visits,
        created_at=url_entry.created_at,
    )
