from datetime import datetime

from pydantic import BaseModel, HttpUrl


class URLCreate(BaseModel):
    url: HttpUrl


class URLResponse(BaseModel):
    original_url: str
    short_code: str
    short_url: str
    visits: int
    created_at: datetime


class URLStats(BaseModel):
    original_url: str
    short_code: str
    visits: int
    created_at: datetime
