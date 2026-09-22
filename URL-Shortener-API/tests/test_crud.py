from app import crud


def test_create_short_url(db):
    url_entry = crud.create_short_url(db, "https://example.com", "abc123")
    assert url_entry.original_url == "https://example.com"
    assert url_entry.short_code == "abc123"
    assert url_entry.visits == 0
    assert url_entry.created_at is not None


def test_get_url_by_code(db):
    crud.create_short_url(db, "https://example.com", "xyz789")
    result = crud.get_url_by_code(db, "xyz789")
    assert result is not None
    assert result.original_url == "https://example.com"
    assert result.short_code == "xyz789"


def test_get_url_by_code_not_found(db):
    result = crud.get_url_by_code(db, "nonexistent")
    assert result is None


def test_increment_visits(db):
    url_entry = crud.create_short_url(db, "https://example.com", "vis123")
    assert url_entry.visits == 0

    crud.increment_visits(db, url_entry)
    assert url_entry.visits == 1

    crud.increment_visits(db, url_entry)
    assert url_entry.visits == 2
