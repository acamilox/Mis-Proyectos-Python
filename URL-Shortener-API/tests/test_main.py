

def test_root(client):
    response = client.get("/")
    assert response.status_code == 200
    data = response.json()
    assert "message" in data
    assert "docs" in data


def test_shorten_url(client, auth_headers):
    response = client.post(
        "/shorten",
        json={"url": "https://example.com/very-long-url"},
        headers=auth_headers,
    )
    assert response.status_code == 201
    data = response.json()
    assert data["original_url"] == "https://example.com/very-long-url"
    assert "short_code" in data
    assert "short_url" in data
    assert data["visits"] == 0
    assert "created_at" in data


def test_shorten_invalid_url(client, auth_headers):
    response = client.post(
        "/shorten",
        json={"url": "not-a-valid-url"},
        headers=auth_headers,
    )
    assert response.status_code == 422


def test_redirect(client, auth_headers):
    # Create a URL first
    create_response = client.post(
        "/shorten",
        json={"url": "https://example.com"},
        headers=auth_headers,
    )
    short_code = create_response.json()["short_code"]

    # Follow redirect manually (TestClient doesn't follow by default)
    response = client.get(
        f"/r/{short_code}", headers=auth_headers, follow_redirects=False
    )
    assert response.status_code == 301
    assert response.headers["location"] == "https://example.com/"


def test_redirect_not_found(client, auth_headers):
    response = client.get("/r/nonexistent", headers=auth_headers)
    assert response.status_code == 404


def test_stats(client, auth_headers):
    # Create a URL first
    create_response = client.post(
        "/shorten",
        json={"url": "https://example.com"},
        headers=auth_headers,
    )
    short_code = create_response.json()["short_code"]

    # Check stats
    response = client.get(f"/stats/{short_code}", headers=auth_headers)
    assert response.status_code == 200
    data = response.json()
    assert data["original_url"] == "https://example.com/"
    assert data["short_code"] == short_code
    assert data["visits"] == 0


def test_stats_not_found(client, auth_headers):
    response = client.get("/stats/nonexistent", headers=auth_headers)
    assert response.status_code == 404


def test_api_key_required(client):
    response = client.post(
        "/shorten",
        json={"url": "https://example.com"},
    )
    assert response.status_code == 401


def test_api_key_invalid(client):
    response = client.post(
        "/shorten",
        json={"url": "https://example.com"},
        headers={"X-API-Key": "wrong-key"},
    )
    assert response.status_code == 401


def test_stats_increments_on_redirect(client, auth_headers):
    # Create a URL
    create_response = client.post(
        "/shorten",
        json={"url": "https://example.com"},
        headers=auth_headers,
    )
    short_code = create_response.json()["short_code"]

    # Redirect twice
    client.get(f"/r/{short_code}", headers=auth_headers, follow_redirects=False)
    client.get(f"/r/{short_code}", headers=auth_headers, follow_redirects=False)

    # Check visits
    response = client.get(f"/stats/{short_code}", headers=auth_headers)
    assert response.status_code == 200
    assert response.json()["visits"] == 2
