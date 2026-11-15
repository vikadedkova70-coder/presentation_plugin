from conftest import make_pptx

PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation"


def post(client, data: bytes, name="deck.pptx", mime=PPTX, title="Отчёт"):
    return client.post(
        "/api/exports",
        data={"title": title, "description": "d", "tags": "a, b ,a"},
        files={"file": (name, data, mime)},
    )


def test_health(client):
    assert client.get("/health").json() == {"status": "ok"}


def test_success(client, fake_drive):
    r = post(client, make_pptx("ok1"))
    assert r.status_code == 201
    body = r.json()
    assert body["status"] == "success" and body["drive_url"].startswith("https://drive.example/")
    assert len(body["sha256"]) == 64
    assert fake_drive.calls == 1


def test_duplicate_returns_409_with_link(client):
    data = make_pptx("dup")
    assert post(client, data).status_code == 201
    r = post(client, data)
    assert r.status_code == 409
    assert r.json()["detail"]["drive_url"]


def test_wrong_extension_415(client):
    assert post(client, make_pptx("x"), name="deck.pdf").status_code == 415


def test_wrong_mime_415(client):
    assert post(client, make_pptx("y"), mime="text/plain").status_code == 415


def test_not_a_pptx_400(client):
    assert post(client, b"hello not a zip").status_code == 400


def test_empty_400(client):
    assert post(client, b"").status_code == 400


def test_too_large_413(client):
    assert post(client, make_pptx("big", extra_bytes=2 * 1024 * 1024)).status_code == 413


def test_missing_title_422(client):
    r = client.post("/api/exports", data={"title": ""}, files={"file": ("a.pptx", make_pptx("t"), PPTX)})
    assert r.status_code == 422


def test_drive_failure_502_and_logged(client, fake_drive):
    fake_drive.fail = True
    assert post(client, make_pptx("fail1")).status_code == 502
    assert any(x["status"] == "failed" for x in client.get("/api/exports").json())
    fake_drive.fail = False  # после сбоя повторная загрузка разрешена
    assert post(client, make_pptx("fail1")).status_code == 201


def test_admin_log_lists_success(client):
    post(client, make_pptx("log1"), title="Журнал")
    log = client.get("/api/exports").json()
    assert log[0]["title"] == "Журнал" and log[0]["status"] == "success"
