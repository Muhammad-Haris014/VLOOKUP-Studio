"""Smoke tests of the Streamlit page (file uploads are covered by the engine tests)."""
from pathlib import Path

from streamlit.testing.v1 import AppTest

APP = str(Path(__file__).resolve().parent.parent / "app.py")


def test_start_page_renders_without_errors(monkeypatch):
    monkeypatch.delenv("VLOOKUP_PASSWORD", raising=False)
    at = AppTest.from_file(APP, default_timeout=30).run()
    assert not at.exception
    assert len(at.sidebar.toggle) >= 5


def test_password_gate(monkeypatch):
    monkeypatch.setenv("VLOOKUP_PASSWORD", "s3cret")
    at = AppTest.from_file(APP, default_timeout=30).run()
    assert not at.exception and len(at.text_input) == 1  # only the password box is shown

    at.text_input[0].input("wrong")
    at.button[0].click().run()
    assert any("Wrong password" in e.value for e in at.error)

    at.text_input[0].input("s3cret")
    at.button[0].click().run()
    assert not at.exception and at.session_state["signed_in"]
    assert len(at.sidebar.toggle) >= 5
