import os

import pytest

streamlit = pytest.importorskip('streamlit')
from streamlit.testing.v1 import AppTest  # noqa: E402

APP = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'vercal_web.py')


def run_app(timeout=300):
    at = AppTest.from_file(APP, default_timeout=timeout)
    at.run()
    return at


class TestWebApp:
    def test_starts_without_exception(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        at = run_app()
        assert not at.exception
        assert [b.label for b in at.button] == ['Create Calendar']

    def test_settings_of_the_same_run_are_used(self, tmp_path, monkeypatch):
        # ボタンをコールバックにしていた頃は，前回の実行時の設定で作られていた
        monkeypatch.chdir(tmp_path)
        at = run_app()
        at.number_input[0].set_value(2030)
        at.button[0].click().run()
        assert not at.exception
        assert (tmp_path / '2030_calendar.pdf').exists()

    def test_empty_day_range_shows_an_error(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        at = run_app()
        at.slider[0].set_range(6, 6)
        at.button[0].click().run()
        assert not at.exception
        assert len(at.error) == 1
        assert not list(tmp_path.glob('*.pdf'))
