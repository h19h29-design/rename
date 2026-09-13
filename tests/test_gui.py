import os
import sys
import pytest
import tkinter as tk
from rename_app.gui import MainWindow
from rename_app.session import Session
from rename_app.rules import Options
from tests.helpers import xlsx

pytestmark = pytest.mark.skipif(sys.platform != 'win32' and not os.environ.get('DISPLAY'), reason='Display required; run with xvfb-run')

@pytest.fixture
def window():
    root = tk.Tk()
    root.withdraw()
    app = MainWindow(root)
    yield app
    root.destroy()


def test_gui_defaults_and_review_loading(window, tmp_path):
    opts = window.get_options()
    assert opts.names and opts.rrn and opts.birth and not opts.amounts and not opts.dates
    source = xlsx(tmp_path/'demo.xlsx')
    window.add_paths([source])
    window.accept_session(Session.load([source],opts))
    assert len(window.tree.get_children()) == len(window.session.candidates)
    assert window.session is not None


def test_options_invalidate_previous_scan(window,tmp_path):
    source = xlsx(tmp_path/'demo.xlsx')
    window.add_paths([source])
    window.accept_session(Session.load([source],Options()))
    window.invalidate()
    assert window.session is None
    assert len(window.tree.get_children()) == 0


def test_candidate_toggle_and_masked_preview(window,tmp_path):
    source = xlsx(tmp_path/'demo.xlsx')
    window.add_paths([source])
    window.accept_session(Session.load([source],Options()))
    window.set_visible_selection(False)
    assert not any(c.selected for c in window.session.candidates)
    window.hide_original.set(True)
    window.refresh_table()
    assert all(window.tree.item(i,'values')[2] == '••••••' for i in window.tree.get_children())
