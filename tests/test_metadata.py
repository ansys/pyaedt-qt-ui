from ansys.aedt_qt_ui.library import __version__


def test_pkg_version():
    assert __version__ == "0.1.dev0"
