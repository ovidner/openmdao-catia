import psutil
import pytest

from openmdao_catia import utils


@pytest.fixture(autouse=True)
def no_catia_running():
    for proc in psutil.process_iter():
        if proc.name() == "CNEXT.exe":
            raise RuntimeError(
                "CATIA is already running! Please close it before running the tests."
            )


def test_start():
    catia = utils.get_catia_session()

    assert utils.catia_alive(catia)
    catia.Quit()


def test_get():
    first = utils.get_catia_session()

    caption = "foobar"
    first.Caption = caption

    second = utils.get_catia_session()

    assert second.Caption == caption
    first.Quit()
