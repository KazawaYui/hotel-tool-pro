import pytest

from app_loader import load_app


@pytest.fixture(scope='session')
def app():
    return load_app()
