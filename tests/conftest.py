import pytest


# https://stackoverflow.com/a/54736376/6107981
pytest_plugins = ["tests.xlwings_fixtures"]


# https://docs.pytest.org/en/latest/example/simple.html#control-skipping-of-tests-according-to-command-line-option
# The idea here is that the tests that depend on xlwings won't run unless you pass the
# --run-xlwings command-line parameter to pytest. That way it won't cause problems in CI
# where Excel is not installed.
def pytest_addoption(parser):
    parser.addoption(
        "--run-xlwings",
        action="store_true",
        default=False,
        help="run tests that depend on xlwings",
    )
    parser.addoption(
        "--max-examples",
        action="store",
        default=100,
        help="max examples to used per property-based test",
        type=int,
    )


CONFIG = {}


def pytest_configure(config):
    config.addinivalue_line(
        "markers", "xlwings: mark test as depending on xlwings package"
    )

    # hacky way of getting custom configs as command-line arguments, without needing to
    # access the data in a test via a fixture. use like this in a test module:
    #
    #   from tests.conftest import CONFIG
    #   my_value = CONFIG["some-key"]

    CONFIG["max-examples"] = config.getoption("--max-examples")


def pytest_collection_modifyitems(config, items):
    if config.getoption("--run-xlwings"):
        return

    skip_xlwings = pytest.mark.skip(reason="need --run-xlwings option to run")
    for item in items:
        if "xlwings" in item.keywords:
            item.add_marker(skip_xlwings)
