"""Every call sheetpeek makes into xlformat, xlrecalc or xlio that takes a ``platform`` passes one.

The providers' defaults follow their own default platform; sheetpeek shows a sheet as the platform
it was read with, so a call that left the platform out would follow the provider's default
instead. The source is read, not run: each call is resolved to the provider's function through
the module's imports, and its signature says whether it takes ``platform``.
"""

import ast
import importlib
import inspect
from pathlib import Path

SOURCE = Path(__file__).resolve().parents[1] / "src" / "sheetpeek"
PROVIDERS = ("xlformat", "xlrecalc", "xlio")


def _target(dotted):
    names = dotted.split(".")
    found = importlib.import_module(names[0])
    for name in names[1:]:
        found = getattr(found, name)
    return found


def calls_without_platform(source=SOURCE):
    """``(file:line, function)`` of each call into a provider that takes no ``platform`` given."""
    missing = []
    for path in sorted(source.rglob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        alias = {}
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for name in node.names:
                    if name.name.split(".")[0] in PROVIDERS:
                        alias[name.asname or name.name] = name.name
            elif (
                isinstance(node, ast.ImportFrom)
                and node.module
                and node.module.split(".")[0] in PROVIDERS
            ):
                for name in node.names:
                    alias[name.asname or name.name] = f"{node.module}.{name.name}"
        for node in ast.walk(tree):
            if not isinstance(node, ast.Call):
                continue
            parts, func = [], node.func
            while isinstance(func, ast.Attribute):
                parts.insert(0, func.attr)
                func = func.value
            if not isinstance(func, ast.Name) or func.id not in alias:
                continue
            dotted = ".".join([alias[func.id], *parts])
            try:
                parameters = inspect.signature(_target(dotted)).parameters
            except (AttributeError, TypeError, ValueError):
                continue
            given = {keyword.arg for keyword in node.keywords}
            if "platform" in parameters and "platform" not in given and None not in given:
                missing.append((f"{path.relative_to(source)}:{node.lineno}", dotted))
    return missing


def test_every_provider_call_that_takes_a_platform_passes_one():
    assert calls_without_platform() == []


def test_the_scan_finds_a_call_without_one(tmp_path):
    # The guard's own check: a call that leaves the platform out is found.
    (tmp_path / "probe.py").write_text(
        "import xlformat\nxlformat.row_metrics(None)\nxlformat.row_metrics(None, platform=1)\n",
        encoding="utf-8",
    )
    assert calls_without_platform(tmp_path) == [("probe.py:2", "xlformat.row_metrics")]
