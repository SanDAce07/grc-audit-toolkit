from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path


REPOSITORY_ROOT = Path(__file__).resolve().parents[1]


def load_audit_script(filename):
    """Load a hyphenated audit script as an importable module for testing."""
    path = REPOSITORY_ROOT / "audit-scripts" / filename
    module_name = filename.replace("-", "_").removesuffix(".py")
    spec = spec_from_file_location(module_name, path)
    module = module_from_spec(spec)
    spec.loader.exec_module(module)
    return module
