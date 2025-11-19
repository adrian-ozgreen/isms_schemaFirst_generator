from pathlib import Path
import yaml

def load_profiles(path: Path):
    return yaml.safe_load(Path(path).read_text(encoding="utf-8"))
