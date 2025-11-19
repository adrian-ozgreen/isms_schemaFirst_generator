from pathlib import Path

def load_blocks(path: Path) -> dict:
    """Load all .md and .txt files in the given folder into a dict."""
    blocks = {}
    if not path.exists():
        return blocks
    for f in list(path.glob("*.md")) + list(path.glob("*.txt")):
        blocks[f.stem] = f.read_text(encoding="utf-8").strip()
    return blocks


def merge_content_blocks(data: dict, blocks: dict):
    """Recursively replace any {"use_block": "<key>"} entries with block text."""
    def replace(node):
        if isinstance(node, dict):
            if "use_block" in node:
                key = node["use_block"]
                return blocks.get(key, f"[Missing content block: {key}]")
            return {k: replace(v) for k, v in node.items()}
        if isinstance(node, list):
            return [replace(v) for v in node]
        return node
    return replace(data)
