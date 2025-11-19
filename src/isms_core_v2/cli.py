"""
cli.py

Tiny CLI wrapper for the schema-driven ISMS document generator (v2).

Commands:
- validate: validate JSON against DocumentModel
- generate: validate JSON and render a .docx using the base template
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pydantic import ValidationError

from .models import DocumentModel
from .renderers.word_renderer import render_document

import json

def _resolve_path(p: str | Path) -> Path:
    return Path(p).expanduser().resolve()


def cmd_validate(args: argparse.Namespace) -> int:
    """
    Validate the JSON input against DocumentModel.
    """
    input_path = _resolve_path(args.input)

    if not input_path.is_file():
        print(f"[ERROR] Input JSON not found: {input_path}", file=sys.stderr)
        return 1

    try:
        with open(input_path, "r", encoding="utf-8") as f:
            raw_json = f.read()
        model = DocumentModel.model_validate_json(raw_json)
    except ValidationError as e:
        print("[ERROR] JSON validation failed:", file=sys.stderr)
        print(e, file=sys.stderr)
        return 2
    except Exception as e:
        print("[ERROR] Failed to load JSON:", file=sys.stderr)
        print(e, file=sys.stderr)
        return 3

    print(f"[OK] JSON is valid for DocumentModel.")
    print(f"      DocID:   {model.metadata.doc_id}")
    print(f"      Title:   {model.metadata.title}")
    print(f"      Type:    {model.metadata.doc_type}")
    print(f"      Version: {model.metadata.version}")
    print(f"      Status:  {model.metadata.status}")
    return 0



def cmd_generate(args: argparse.Namespace) -> int:
    """
    Validate the JSON input and generate a Word document based on the template.
    """
    input_path = _resolve_path(args.input)
    template_path = _resolve_path(args.template)
    output_path = _resolve_path(args.output)

    if not input_path.is_file():
        print(f"[ERROR] Input JSON not found: {input_path}", file=sys.stderr)
        return 1

    if not template_path.is_file():
        print(f"[ERROR] Template not found: {template_path}", file=sys.stderr)
        return 1

    try:
        with open(input_path, "r", encoding="utf-8") as f:
            raw_json = f.read()
        model = DocumentModel.model_validate_json(raw_json)
    except ValidationError as e:
        print("[ERROR] JSON validation failed:", file=sys.stderr)
        print(e, file=sys.stderr)
        return 2

    # Ensure output directory exists
    output_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        render_document(model, template_path, output_path)
    except Exception as e:  # keep this broad in the CLI to catch unexpected issues
        print("[ERROR] Failed to render document:", file=sys.stderr)
        print(repr(e), file=sys.stderr)
        return 3

    print(f"[OK] Document generated:")
    print(f"      Input JSON: {input_path}")
    print(f"      Template:   {template_path}")
    print(f"      Output:     {output_path}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="isms_core_v2",
        description="Schema-driven ISMS document generator (v2).",
    )

    subparsers = parser.add_subparsers(dest="command", required=True)

    # validate
    p_validate = subparsers.add_parser(
        "validate", help="Validate JSON against the DocumentModel schema."
    )
    p_validate.add_argument(
        "-i",
        "--input",
        required=True,
        help="Path to input JSON file.",
    )
    p_validate.set_defaults(func=cmd_validate)

    # generate
    p_generate = subparsers.add_parser(
        "generate",
        help="Validate JSON and generate a Word document from the template.",
    )
    p_generate.add_argument(
        "-i",
        "--input",
        required=True,
        help="Path to input JSON file.",
    )
    p_generate.add_argument(
        "-t",
        "--template",
        default="templates/ISMS_Base_Master.docx",
        help="Path to the Word template (.docx). "
             "Defaults to templates/ISMS_Base_Master.docx",
    )
    p_generate.add_argument(
        "-o",
        "--output",
        required=True,
        help="Path to output .docx file.",
    )
    p_generate.set_defaults(func=cmd_generate)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    func = getattr(args, "func", None)
    if func is None:
        parser.print_help()
        return 1

    return func(args)


if __name__ == "__main__":
    raise SystemExit(main())
