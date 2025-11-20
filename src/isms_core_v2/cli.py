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
from .importers.word_importer import import_word_to_document_dict

#from src.isms_core_v2 import registers  # or from . import registers if cli.py is inside the package
from . import registers

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


    # After successful generation, optionally update registers
    if getattr(args, "update_dcr", None):
        registers.update_document_control_register(
            register_path=args.update_dcr,
            model=model,
            output_path=output_path,
        )
        print(f"[INFO] Updated Document Control Register: {args.update_dcr}")

    if getattr(args, "update_mrr", None):
        registers.update_master_reference_register(
            register_path=args.update_mrr,
            model=model,
        )
        print(f"[INFO] Updated Master Reference Register: {args.update_mrr}")





    return 0


def cmd_import_word(args: argparse.Namespace) -> int:
    """
    Import a Word document and convert it into JSON compatible with DocumentModel.
    """
    input_path = _resolve_path(args.input)
    output_path = _resolve_path(args.output)

    if not input_path.is_file():
        print(f"[ERROR] Input Word document not found: {input_path}", file=sys.stderr)
        return 1

    try:
        doc_dict = import_word_to_document_dict(
            path=input_path,
            doc_type=args.doc_type,
            default_doc_id=args.doc_id,
        )
        # Optional: soft validation
        try:
            _ = DocumentModel.model_validate(doc_dict)
            print("[OK] Imported document validates against DocumentModel.")
        except ValidationError as e:
            print("[WARN] Imported document did not fully validate:", file=sys.stderr)
            print(e, file=sys.stderr)
    except Exception as e:
        print("[ERROR] Failed to import Word document:", file=sys.stderr)
        print(e, file=sys.stderr)
        return 2

    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with output_path.open("w", encoding="utf-8") as f:
            json.dump(doc_dict, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print("[ERROR] Failed to write output JSON:", file=sys.stderr)
        print(e, file=sys.stderr)
        return 3

    print(f"[INFO] Wrote imported JSON to: {output_path}")
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
    p_generate.add_argument(
        "--update-dcr",
        metavar="PATH",
        type=Path,
        help="Optional path to Document Control Register CSV to update/create.",
    )

    p_generate.add_argument(
        "--update-mrr",
        metavar="PATH",
        type=Path,
        help="Optional path to Master Reference Register CSV to update/create.",
    )
    p_generate.set_defaults(func=cmd_generate)

    # import-word
    p_import = subparsers.add_parser(
        "import-word",
        help="Import a Word document and produce a JSON file for the generator.",
    )
    p_import.add_argument(
        "input",
        help="Path to input Word document (.docx).",
    )
    p_import.add_argument(
        "output",
        help="Path to output JSON file.",
    )
    p_import.add_argument(
        "--doc-type",
        default="Record",
        help="ISMS document type (e.g. Record, Policy, Procedure). Defaults to Record.",
    )
    p_import.add_argument(
        "--doc-id",
        default="REC-UNKNOWN-000",
        help="DocID to use if not derivable from the Word file.",
    )
    p_import.set_defaults(func=cmd_import_word)
    
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
