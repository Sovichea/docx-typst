"""Command line entry points for docx-typst."""

from __future__ import annotations

import argparse
from importlib import resources
from pathlib import Path

from .docx_to_typst import convert as convert_to_typst
from .docx_to_typst import main as _docx_to_typst_main
from .typst_to_docx import convert as convert_to_docx
from .typst_to_docx import main as _typst_to_docx_main


def docx_to_typst() -> None:
    _docx_to_typst_main()


def typst_to_docx() -> None:
    _typst_to_docx_main()


def _read_resource(name: str) -> str:
    return resources.files("docx_typst").joinpath(name).read_text(encoding="utf-8")


def _write(path: Path, content: str, force: bool) -> None:
    if path.exists() and not force:
        raise SystemExit(f"Refusing to overwrite {path}; pass --force to replace it.")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")
    print(f"Created {path}")


def _starter_assets() -> list[tuple[str, bytes]]:
    directory = resources.files("docx_typst").joinpath("templates", "assets")
    return [
        (item.name, item.read_bytes())
        for item in sorted(directory.iterdir(), key=lambda item: item.name)
        if item.is_file()
    ]


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="docx-typst",
        description="Create or convert a Typst project for the supported DOCX subset.",
    )
    commands = parser.add_subparsers(dest="command", required=True)

    module = commands.add_parser("module", help="write only the helper module")
    module.add_argument("--output", type=Path, default=Path("docx-typst.typ"))
    module.add_argument("--force", action="store_true", help="replace an existing file")

    init = commands.add_parser("init", help="create a starter Typst document")
    init.add_argument("directory", nargs="?", type=Path, default=Path("docx-typst-project"))
    init.add_argument(
        "--force", action="store_true", help="replace generated files that already exist"
    )

    to_typst = commands.add_parser("to-typst", help="convert DOCX to a Typst project")
    to_typst.add_argument("input_docx", type=Path)
    to_typst.add_argument("--output", type=Path, default=Path("typst-output"))
    to_typst.add_argument("--strict", action="store_true")
    to_typst.add_argument("--reconstruct", action="store_true", help="rebuild from Word instead of restoring unchanged packaged source")

    to_docx = commands.add_parser("to-docx", help="convert Typst to DOCX")
    to_docx.add_argument("main_typ", type=Path)
    to_docx.add_argument("--output", type=Path, default=Path("converted.docx"))
    to_docx.add_argument("--reference", type=Path)
    to_docx.add_argument("--strict", action="store_true")

    args = parser.parse_args()
    if args.command == "module":
        _write(args.output, _read_resource("template.typ"), args.force)
        return
    if args.command == "init":
        project = args.directory
        outputs = [
            project / "docx-typst.typ",
            project / "main.typ",
            *[project / "assets" / name for name, _ in _starter_assets()],
        ]
        conflicts = [path for path in outputs if path.exists()]
        if conflicts and not args.force:
            names = ", ".join(str(path) for path in conflicts)
            raise SystemExit(f"Refusing to overwrite {names}; pass --force to replace them.")
        _write(project / "docx-typst.typ", _read_resource("template.typ"), True)
        _write(project / "main.typ", _read_resource("templates/main.typ"), True)
        for name, content in _starter_assets():
            destination = project / "assets" / name
            destination.parent.mkdir(parents=True, exist_ok=True)
            destination.write_bytes(content)
            print(f"Created {destination}")
        return
    if args.command == "to-typst":
        report = convert_to_typst(
            args.input_docx, args.output, strict=args.strict, restore_source=not args.reconstruct
        )
        print(f"Created {args.output / 'main.typ'} ({report['status']})")
        return
    report = convert_to_docx(
        args.main_typ,
        args.output,
        reference=args.reference,
        strict=args.strict,
    )
    print(f"Created {args.output} ({report['status']})")


if __name__ == "__main__":
    main()
