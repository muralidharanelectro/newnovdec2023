#!/usr/bin/env python3
"""Automate the full `combine_results.py` workflow.

This helper script performs the following steps:

1. Confirms the required Excel workbooks are present.
2. Creates (or reuses) a virtual environment.
3. Installs Python dependencies from ``requirements.txt``.
4. Executes ``combine_results.py`` with the provided arguments.
5. Optionally runs ``generate_class_performance_summary.py`` to build the
   gender/community snapshot.

The script is designed to be idempotent and safe to run multiple times.
"""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Iterable, Sequence

PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_REQUIREMENTS = PROJECT_ROOT / "requirements.txt"
DEFAULT_VENV = PROJECT_ROOT / ".venv"


class CommandError(RuntimeError):
    """Raised when a subprocess invocation fails."""


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Create a virtual environment, install dependencies, and run "
            "combine_results.py using the Excel workbooks located in the "
            "current directory."
        )
    )
    parser.add_argument(
        "--input",
        default="ALL SEM RESULTS.xlsx",
        help="Path to the master results workbook (default: %(default)s)",
    )
    parser.add_argument(
        "--subject-catalog",
        default="CGPA.xlsx",
        help="Path to the subject metadata workbook (default: %(default)s)",
    )
    parser.add_argument(
        "--biodata",
        default="biodata.xlsx",
        help="Path to the biodata workbook (default: %(default)s)",
    )
    parser.add_argument(
        "--config",
        default="config.json",
        help="Path to the JSON configuration file (default: %(default)s)",
    )
    parser.add_argument(
        "--outdir",
        default="outputs",
        help="Directory where analysis artefacts will be written",
    )
    parser.add_argument(
        "--current-semester",
        type=int,
        default=None,
        help=(
            "Semester number currently being analysed. Falls back to the value "
            "defined in config.json when omitted."
        ),
    )
    parser.add_argument(
        "--python",
        default=sys.executable,
        help=(
            "Python interpreter used to create the virtual environment. "
            "Defaults to the interpreter running this script."
        ),
    )
    parser.add_argument(
        "--venv",
        default=str(DEFAULT_VENV),
        help="Location of the virtual environment to create/use",
    )
    parser.add_argument(
        "--requirements",
        default=str(DEFAULT_REQUIREMENTS),
        help="Path to requirements file used for installation",
    )
    parser.add_argument(
        "--class-performance-output",
        default="class_performance_summary.xlsx",
        help=(
            "Write the optional class performance snapshot to this path. "
            "Use '-' to skip the snapshot entirely."
        ),
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print the actions that would be performed without executing them",
    )
    parser.add_argument(
        "--keep-venv",
        action="store_true",
        help="Do not delete the virtual environment when a step fails",
    )
    parser.add_argument(
        "--extra-combine-args",
        nargs=argparse.REMAINDER,
        help=(
            "Additional arguments forwarded to combine_results.py after the "
            "generated CLI parameters."
        ),
    )
    return parser.parse_args(argv)


def check_files_exist(paths: Iterable[Path]) -> None:
    missing = [p for p in paths if not p.exists()]
    if missing:
        formatted = "\n".join(f"  - {p}" for p in missing)
        raise FileNotFoundError(
            "The following required files were not found:\n" f"{formatted}"
        )


def run_command(cmd: Sequence[str], *, dry_run: bool = False, env: dict | None = None) -> None:
    display_cmd = " ".join(cmd)
    print(f"[CMD] {display_cmd}")
    if dry_run:
        return
    completed = subprocess.run(cmd, env=env)
    if completed.returncode != 0:
        raise CommandError(f"Command failed with exit code {completed.returncode}: {display_cmd}")


def ensure_virtualenv(venv_path: Path, *, python: str, dry_run: bool) -> None:
    if venv_path.exists():
        print(f"[INFO] Reusing existing virtual environment at {venv_path}")
        return
    print(f"[INFO] Creating virtual environment at {venv_path}")
    run_command([python, "-m", "venv", str(venv_path)], dry_run=dry_run)


def python_in_venv(venv_path: Path) -> Path:
    if os.name == "nt":
        return venv_path / "Scripts" / "python.exe"
    return venv_path / "bin" / "python"


def pip_in_venv(venv_path: Path) -> Path:
    if os.name == "nt":
        return venv_path / "Scripts" / "pip.exe"
    return venv_path / "bin" / "pip"


def install_dependencies(venv_path: Path, requirements: Path, *, dry_run: bool) -> None:
    python_exe = python_in_venv(venv_path)
    pip_exe = pip_in_venv(venv_path)
    run_command([str(python_exe), "-m", "pip", "install", "--upgrade", "pip"], dry_run=dry_run)
    run_command([str(pip_exe), "install", "-r", str(requirements)], dry_run=dry_run)


def run_combine_results(
    venv_path: Path,
    *,
    input_path: Path,
    subject_catalog: Path,
    biodata: Path,
    config: Path,
    outdir: Path,
    current_semester: int | None,
    extra_args: Sequence[str] | None,
    dry_run: bool,
) -> None:
    python_exe = python_in_venv(venv_path)
    cmd = [
        str(python_exe),
        str(PROJECT_ROOT / "combine_results.py"),
        "--input",
        str(input_path),
        "--subject-catalog",
        str(subject_catalog),
        "--biodata",
        str(biodata),
        "--outdir",
        str(outdir),
        "--config",
        str(config),
    ]
    if current_semester is not None:
        cmd.extend(["--current-semester", str(current_semester)])
    if extra_args:
        cmd.extend(extra_args)
    run_command(cmd, dry_run=dry_run)


def run_class_snapshot(
    venv_path: Path,
    *,
    overall_path: Path,
    output_path: Path,
    dry_run: bool,
    working_dir: Path,
) -> None:
    python_exe = python_in_venv(venv_path)
    cmd = [
        str(python_exe),
        str(PROJECT_ROOT / "generate_class_performance_summary.py"),
        "--overall",
        str(overall_path),
        "--output",
        str(output_path),
    ]
    # Run from the project root so local module imports (e.g. class_performance)
    # resolve the same way they do when invoking the script manually.
    display_cmd = " ".join(cmd)
    print(f"[CMD] {display_cmd}")
    if dry_run:
        return
    completed = subprocess.run(cmd, cwd=str(working_dir))
    if completed.returncode != 0:
        raise CommandError(
            f"Command failed with exit code {completed.returncode}: {display_cmd}"
        )


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)

    venv_path = Path(args.venv).resolve()
    requirements_path = Path(args.requirements).resolve()
    input_path = Path(args.input).resolve()
    subject_catalog = Path(args.subject_catalog).resolve()
    biodata = Path(args.biodata).resolve()
    config_path = Path(args.config).resolve()
    outdir = Path(args.outdir).resolve()

    try:
        check_files_exist([input_path, subject_catalog, biodata, config_path, requirements_path])
        ensure_virtualenv(venv_path, python=args.python, dry_run=args.dry_run)
        install_dependencies(venv_path, requirements_path, dry_run=args.dry_run)
        if not args.dry_run:
            outdir.mkdir(parents=True, exist_ok=True)
        run_combine_results(
            venv_path,
            input_path=input_path,
            subject_catalog=subject_catalog,
            biodata=biodata,
            config=config_path,
            outdir=outdir,
            current_semester=args.current_semester,
            extra_args=args.extra_combine_args,
            dry_run=args.dry_run,
        )
        if args.class_performance_output != "-":
            snapshot_path = Path(args.class_performance_output).resolve()
            overall_csv = outdir / "analytics_student_overall.csv"
            run_class_snapshot(
                venv_path,
                overall_path=overall_csv,
                output_path=snapshot_path,
                dry_run=args.dry_run,
                working_dir=PROJECT_ROOT,
            )
        print("[SUCCESS] Workflow completed successfully.")
        return 0
    except (FileNotFoundError, CommandError) as exc:
        print(f"[ERROR] {exc}")
        if not args.keep_venv and not args.dry_run:
            if venv_path.exists():
                print(f"[INFO] Removing virtual environment at {venv_path}")
                shutil.rmtree(venv_path, ignore_errors=True)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
