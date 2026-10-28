"""Executable shim that starts the JSON-line Python engine."""

from excel_engine.main import main


if __name__ == "__main__":
    # Return the engine's exit status to the operating system.
    raise SystemExit(main())
