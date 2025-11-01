"""Entry point for the BOM check application."""
from __future__ import annotations

from pathlib import Path

from config_manager import ConfigManager
from gui import Application


def main() -> None:
    config_path = Path("config.json")
    manager = ConfigManager(config_path)
    app = Application(manager)
    app.mainloop()


if __name__ == "__main__":
    main()
