"""Entry point: python -m dashboard (serves with waitress on PORT, default 8000)."""

import logging
import os

from waitress import serve

from .app import create_app

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")

if __name__ == "__main__":
    app = create_app()
    port = int(os.environ.get("PORT", "8000"))
    logging.getLogger("inventory-dashboard").info(
        "Serving on port %s, workbook %s", port, os.environ.get("WORKBOOK_PATH", "/data/Hardware Inventory.xlsx"))
    serve(app, host="0.0.0.0", port=port, threads=4)
