"""Flask app: serves the dashboard page and a read-only JSON API."""

from __future__ import annotations

import datetime as dt
import logging
import os
import threading

from flask import Flask, Response, abort, jsonify, request, send_from_directory

from .workbook import build_snapshot

log = logging.getLogger("inventory-dashboard")

STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")


class SnapshotCache:
    """Re-reads the workbook only when its size or modified time changes.

    If a new version cannot be read, the last good snapshot keeps being served
    with an error attached, so a half-synced file never blanks the page.
    """

    def __init__(self, path, raw_dir):
        self.path = path
        self.raw_dir = raw_dir
        self._lock = threading.Lock()
        self._sig = None
        self._snap = None
        self._error = None
        self._error_at = None

    def _signature(self):
        st = os.stat(self.path)
        return (st.st_mtime_ns, st.st_size)

    def get(self):
        with self._lock:
            try:
                sig = self._signature()
            except OSError as e:
                self._error = f"Cannot open the workbook at {self.path}: {e.strerror or e}"
                self._error_at = _now()
                return self._snap, self._error
            if sig != self._sig or self._snap is None:
                try:
                    self._snap = build_snapshot(self.path, self.raw_dir)
                    self._sig = sig
                    self._error = None
                except Exception as e:  # noqa: BLE001 - report any read failure
                    log.exception("Failed to read workbook")
                    self._error = f"Could not read the workbook: {type(e).__name__}: {e}"
                    self._error_at = _now()
            return self._snap, self._error


def _now():
    return dt.datetime.now(dt.timezone.utc).isoformat()


def create_app(workbook_path=None, raw_dir=None):
    workbook_path = workbook_path or os.environ.get(
        "WORKBOOK_PATH", "/data/Hardware Inventory.xlsx")
    if raw_dir is None:
        raw_dir = os.environ.get("RAW_DIR") or os.path.join(
            os.path.dirname(workbook_path), "inventory-raw")

    app = Flask(__name__, static_folder=None)
    app.config["JSON_SORT_KEYS"] = False
    app.json.sort_keys = False
    cache = SnapshotCache(workbook_path, raw_dir)
    app.extensions["snapshot_cache"] = cache

    @app.after_request
    def no_cache(resp):
        if request.path.startswith("/api/"):
            resp.headers["Cache-Control"] = "no-store"
        resp.headers["X-Content-Type-Options"] = "nosniff"
        return resp

    @app.get("/")
    def index():
        return send_from_directory(STATIC_DIR, "index.html")

    @app.get("/static/<path:name>")
    def static_files(name):
        return send_from_directory(STATIC_DIR, name)

    @app.get("/api/data")
    def data():
        snap, err = cache.get()
        if snap is None:
            return jsonify({"error": err or "Workbook not loaded."}), 503
        out = dict(snap)
        out["meta"] = dict(snap["meta"])
        out["meta"]["error"] = err
        return jsonify(out)

    @app.get("/api/status")
    def status():
        snap, err = cache.get()
        return jsonify({
            "ok": snap is not None and err is None,
            "mtime": snap["meta"]["mtime"] if snap else None,
            "error": err,
        })

    @app.get("/api/raw/<asset_id>")
    def raw(asset_id):
        snap, _ = cache.get()
        if snap is None:
            abort(503)
        comp = next((c for c in snap["computers"] if c["asset_id"] == asset_id), None)
        if comp is None or not comp.get("scan_files"):
            abort(404)
        # Only file names already listed for this asset are served, so no
        # request can reach a path outside the scan folder.
        name = request.args.get("file") or comp.get("detail_file") or comp["scan_files"][-1]
        if name not in comp["scan_files"]:
            abort(404)
        path = os.path.join(cache.raw_dir, name)
        try:
            with open(path, "r", encoding="utf-8", errors="replace") as f:
                text = f.read()
        except OSError:
            abort(404)
        return Response(text, mimetype="text/plain; charset=utf-8",
                        headers={"Cache-Control": "no-store"})

    @app.get("/healthz")
    def healthz():
        return "ok\n", 200, {"Content-Type": "text/plain"}

    return app
