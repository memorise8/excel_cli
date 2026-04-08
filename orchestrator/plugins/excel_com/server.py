#!/usr/bin/env python3
"""HTTP bridge server — runs on Windows, accepts requests from Linux harness.

Usage (Windows):
    pip install pywin32 flask
    python server.py --port 8765

The server exposes Excel COM operations as REST endpoints.
The Linux harness calls these endpoints instead of local file manipulation.
"""

import argparse
import json
import os
import sys
from pathlib import Path

from flask import Flask, request, jsonify

from bridge import ExcelBridge

app = Flask(__name__)

# Active workbook sessions
sessions: dict[str, ExcelBridge] = {}


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "platform": sys.platform, "sessions": len(sessions)})


@app.route("/open", methods=["POST"])
def open_workbook():
    data = request.json
    path = data.get("path", "")
    session_id = data.get("session_id", "default")
    visible = data.get("visible", False)

    if not path or not Path(path).exists():
        return jsonify({"error": f"File not found: {path}"}), 400

    try:
        bridge = ExcelBridge(visible=visible)
        info = bridge.open(path)
        sessions[session_id] = bridge
        return jsonify({"session_id": session_id, **info})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/close", methods=["POST"])
def close_workbook():
    data = request.json
    session_id = data.get("session_id", "default")
    save = data.get("save", False)

    bridge = sessions.pop(session_id, None)
    if not bridge:
        return jsonify({"error": "Session not found"}), 404

    try:
        bridge.close(save=save)
        return jsonify({"status": "ok"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/save", methods=["POST"])
def save_workbook():
    data = request.json
    session_id = data.get("session_id", "default")
    path = data.get("path")

    bridge = sessions.get(session_id)
    if not bridge:
        return jsonify({"error": "Session not found"}), 404

    try:
        result = bridge.save(path)
        return jsonify(json.loads(result))
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/cell/get", methods=["POST"])
def get_cell():
    data = request.json
    session_id = data.get("session_id", "default")
    bridge = sessions.get(session_id)
    if not bridge:
        return jsonify({"error": "Session not found"}), 404

    try:
        result = bridge.get_cell(data["sheet"], data["cell"])
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/cell/set", methods=["POST"])
def set_cell():
    data = request.json
    session_id = data.get("session_id", "default")
    bridge = sessions.get(session_id)
    if not bridge:
        return jsonify({"error": "Session not found"}), 404

    try:
        result = bridge.set_cell(data["sheet"], data["cell"], data["value"])
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/range/get", methods=["POST"])
def get_range():
    data = request.json
    session_id = data.get("session_id", "default")
    bridge = sessions.get(session_id)
    if not bridge:
        return jsonify({"error": "Session not found"}), 404

    try:
        result = bridge.get_range(data["sheet"], data["range"])
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/range/set", methods=["POST"])
def set_range():
    data = request.json
    session_id = data.get("session_id", "default")
    bridge = sessions.get(session_id)
    if not bridge:
        return jsonify({"error": "Session not found"}), 404

    try:
        result = bridge.set_range(data["sheet"], data["range"], data["data"])
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/recalculate", methods=["POST"])
def recalculate():
    data = request.json
    session_id = data.get("session_id", "default")
    bridge = sessions.get(session_id)
    if not bridge:
        return jsonify({"error": "Session not found"}), 404

    try:
        result = bridge.recalculate()
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/formula/set", methods=["POST"])
def set_formula():
    data = request.json
    session_id = data.get("session_id", "default")
    bridge = sessions.get(session_id)
    if not bridge:
        return jsonify({"error": "Session not found"}), 404

    try:
        result = bridge.set_formula(data["sheet"], data["cell"], data["formula"])
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/formula/result", methods=["POST"])
def formula_result():
    data = request.json
    session_id = data.get("session_id", "default")
    bridge = sessions.get(session_id)
    if not bridge:
        return jsonify({"error": "Session not found"}), 404

    try:
        result = bridge.get_formula_result(data["sheet"], data["cell"])
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/pivot/create", methods=["POST"])
def create_pivot():
    data = request.json
    session_id = data.get("session_id", "default")
    bridge = sessions.get(session_id)
    if not bridge:
        return jsonify({"error": "Session not found"}), 404

    try:
        result = bridge.create_pivot(
            source_sheet=data["source_sheet"],
            source_range=data["source_range"],
            dest_sheet=data.get("dest_sheet", "PivotResult"),
            dest_cell=data.get("dest_cell", "A1"),
            name=data.get("name", "PivotTable1"),
            row_fields=data.get("row_fields", []),
            value_fields=data.get("value_fields", []),
            col_fields=data.get("col_fields"),
        )
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/pivot/refresh", methods=["POST"])
def refresh_pivot():
    data = request.json
    session_id = data.get("session_id", "default")
    bridge = sessions.get(session_id)
    if not bridge:
        return jsonify({"error": "Session not found"}), 404

    try:
        result = bridge.refresh_pivot(data["name"])
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/pivot/list", methods=["POST"])
def list_pivots():
    data = request.json
    session_id = data.get("session_id", "default")
    bridge = sessions.get(session_id)
    if not bridge:
        return jsonify({"error": "Session not found"}), 404

    try:
        result = bridge.list_pivots()
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/chart/create", methods=["POST"])
def create_chart():
    data = request.json
    session_id = data.get("session_id", "default")
    bridge = sessions.get(session_id)
    if not bridge:
        return jsonify({"error": "Session not found"}), 404

    try:
        result = bridge.create_chart(
            sheet=data["sheet"],
            source_range=data["source_range"],
            chart_type=data.get("chart_type", "column"),
            title=data.get("title", ""),
            dest_sheet=data.get("dest_sheet"),
        )
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/sheets", methods=["POST"])
def list_sheets():
    data = request.json
    session_id = data.get("session_id", "default")
    bridge = sessions.get(session_id)
    if not bridge:
        return jsonify({"error": "Session not found"}), 404

    try:
        result = bridge.list_sheets()
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/macro/run", methods=["POST"])
def run_macro():
    data = request.json
    session_id = data.get("session_id", "default")
    bridge = sessions.get(session_id)
    if not bridge:
        return jsonify({"error": "Session not found"}), 404

    try:
        result = bridge.run_macro(data["name"])
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/export/pdf", methods=["POST"])
def export_pdf():
    data = request.json
    session_id = data.get("session_id", "default")
    bridge = sessions.get(session_id)
    if not bridge:
        return jsonify({"error": "Session not found"}), 404

    try:
        result = bridge.export_pdf(
            output_path=data["output"],
            sheets=data.get("sheets"),
        )
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Excel COM Bridge Server (Windows)")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--host", default="0.0.0.0")
    args = parser.parse_args()

    print(f"Excel COM Bridge Server: http://localhost:{args.port}")
    print(f"Platform: {sys.platform}")
    print("Endpoints: /health, /open, /close, /cell/get, /cell/set, /range/get, /range/set")
    print("           /recalculate, /formula/set, /formula/result")
    print("           /pivot/create, /pivot/refresh, /pivot/list")
    print("           /chart/create, /sheets, /macro/run, /export/pdf")

    app.run(host=args.host, port=args.port, debug=False)
