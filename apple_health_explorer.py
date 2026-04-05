from __future__ import annotations

from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
from pathlib import Path
from urllib.parse import parse_qs, urlparse
import xml.etree.ElementTree as ET


PROJECT_ROOT = Path(__file__).resolve().parent
EXPORT_XML_PATH = PROJECT_ROOT / "apple_health_export" / "export.xml"
CACHE_PATH = PROJECT_ROOT / ".apple_health_explorer_cache.json"
DEFAULT_PORT = 8765
MAX_RESULTS = 500


HTML_PAGE = """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Apple Health Explorer</title>
  <style>
    :root {
      --bg: #f3efe7;
      --panel: #fffdf8;
      --ink: #1d2a33;
      --muted: #65737e;
      --accent: #0b6e4f;
      --line: #d9d1c3;
      --chip: #e9f4ef;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: Georgia, "Times New Roman", serif;
      color: var(--ink);
      background:
        radial-gradient(circle at top left, rgba(11,110,79,.08), transparent 30%),
        linear-gradient(180deg, #f8f4ec 0%, var(--bg) 100%);
    }
    .wrap {
      max-width: 1280px;
      margin: 0 auto;
      padding: 24px;
    }
    .hero {
      display: grid;
      gap: 10px;
      margin-bottom: 22px;
    }
    h1 {
      margin: 0;
      font-size: clamp(2rem, 4vw, 3.5rem);
      line-height: 1;
      letter-spacing: -0.03em;
    }
    .sub {
      color: var(--muted);
      max-width: 900px;
      font-size: 1rem;
    }
    .grid {
      display: grid;
      grid-template-columns: 320px 1fr;
      gap: 20px;
      align-items: start;
    }
    .card {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 18px;
      padding: 18px;
      box-shadow: 0 10px 30px rgba(41, 51, 57, 0.06);
    }
    .card h2 {
      margin: 0 0 14px;
      font-size: 1.1rem;
    }
    label {
      display: block;
      font-size: .88rem;
      color: var(--muted);
      margin-bottom: 6px;
    }
    input, select, button {
      width: 100%;
      border-radius: 12px;
      border: 1px solid var(--line);
      padding: 11px 12px;
      font: inherit;
      background: white;
      color: var(--ink);
    }
    button {
      background: var(--accent);
      color: white;
      border: 0;
      cursor: pointer;
      font-weight: 600;
    }
    button:hover { opacity: .94; }
    .field { margin-bottom: 14px; }
    .meta {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-bottom: 16px;
    }
    .chip {
      background: var(--chip);
      color: var(--accent);
      border-radius: 999px;
      padding: 7px 11px;
      font-size: .85rem;
      font-weight: 600;
    }
    .results-head {
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: center;
      margin-bottom: 14px;
      flex-wrap: wrap;
    }
    .status {
      color: var(--muted);
      font-size: .92rem;
    }
    .table-wrap {
      overflow: auto;
      border: 1px solid var(--line);
      border-radius: 14px;
      background: white;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: .9rem;
    }
    th, td {
      text-align: left;
      padding: 10px 12px;
      border-bottom: 1px solid #eee7db;
      vertical-align: top;
      white-space: nowrap;
    }
    th {
      position: sticky;
      top: 0;
      background: #fcfaf5;
      z-index: 1;
    }
    tr:hover td { background: #fbf8f1; }
    .muted { color: var(--muted); }
    .small { font-size: .82rem; }
    @media (max-width: 960px) {
      .grid { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <div class="wrap">
    <section class="hero">
      <h1>Apple Health Export Explorer</h1>
      <div class="sub">Browse raw Apple Health records directly from <code>export.xml</code>. Filter by record type, date range, source name, and result limit to understand what the export actually contains.</div>
    </section>

    <div class="grid">
      <aside class="card">
        <h2>Filters</h2>
        <div class="field">
          <label for="recordType">Record type</label>
          <select id="recordType"></select>
        </div>
        <div class="field">
          <label for="startDate">Start date</label>
          <input id="startDate" type="datetime-local">
        </div>
        <div class="field">
          <label for="endDate">End date</label>
          <input id="endDate" type="datetime-local">
        </div>
        <div class="field">
          <label for="sourceFilter">Source contains</label>
          <input id="sourceFilter" type="text" placeholder="Apple Watch">
        </div>
        <div class="field">
          <label for="limit">Limit</label>
          <input id="limit" type="number" min="1" max="500" value="100">
        </div>
        <button id="loadButton">Load records</button>
        <button id="refreshCacheButton" style="margin-top:10px;background:#c96d2d">Update cache</button>
        <div class="small muted" style="margin-top:12px">
          Date filters are interpreted in local browser time and matched against Apple Health record start/end dates.
        </div>
      </aside>

      <main class="card">
        <div class="results-head">
          <div>
            <h2 style="margin:0 0 4px">Results</h2>
            <div class="status" id="status">Loading record catalog…</div>
          </div>
          <div class="meta" id="summary"></div>
        </div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Type</th>
                <th>Value</th>
                <th>Unit</th>
                <th>Start</th>
                <th>End</th>
                <th>Source</th>
                <th>Metadata</th>
              </tr>
            </thead>
            <tbody id="rows"></tbody>
          </table>
        </div>
      </main>
    </div>
  </div>

  <script>
    const state = { recordTypes: [] };

    async function fetchJson(url, options = {}) {
      const response = await fetch(url, options);
      if (!response.ok) {
        const text = await response.text();
        throw new Error(text || `Request failed: ${response.status}`);
      }
      return response.json();
    }

    function fmtMeta(metadata) {
      if (!metadata || !metadata.length) return "";
      return metadata.map(item => `${item.key}=${item.value}`).join(" | ");
    }

    function buildSummary(items) {
      const summary = document.getElementById("summary");
      summary.innerHTML = "";
      items.forEach(text => {
        const chip = document.createElement("div");
        chip.className = "chip";
        chip.textContent = text;
        summary.appendChild(chip);
      });
    }

    function renderRows(records) {
      const tbody = document.getElementById("rows");
      tbody.innerHTML = "";
      for (const record of records) {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${record.type || ""}</td>
          <td>${record.value || ""}</td>
          <td>${record.unit || ""}</td>
          <td>${record.startDate || ""}</td>
          <td>${record.endDate || ""}</td>
          <td>${record.sourceName || ""}</td>
          <td class="small">${fmtMeta(record.metadata)}</td>
        `;
        tbody.appendChild(tr);
      }
    }

    async function loadRecordTypes() {
      const data = await fetchJson("/api/record-types");
      state.recordTypes = data.record_types;
      const select = document.getElementById("recordType");
      select.innerHTML = `<option value="">Choose a record type…</option>`;
      for (const item of state.recordTypes) {
        const option = document.createElement("option");
        option.value = item.type;
        option.textContent = `${item.type} (${item.count})`;
        select.appendChild(option);
      }
      document.getElementById("status").textContent = "Pick a record type and load some rows.";
      buildSummary([
        `${data.record_types.length} record types`,
        `${data.metadata_key_count} metadata keys`,
        data.cache_updated_at ? `cached ${data.cache_updated_at}` : "fresh scan"
      ]);
    }

    async function refreshCache() {
      document.getElementById("status").textContent = "Refreshing cache from export.xml…";
      const data = await fetchJson("/api/refresh-cache", { method: "POST" });
      document.getElementById("status").textContent = "Cache refreshed.";
      buildSummary([
        `${data.record_types.length} record types`,
        `${data.metadata_key_count} metadata keys`,
        `cached ${data.cache_updated_at}`
      ]);
      const select = document.getElementById("recordType");
      const current = select.value;
      select.innerHTML = `<option value="">Choose a record type…</option>`;
      for (const item of data.record_types) {
        const option = document.createElement("option");
        option.value = item.type;
        option.textContent = `${item.type} (${item.count})`;
        if (item.type === current) option.selected = true;
        select.appendChild(option);
      }
    }

    async function loadRecords() {
      const type = document.getElementById("recordType").value;
      if (!type) {
        document.getElementById("status").textContent = "Select a record type first.";
        return;
      }

      const params = new URLSearchParams();
      params.set("type", type);
      const startDate = document.getElementById("startDate").value;
      const endDate = document.getElementById("endDate").value;
      const source = document.getElementById("sourceFilter").value.trim();
      const limit = document.getElementById("limit").value;
      if (startDate) params.set("start", startDate);
      if (endDate) params.set("end", endDate);
      if (source) params.set("source", source);
      if (limit) params.set("limit", limit);

      document.getElementById("status").textContent = "Scanning export.xml…";
      const data = await fetchJson(`/api/records?${params.toString()}`);
      renderRows(data.records);
      document.getElementById("status").textContent =
        `Showing ${data.records.length} of ${data.matched_count} matching records.`;
      buildSummary([
        data.record_type,
        `${data.records.length} shown`,
        `${data.matched_count} matched`,
        `limit ${data.limit}`
      ]);
    }

    document.getElementById("loadButton").addEventListener("click", () => {
      loadRecords().catch(error => {
        document.getElementById("status").textContent = error.message;
      });
    });

    document.getElementById("refreshCacheButton").addEventListener("click", () => {
      refreshCache().catch(error => {
        document.getElementById("status").textContent = error.message;
      });
    });

    loadRecordTypes().catch(error => {
      document.getElementById("status").textContent = error.message;
    });
  </script>
</body>
</html>
"""


def parse_export_datetime(value: str | None) -> datetime | None:
    if not value:
        return None
    return datetime.strptime(value, "%Y-%m-%d %H:%M:%S %z")


def parse_ui_datetime(value: str | None) -> datetime | None:
    if not value:
        return None
    parsed = datetime.fromisoformat(value)
    return parsed.astimezone()


def load_record_type_summary() -> list[dict[str, object]]:
    counts: dict[str, int] = {}
    for _, elem in ET.iterparse(EXPORT_XML_PATH, events=("end",)):
        if elem.tag == "Record":
            record_type = elem.attrib.get("type")
            if record_type:
                counts[record_type] = counts.get(record_type, 0) + 1
        elem.clear()

    return [
        {"type": record_type, "count": count}
        for record_type, count in sorted(counts.items())
    ]


def count_metadata_keys() -> int:
    keys: set[str] = set()
    for _, elem in ET.iterparse(EXPORT_XML_PATH, events=("end",)):
        if elem.tag == "MetadataEntry":
            key = elem.attrib.get("key")
            if key:
                keys.add(key)
        elem.clear()
    return len(keys)


def build_catalog_payload() -> dict[str, object]:
    payload = {
        "record_types": load_record_type_summary(),
        "metadata_key_count": count_metadata_keys(),
        "cache_updated_at": datetime.now().astimezone().strftime("%Y-%m-%d %H:%M:%S %Z"),
    }
    CACHE_PATH.write_text(json.dumps(payload), encoding="utf-8")
    return payload


def load_catalog_payload() -> dict[str, object]:
    if CACHE_PATH.exists():
        try:
            return json.loads(CACHE_PATH.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            pass
    return build_catalog_payload()


def query_records(
    record_type: str,
    start: datetime | None,
    end: datetime | None,
    source_contains: str | None,
    limit: int,
) -> tuple[list[dict[str, object]], int]:
    records: list[dict[str, object]] = []
    matched_count = 0
    source_filter = source_contains.lower() if source_contains else None

    for _, elem in ET.iterparse(EXPORT_XML_PATH, events=("end",)):
        if elem.tag != "Record" or elem.attrib.get("type") != record_type:
            elem.clear()
            continue

        record_start = parse_export_datetime(elem.attrib.get("startDate"))
        record_end = parse_export_datetime(elem.attrib.get("endDate"))
        if start and record_end and record_end < start:
            elem.clear()
            continue
        if end and record_start and record_start > end:
            elem.clear()
            continue
        if source_filter:
            source_name = elem.attrib.get("sourceName", "").replace("\xa0", " ")
            if source_filter not in source_name.lower():
                elem.clear()
                continue

        matched_count += 1
        if len(records) < limit:
            records.append(
                {
                    "type": elem.attrib.get("type", ""),
                    "value": elem.attrib.get("value", ""),
                    "unit": elem.attrib.get("unit", ""),
                    "startDate": elem.attrib.get("startDate", ""),
                    "endDate": elem.attrib.get("endDate", ""),
                    "sourceName": elem.attrib.get("sourceName", "").replace("\xa0", " "),
                    "metadata": [
                        {
                            "key": child.attrib.get("key", ""),
                            "value": child.attrib.get("value", ""),
                        }
                        for child in elem
                        if child.tag == "MetadataEntry"
                    ],
                }
            )
        elem.clear()

    return records, matched_count


class ExplorerHandler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self.respond_html(HTML_PAGE)
            return

        if parsed.path == "/api/record-types":
            self.respond_json(load_catalog_payload())
            return

        if parsed.path == "/api/records":
            query = parse_qs(parsed.query)
            record_type = query.get("type", [""])[0]
            if not record_type:
                self.respond_json({"error": "type is required"}, status=HTTPStatus.BAD_REQUEST)
                return

            try:
                start = parse_ui_datetime(query.get("start", [""])[0] or None)
                end = parse_ui_datetime(query.get("end", [""])[0] or None)
                limit = min(int(query.get("limit", [str(100)])[0]), MAX_RESULTS)
            except ValueError as exc:
                self.respond_json({"error": str(exc)}, status=HTTPStatus.BAD_REQUEST)
                return

            source_contains = query.get("source", [""])[0] or None
            records, matched_count = query_records(
                record_type=record_type,
                start=start,
                end=end,
                source_contains=source_contains,
                limit=max(1, limit),
            )
            self.respond_json(
                {
                    "record_type": record_type,
                    "records": records,
                    "matched_count": matched_count,
                    "limit": max(1, limit),
                }
            )
            return

        self.respond_json({"error": "not found"}, status=HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        if parsed.path == "/api/refresh-cache":
            self.respond_json(build_catalog_payload())
            return

        self.respond_json({"error": "not found"}, status=HTTPStatus.NOT_FOUND)

    def log_message(self, format: str, *args: object) -> None:
        return

    def respond_html(self, content: str, status: HTTPStatus = HTTPStatus.OK) -> None:
        body = content.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def respond_json(self, payload: dict[str, object], status: HTTPStatus = HTTPStatus.OK) -> None:
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def main() -> None:
    if not EXPORT_XML_PATH.exists():
        raise FileNotFoundError(f"Missing export.xml: {EXPORT_XML_PATH}")

    server = ThreadingHTTPServer(("127.0.0.1", DEFAULT_PORT), ExplorerHandler)
    print(f"Apple Health Explorer running at http://127.0.0.1:{DEFAULT_PORT}")
    print("Press Ctrl+C to stop.")
    server.serve_forever()


if __name__ == "__main__":
    main()
