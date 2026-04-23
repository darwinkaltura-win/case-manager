#!/usr/bin/env python3
"""
HAR file analyzer: shows HTTP errors and extracts x-me session headers.
Usage: python parse_har.py <path_to_file.har>
"""

import json
import sys
from pathlib import Path


def load_har(path: str) -> dict:
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        return json.load(f)


def find_header(headers: list, name: str) -> str | None:
    name_lower = name.lower()
    for h in headers:
        if h.get("name", "").lower() == name_lower:
            return h.get("value")
    return None


def find_headers_containing(headers: list, substring: str) -> dict:
    result = {}
    for h in headers:
        if substring.lower() in h.get("name", "").lower():
            result[h["name"]] = h.get("value", "")
    return result


def analyze(har_path: str):
    data = load_har(har_path)
    entries = data.get("log", {}).get("entries", [])

    errors = []
    x_me_sessions = {}  # url -> {header: value}

    for entry in entries:
        req = entry.get("request", {})
        resp = entry.get("response", {})
        url = req.get("url", "")
        status = resp.get("status", 0)

        # Collect errors (4xx / 5xx)
        if status >= 400:
            errors.append({
                "status": status,
                "method": req.get("method", ""),
                "url": url,
                "status_text": resp.get("statusText", ""),
            })

        # Extract x-me* headers from request and response
        req_xme = find_headers_containing(req.get("headers", []), "x-me")
        resp_xme = find_headers_containing(resp.get("headers", []), "x-me")

        combined = {**req_xme, **resp_xme}
        if combined:
            x_me_sessions[url] = combined

    # --- Print errors ---
    print(f"\n{'='*70}")
    print(f"HAR FILE: {har_path}")
    print(f"Total entries: {len(entries)}")
    print(f"{'='*70}")

    print(f"\n[ERRORS] ({len(errors)} requests with 4xx/5xx status)")
    print("-" * 70)
    if errors:
        for e in errors:
            print(f"  {e['status']} {e['status_text']:<20} {e['method']} {e['url']}")
    else:
        print("  No errors found.")

    # --- Print x-me sessions ---
    print(f"\n[X-ME SESSIONS] ({len(x_me_sessions)} URLs with x-me headers)")
    print("-" * 70)
    if x_me_sessions:
        seen_values = {}
        for url, headers in x_me_sessions.items():
            for hname, hval in headers.items():
                key = f"{hname}: {hval}"
                if key not in seen_values:
                    seen_values[key] = []
                seen_values[key].append(url)

        for key, urls in seen_values.items():
            print(f"\n  {key}")
            print(f"  Found in {len(urls)} request(s):")
            for u in urls[:5]:
                print(f"    {u}")
            if len(urls) > 5:
                print(f"    ... and {len(urls) - 5} more")
    else:
        print("  No x-me headers found.")

    print()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python parse_har.py <path_to_file.har>")
        sys.exit(1)
    analyze(sys.argv[1])
