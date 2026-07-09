#!/usr/bin/env python3
"""
Production entrypoint for the HR portal, meant to be kept running
continuously (as a Windows service, systemd unit, or Docker container) on
whatever always-on machine hosts it on the office LAN.

Uses waitress instead of Flask's dev server since it's safe for unattended,
long-running use. See README.md for how to run this as a persistent service.
"""

import socket

from waitress import serve

from portal_app import app

HOST = "0.0.0.0"
PORT = 5000


def _local_ip_hint():
    """Best-effort guess at the LAN IP, just for the startup log message."""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        return s.getsockname()[0]
    except OSError:
        return "127.0.0.1"
    finally:
        s.close()


if __name__ == "__main__":
    ip = _local_ip_hint()
    print(f"HR Portal starting on http://{ip}:{PORT} (reachable from any device on this LAN)")
    print(f"Local access: http://localhost:{PORT}")
    serve(app, host=HOST, port=PORT)
