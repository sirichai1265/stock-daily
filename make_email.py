"""Open an Outlook draft (not sent) for a day's report: body pasted from the model, HTML attached.

Usage: python make_email.py <M-D>        e.g. python make_email.py 10-22
Needs <M-D>/Stock_Daily_<M-D>.model.json and .html (produced by build_stock_daily.py).
"""
import importlib.util
import json
import subprocess
import sys
from pathlib import Path

FOLDER = Path(__file__).resolve().parent
TO = ("tblim@heungaline.com; shhwang@heungaline.com; logistics@heungaline.co.th; "
      "csoutbound@heungaline.co.th; sales@heungaline.co.th")

tag = sys.argv[1]
day = FOLDER / tag
model = json.loads((day / f"Stock_Daily_{tag}.model.json").read_text(encoding="utf-8"))
spec = importlib.util.spec_from_file_location("bsd", FOLDER / "build_stock_daily.py")
bsd = importlib.util.module_from_spec(spec)
spec.loader.exec_module(bsd)

body = day / f"Email_{tag}.html"
body.write_text(bsd.build_email_html(model), encoding="ascii")
d = bsd._md(model["date"]) + " " + model["date"][:4]
subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
                str(FOLDER / "make_outlook_draft.ps1"),
                "-HtmlBody", str(body), "-Attach", str(day / f"Stock_Daily_{tag}.html"),
                "-Subject", f"Daily Container Stock vs Booking - {d}", "-To", TO], check=True)
