"""Open an Outlook draft (not sent) for a day's report: report pasted into the user's own mail template, HTML attached.

Usage: python make_email.py <M-D>        e.g. python make_email.py 10-22
Needs <M-D>/Stock_Daily_<M-D>.model.json and .html (produced by build_stock_daily.py).

The mail body takes only two pieces of the user's Outlook template "STOCK DAILY ON BKK & LCH": the greeting
(rewritten to "Dear      All concerned,") and the "Thanks & Best regards ..." signature + company block. Its header
line and the EMPTY TO TH / HAL shared to SKR / EQ SPECIAL / INBOUND FULL / DEPOT STOCK BALANCE headings are dropped.
In between: a "Please find ..." line and the report. If the template file is missing, a plain wrapper is used.
"""
import importlib.util
import json
import os
import re
import subprocess
import sys
from pathlib import Path

FOLDER = Path(__file__).resolve().parent
TO = ("tblim@heungaline.com; shhwang@heungaline.com; logistics@heungaline.co.th; "
      "csoutbound@heungaline.co.th; sales@heungaline.co.th")
TEMPLATE = (Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Signatures"
            / "STOCK DAILY ON BKK & LCH (Sirichai@heungaline.co.th).htm")
P = "font-family:Calibri,sans-serif;font-size:11pt;color:#000000;margin:0"

tag = sys.argv[1]
day = FOLDER / tag
model = json.loads((day / f"Stock_Daily_{tag}.model.json").read_text(encoding="utf-8"))
spec = importlib.util.spec_from_file_location("bsd", FOLDER / "build_stock_daily.py")
bsd = importlib.util.module_from_spec(spec)
spec.loader.exec_module(bsd)

d = bsd._md(model["date"]) + " " + model["date"][:4]
report = bsd.build_email_html(model)
inner = re.search(r"<body[^>]*>(.*)</body>", report, re.S).group(1)
please = (f"<p style='{P}'>Please find the daily container stock vs booking as of {d} below "
          "(HTML file attached).</p><p style='" + P + "'>&nbsp;</p>")
block = please + inner

if TEMPLATE.exists():
    tpl = TEMPLATE.read_bytes().decode("utf-8", "replace")
    dear = re.search(r"<p[^>]*>(?:(?!</p>).)*?Dear.*?</p>", tpl, re.S)
    thanks = re.search(r"<p[^>]*>(?:(?!</p>).)*?Thanks", tpl, re.S)
    if not dear or not thanks:
        raise SystemExit("template has no Dear / Thanks paragraph - not creating the draft")
    dear_p = dear.group(0).replace("CS Outbound and Sale team,", "All concerned,")
    # template's header line and the EMPTY TO TH ... DEPOT STOCK BALANCE headings are dropped on purpose
    body = dear_p + "<p style='" + P + "'>&nbsp;</p>" + block + "<p style='" + P + "'>&nbsp;</p>" + tpl[thanks.start():]
else:
    body = (f"<p style='{P}'>Dear&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; All concerned,</p><p style='{P}'>&nbsp;</p>{block}"
            f"<p style='{P}'>&nbsp;</p><p style='{P}'>Thanks &amp; Best regards,<br>"
            "Mr. Sirichai Prayochananon / Operation team</p>")

page = ("<html><head><meta charset='utf-8'></head><body>" + body + "</body></html>")
out = day / f"Email_{tag}.html"
out.write_text(page.encode("ascii", "xmlcharrefreplace").decode("ascii"), encoding="ascii")
subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
                str(FOLDER / "make_outlook_draft.ps1"),
                "-HtmlBody", str(out), "-Attach", str(day / f"Stock_Daily_{tag}.html"),
                "-Subject", f"Daily Container Stock vs Booking - {d}", "-To", TO], check=True)
