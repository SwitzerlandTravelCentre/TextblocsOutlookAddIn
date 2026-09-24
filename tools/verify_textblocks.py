"""Compare the deployed JSON with this run's exact candidate without logging content."""
import hashlib
from pathlib import Path
import time
import requests

expected = hashlib.sha256(Path("dist/data/textblocks.json").read_bytes()).hexdigest()
url = "https://calm-rock-03c40f810.7.azurestaticapps.net/data/textblocks.json"
for attempt in range(12):
    response = requests.get(url, params={"release": expected}, headers={"Cache-Control": "no-cache"}, timeout=30)
    if response.status_code == 200 and hashlib.sha256(response.content).hexdigest() == expected:
        print("Published JSON matches the tested candidate: %s (%d text blocks)" % (expected, response.json()["blockCount"]))
        break
    if attempt < 11:
        time.sleep(10)
else:
    raise SystemExit("Published JSON differs from the candidate. Verify the release and use the documented rollback.")
