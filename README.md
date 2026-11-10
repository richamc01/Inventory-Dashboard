# Inventory Dashboard

This is a read-only web dashboard for the hardware lot. It reads `Hardware Inventory.xlsx` and the `inventory-raw` scan files straight from the Nextcloud share. It never writes to the workbook.

## What it shows

| Tab | Contents |
|---|---|
| Overview | Net sales, how much of the lot cost is recovered, profit or amount still to recover, net sales potential, total profit potential, units sold, average net, machines left, cumulative net vs lot cost, sales by machine type, platform, payment method and model, and a list of rows that need attention |
| Sales | Every sale with search, filters, sortable columns, totals, multi-machine orders grouped together, and daily totals |
| Listings | Every machine and item with its listing status (Sold, Pending, Listed, Needs listed, Not listed, Removed), asking price, date listed and platform, plus net sales potential and total profit potential. Search, filters and sortable columns. |
| Inventory | Every machine, including Master-only machines that were never scanned, with search, filters and sorting. Click a machine to see every scanned field, its Master notes, its sale, and the full scan file split into sections. Every scan for an asset is listed. |
| Fleet | Breakdowns of unsold or all machines by model, CPU, generation, RAM, Windows 11, TPM, drive, type and status. Also groups identical machines for writing listings. |
| Other hardware | The Items tab (RAM grouped by kind) and the Bulk tab |
| Notes | The Master tab and the attention list |
| Data | File details, warnings, a cross-check of totals against the Selling Dashboard tab, and the Summary and Lists tabs |

The page checks every 60 seconds and reloads on its own when the workbook changes.

### How the numbers are worked out

- The Computers tab's columns U to Y (`win11_override`, `manual_notes`, `win11_effective`, `ram_mb_effective`, `sold`) are recomputed from the Master and Sales tabs. The rules are the same ones in the workbook README, section A, step 10. This keeps the values right even when a paste has wiped one of those formulas.
- The Sales tab's calculated columns (Q to X) come from the values Excel saved. If Excel saved no results, the dashboard calculates them itself using the same rules.
- The Listings tab is optional. Each machine (Computers and Master tabs) and each item (Items tab, matched by Asset ID or Serial) appears once. A unit on the Sales tab is Sold, whatever its Listings row says.
- **Net sales potential** is the actual net of everything sold plus the asking price of every unsold unit whose listing status is Listed or Pending. No fees or shipping are taken off asking prices. **Total profit potential** is net sales potential minus the lot price. A sold unit counts once, at its sale net, never at its asking price.
- The **Needs attention** list flags the following:
  - Sales rows whose Check is not OK
  - Machines whose scan status is not OK
  - Machines that appear only on the Master tab
  - Master notes that mention the scan row
  - Rows where a formula is missing in Excel
  - Saved values that are out of date
  - Totals that don't match the Selling Dashboard tab
  - Listings problems: the same ID on two rows (only the first is used), an ID that is not in the inventory, Listed or Pending with no asking price, and a status that is not on the list

## Install on UnRAID

Run the commands below in the UnRAID terminal (**Tools** > **Terminal**, or the `>_` icon at the top right of the web UI).

### 1. Create a read-only GitHub token (one time)

The repository is private, so cloning it needs a token.

1. On GitHub, click your avatar (top right) > **Settings**.
1. In the left sidebar, click **Developer settings** at the bottom.
1. Click **Personal access tokens** > **Fine-grained tokens**.
1. Click **Generate new token**.
1. **Token name:** `unraid-inventory-dashboard`.
1. **Expiration:** pick one.
1. **Repository access:** select **Only select repositories**, then pick `richamc01/Inventory-Dashboard`.
1. **Permissions:** click **Add permissions**, choose **Contents**, and leave it on **Read-only**.
1. Click **Generate token** and copy it. GitHub shows it only once.

### 2. Clone the repository

```bash
git clone https://github.com/richamc01/Inventory-Dashboard.git /mnt/user/appdata/inventory-dashboard
```

At the `Username` prompt, type `richamc01`. At the `Password` prompt, paste the token. Nothing is saved, so you'll paste the token again for each update.

### 3. Check that port 8095 is free

```bash
docker ps --format '{{.Names}}  {{.Ports}}' | grep 8095
```

No output means the port is free. If something is listed, choose another port and put `HOST_PORT=<port>` in front of `bash deploy.sh` in step 4.

### 4. Build and start

```bash
cd /mnt/user/appdata/inventory-dashboard
bash deploy.sh
```

The script does the following:

1. Builds the image.
1. Checks that the container user (`99:100`, nobody:users) can read the workbook.
1. Starts the container with `--restart unless-stopped`.
1. Mounts the Inventory folder **read-only** at `/data`.
1. Prints the status.

It ends with `== Done.`

### 5. Open it

Browse to `http://<capacitor's LAN IP>:8095`.

The container appears on UnRAID's **Docker** tab as `inventory-dashboard`. You can start, stop and view logs there. It was created from the command line, so it has no edit template.

### If the read check in step 4 fails

If you see `Permission denied`, the container user can't reach the file through the Nextcloud folders. Run this and send the output:

```bash
namei -l "/mnt/user/Nextcloud/richamc01/files/Documents/Personal Assets/Home Network/Inventory/Hardware Inventory.xlsx"
```

## Update to a new version

```bash
cd /mnt/user/appdata/inventory-dashboard
git pull
bash deploy.sh
```

## Settings

`deploy.sh` reads these environment variables:

| Variable | Default |
|---|---|
| `HOST_PORT` | `8095` |
| `DATA_DIR` | `/mnt/user/Nextcloud/richamc01/files/Documents/Personal Assets/Home Network/Inventory` |
| `RUN_AS` | `99:100` |

These are set inside the container:

| Variable | Default |
|---|---|
| `WORKBOOK_PATH` | `/data/Hardware Inventory.xlsx` |
| `RAW_DIR` | `inventory-raw` next to the workbook |
| `PORT` | `8000` |

`docker-compose.yml` is included for UnRAID's Compose Manager plugin. You don't need it if you use `deploy.sh`.

## Notes

- LAN only. There is no login.
- The workbook is copied into memory before it is read. If a read fails partway through a Nextcloud sync, the dashboard keeps showing the last good version with an error banner, and retries on the next check.
- `.gitignore` blocks `.xlsx`, `.csv` and `inventory-raw/`, so inventory data can't be committed. Tests use a fake workbook built by `tests/make_fixture.py`.

## Development

```bash
pip install -r requirements.txt pytest
python -m pytest -q
REAL_WORKBOOK="/path/to/Hardware Inventory.xlsx" python -m pytest -q   # also check the real file
WORKBOOK_PATH="/path/to/Hardware Inventory.xlsx" python -m dashboard   # http://localhost:8000
```

GitHub Actions runs the tests on every push. It also builds the Docker image, runs it as `99:100` against the fake workbook, and runs `deploy.sh`.
