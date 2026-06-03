# MTE-SDWANAPI

Connects to a Cisco SD-WAN Manager (vManage), collects device and template data, and generates an Excel report.

---

## Project structure

```text
MTE-SDWANAPI/
  main.py               # Entry point — runs the workflow
  requirements.txt      # Python dependencies
  mte_sdwanapi/
    api.py              # API calls and data helpers
    report.py           # Excel report builder
    cli.py              # Command-line argument handling
  credentials.sh        # Environment variables (not committed)
```

---

## Requirements

- Python 3.9+
- A Cisco SD-WAN Manager instance reachable from your machine

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## Setup

- Create a `credentials.sh` file with your vManage details:

```bash
export VMANAGE_IP="192.168.1.1"
export VMANAGE_USER="admin"
export VMANAGE_PASSWORD="yourpassword"
export VMANAGE_PORT="443"        # optional, defaults to 443
```

- Source the file to load the variables into your shell:

```bash
source credentials.sh
```

---

## Usage

Run the report:

```bash
python main.py
```

Run with verbose output to see all collected API data printed to the terminal:

```bash
python main.py --verbose
```

Show available options:

```bash
python main.py --help
```

The report is saved as **`device_report.xlsx`** in the current directory.

---

## Report contents

| Sheet | Contents |
| --- | --- |
| reachable devices | All reachable edge devices with key attributes |
| One sheet per device template | Template summary, attached feature templates, and per-device variable values |

---

## Notes

- `credentials.sh` is excluded from version control via `.gitignore`. Never commit credentials.
- The report file `device_report.xlsx` is also excluded — regenerate it by running the script.
