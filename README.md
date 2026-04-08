# hr-admission-automation
Hybrid automation pipeline for onboarding data validation — combines RPA browser automation and REST APIs (ViaCEP, BrasilAPI) to validate addresses, bank codes, and national holidays, generating a prioritized output report from Excel input.

---

## Features

- Reads employee spreadsheet (`entrada/colaboradores.xlsx`)
- ZIP code validation via **RPA** (automated browser navigation on consultarcep.com.br)
- ZIP code validation via **REST API** (ViaCEP)
- Cross-validation between RPA and API — divergent records are flagged as inconsistent
- Bank data validation via **Brasil API**
- National holiday check with **per-year cache** (avoids repetitive API calls)
- Generates `Prioridade_Agi` column with applied business rules
- Final report exported to `saida/resultado_final.xlsx`

---

## Priority Rules

| Priority | Condition |
|---|---|
| `ALTA` | ZIP code and bank successfully validated through both channels (RPA and API) |
| `BLOQUEADO` | Admission date falls on a national holiday |
| `BAIXA` | Any validation error, RPA/API divergence, or missing data |

---

## Tech Stack

- **Python 3.x**
- **Pandas** — data manipulation and Excel handling
- **Requests** — REST API consumption
- **Selenium / Playwright** — browser automation (RPA)
- **Openpyxl** — reading and writing `.xlsx` files
- **WebDriver Manager** — automatic browser driver management

### APIs Used

| API | Purpose |
|---|---|
| [ViaCEP](https://viacep.com.br) | ZIP code address validation |
| [consultarcep.com.br](https://www.consultarcep.com.br) | ZIP code validation via RPA |
| [Brasil API — Banks](https://brasilapi.com.br/docs#tag/BANKS) | Bank code validation |
| [Brasil API — Holidays](https://brasilapi.com.br/docs#tag/Feriados) | National holiday lookup |

---

## Project Structure

```
hr-data-pipeline/
│
├── entrada/
│   └── colaboradores.xlsx       # Input spreadsheet with employee data
│
├── saida/
│   └── resultado_final.xlsx     # Generated report with validations and priority
│
├── main.py                      # Main script
├── requirements.txt             # Project dependencies
└── README.md
```

---

## How to Run

### 1. Clone the repository

```bash
git clone git clone https://github.com/MajuFantagussi/hr-admission-automation.git
cd hr-data-pipeline
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Add the input spreadsheet

Place the `colaboradores.xlsx` file inside the `entrada/` folder.

### 4. Run the pipeline

```bash
python main.py
```

The report will be generated at `saida/resultado_final.xlsx`.

---

## Dependencies

```
pandas
requests
openpyxl
selenium
playwright
webdriver-manager
```

---

## Notes

- The script performs a **single holiday API call per year**, caching the result to avoid repetitive requests per processed row.
- If any validation fails on either channel (RPA or API), the record is automatically classified as `BAIXA`.

