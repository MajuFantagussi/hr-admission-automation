# hr-admission-automation

Hybrid automation pipeline for onboarding data validation — combines RPA browser automation and REST APIs (ViaCEP, BrasilAPI) to validate addresses, bank codes, and national holidays, generating a prioritized output report from Excel input.

---

## 💡 Business Impact

This project simulates a real-world HR onboarding validation pipeline, ensuring:

* Data consistency across multiple validation sources
* Reduction of manual validation effort
* Early detection of inconsistencies in employee data
* Prioritization logic to support operational decision-making

---

## 🚀 Features

* Reads employee spreadsheet (`entrada/colaboradores.xlsx`)
* ZIP code validation via **RPA** (automated browser navigation on consultarcep.com.br)
* ZIP code validation via **REST API** (ViaCEP)
* Cross-validation between RPA and API — divergent records are flagged as inconsistent
* Bank data validation via **Brasil API**
* National holiday check with **per-year cache** (avoids repetitive API calls)
* Generates `Prioridade_Agi` column with applied business rules
* Final report exported to `saida/resultado_final.xlsx`

---

## 📊 Priority Rules

| Priority    | Condition                                                                    |
| ----------- | ---------------------------------------------------------------------------- |
| `ALTA`      | ZIP code and bank successfully validated through both channels (RPA and API) |
| `BLOQUEADO` | Admission date falls on a national holiday                                   |
| `BAIXA`     | Any validation error, RPA/API divergence, or missing data                    |

---

## 🧰 Prerequisites

Make sure you have the following installed:

* Python 3.10+
* pip (Python package manager)

To verify:

```bash
python --version
py -m pip --version
```

---

## 🛠️ Tech Stack

* **Python 3.x**
* **Pandas** — data manipulation and Excel handling
* **Requests** — REST API consumption
* **Selenium / Playwright** — browser automation (RPA)
* **Openpyxl** — reading and writing `.xlsx` files
* **WebDriver Manager** — automatic browser driver management

### APIs Used

| API                   | Purpose                     |
| --------------------- | --------------------------- |
| ViaCEP                | ZIP code address validation |
| consultarcep.com.br   | ZIP code validation via RPA |
| Brasil API — Banks    | Bank code validation        |
| Brasil API — Holidays | National holiday lookup     |

---

## 📁 Project Structure

hr-admission-automation/
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

---

## ▶️ How to Run

### 1. Clone the repository

```bash
git clone https://github.com/MajuFantagussi/hr-admission-automation.git
cd hr-admission-automation
```

---

### 2. Create a virtual environment (recommended)

```bash
py -m venv venv
venv\Scripts\activate
```

### 3. Install dependencies

```bash
py -m pip install -r requirements.txt
```

---

### 4. Add the input spreadsheet

Place the `colaboradores.xlsx` file inside the `entrada/` folder.

---

### 5. Run the pipeline

```bash
python main.py
```

The report will be generated at:

saida/resultado_final.xlsx

---
---

## 📦 Dependencies

pandas
requests
openpyxl
selenium
playwright
webdriver-manager

---

## 📝 Notes

* The pipeline performs a **single holiday API call per year**, caching results to improve performance.
* Cross-validation between RPA and API ensures higher reliability of ZIP code data.
* Any inconsistency or validation failure automatically downgrades priority to `BAIXA`.

---
