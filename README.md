# AllocationModel

A desktop application that optimally allocates financial deals to purchasers using Integer Linear Programming (ILP), with ML-assisted deal classification and interactive What-If budget analysis.

Built for non-technical users — no coding required. Load your data, click **Run Optimisation**, and export results to CSV or Excel.

---

## Features

- **Optimal allocation** — maximises total deal value under purchaser budget constraints
- **Auto-classify deals with ML** — K-Means clustering suggests Prepay / PPA labels from deal values alone
- **What-If analysis** — instantly estimates how changing a purchaser's budget affects total allocated value
- **CSV and Excel support** — load `.csv`, `.xlsx`, or `.xlsm` files; export results to either format
- **Cross-platform** — runs on Windows and macOS

---

## Requirements

- Python **3.9** or higher  
- Internet connection for the one-time `pip install` step

---

## Installation

**1. Clone the repository**

```bash
git clone https://github.com/your-username/AllocationModel.git
cd AllocationModel
```

**2. (Recommended) Create a virtual environment**

```bash
# macOS / Linux
python3 -m venv .venv
source .venv/bin/activate

# Windows
python -m venv .venv
.venv\Scripts\activate
```

**3. Install dependencies**

```bash
pip install -r requirements.txt
```

**4. Launch the app**

```bash
python main.py
```

A window will open. No further setup needed.

---

## Input File Format

The app accepts two formats:

### Option A — Two CSV files

**deals.csv** — one row per deal

| Column | Type | Required | Description |
|---|---|---|---|
| `deal_id` | text | Yes | Unique deal identifier |
| `deal_value` | number | Yes | Monetary value of the deal |
| `deal_type` | text | No* | `Prepay` or `PPA` |

*If omitted or unknown, use the **Auto-classify** button to let the ML model suggest types.

**purchasers.csv** — one row per purchaser

| Column | Type | Required | Description |
|---|---|---|---|
| `purchaser_id` | text | Yes | Unique purchaser identifier |
| `purchaser_max` | number | Yes | Maximum allocation budget |
| `purchaser_preference` | text | Yes | `Prepay` or `PPA` — preferred deal type |

Column names are **case-insensitive** and leading/trailing spaces are ignored.

### Option B — Single Excel file (`.xlsx` or `.xlsm`)

Same columns as above, split across two sheets named **`Deals`** and **`Purchasers`**.

### Sample data

Ready-to-use example files are included in `sample_data/`:

```
sample_data/example_deals.csv
sample_data/example_purchasers.csv
```

---

## How to Use

### Screen 1 — Load Data

1. Select **CSV (2 files)** or **Excel (1 file)** using the toggle.
2. Click **Browse** to select your file(s).
3. Click **Load & Preview** — a table preview of your deals and purchasers will appear.
4. *(Optional)* Click **Auto-classify Deals with ML** to have the app suggest Prepay / PPA labels based on deal value distribution. Deals with suggested labels different from their original labels are flagged as overrides for your review.
5. Click **Next: Configure →**.

### Screen 2 — Configure & Run

| Setting | Description | Default |
|---|---|---|
| Time limit | Maximum seconds the solver runs before returning the best solution found | 60 s |
| Each purchaser gets at least 1 deal | Forces the solver to assign at least one deal per purchaser (when budget permits) | On |
| Penalise preference mismatches | Soft penalty when a purchaser receives a deal type they did not prefer | On |

Click **Run Optimisation**. The solver log appears in real time. Click **View Results →** when it finishes.

### Screen 3 — Results & What-If

**Summary cards** show solver status, total allocated value, deals allocated (%), and budget used (%).

**Allocation table** lists every deal with its assigned purchaser. Unallocated deals are highlighted in orange.

**What-If Analysis:**

1. Select a purchaser from the dropdown.
2. Move the slider to set the budget change percentage (−30% to +30%).
3. Click **Analyse** — builds a sensitivity model (~30 seconds) and shows an instant estimate.
4. *(Optional)* Click **Run Exact** for a precise solver result at that exact budget level.

**Export:** use **Export CSV** or **Export Excel** to save results. The Excel file contains two sheets: *Allocation Results* and *Summary* (per-purchaser budget usage).

---

## ML Features

### Deal Auto-Classification (K-Means)

When deal types are unknown or need verification, the app clusters deals by value using K-Means:

- The cluster with **higher** average value is labelled **PPA**
- The cluster with **lower** average value is labelled **Prepay**
- Each suggestion includes a **confidence score** (0–1). Scores near 0.5 indicate the deal sits close to the cluster boundary — review those manually.
- A **silhouette score** check prevents the model from creating artificial splits when data is unimodal (i.e. all deals have similar values).

The classifier never modifies your data silently. You review suggestions on Screen 1 before they are applied.

### What-If Sensitivity Analysis (Linear Regression)

The What-If panel answers: *"What happens to total allocated value if I change this purchaser's budget?"*

Internally:
1. The solver runs several times with the purchaser's capacity perturbed by ±5%, ±10%, ±20%.
2. A linear regression is fitted on (capacity change % → value change).
3. The fitted line gives instant estimates for any budget level without re-running the full solver.

The **R²** of the fit is reported. If R² < 0.7 the explanation note warns that the relationship is non-linear and suggests using **Run Exact** for accuracy.

---

## Project Structure

```
AllocationModel/
├── app/
│   ├── data_loader.py      # CSV + Excel input, AllocationInput dataclass
│   ├── optimizer.py        # ILP model (PuLP + CBC solver)
│   ├── ml_classifier.py    # K-Means deal classification
│   ├── sensitivity.py      # What-If regression analysis
│   ├── exporter.py         # CSV + Excel export
│   └── gui.py              # CustomTkinter 3-screen interface
├── tests/
│   ├── test_optimizer.py
│   └── test_ml_classifier.py
├── sample_data/
│   ├── example_deals.csv
│   └── example_purchasers.csv
├── main.py                 # Entry point
├── requirements.txt
└── PLAN.md
```

---

## Running Tests

```bash
python -m pytest tests/ -v
```

All 32 tests should pass in under 15 seconds.

---

## Troubleshooting

**`ModuleNotFoundError: No module named 'customtkinter'`** (or any other module)  
Run `pip install -r requirements.txt` again. If you are using a virtual environment, make sure it is activated first.

**`python: command not found` on macOS**  
Use `python3` instead of `python`. On modern macOS, `python` may not be aliased to Python 3.

**The window opens but is very small or blurry on a high-DPI screen**  
This is a known DPI scaling behaviour on some Windows configurations. Resize the window or set your display scaling to 100% in Windows Display Settings.

**`Solver failed` error when running the optimisation**  
Check that:  
- All `deal_value` and `purchaser_max` columns contain positive numbers (no text, no `$` signs).  
- At least one purchaser has a budget larger than the smallest deal value.  
- The `deal_type` and `purchaser_preference` columns contain only `Prepay` or `PPA` (case-insensitive).

**`Sheet 'Purchasers' not found`**  
When using an Excel file, the two sheets must be named exactly **`Deals`** and **`Purchasers`** (capital P, capital D). Check the tab names in your workbook.

**The app is slow on the first run**  
PuLP writes a temporary model file and calls the CBC solver as a subprocess. This is normal. Subsequent runs in the same session are faster.

---

## Dependencies

| Package | Purpose |
|---|---|
| `pulp` | ILP model formulation + CBC solver (bundled) |
| `pandas` | CSV and Excel data loading |
| `openpyxl` | Excel read/write engine |
| `customtkinter` | Modern desktop GUI framework |
| `scikit-learn` | K-Means clustering + Linear Regression |

---

## License

MIT
