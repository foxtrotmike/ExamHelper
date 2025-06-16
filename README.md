# 📝 Marks Integration and Summary Tools

This repository provides utilities for managing and summarising student marks in module. It includes:

1. **`copyit.py`** – A Python script to update assignment marks in the main gradebook based on adjustments from the feedback sheet.
2. **`summary_generator.vba`** – A VBA macro to generate per-route summaries and overall grade distributions.

---

## 📂 Repository Contents

| File                | Description                                                                 |
|---------------------|-----------------------------------------------------------------------------|
| `copyit.py`         | Python script to copy final assignment marks into the main gradebook. |
| `summary_generator.vba` | VBA macro to compute coursework/exam/total averages and grade distributions. |

---

## 🚀 Quick Start

1. **Run `copyit.py`**  
   Updates the gradebook (`CS429-15.xlsx`) using final marks from the feedback sheet (`CS429-Assignment-2.xlsx`).

2. **Run `summary_generator.vba`**  
   In Excel, generates a `Summary` sheet with route-level statistics and overall performance breakdown.

> 📄 **For full configuration and usage instructions**, please see the **docstring in `copyit.py`** and the **comments in `summary_generator.vba`**.

---

## 📌 Requirements

- Python 3 with `pandas` installed
- Microsoft Excel (for VBA)

---

## ✅ Example Outputs

- A new gradebook file: `CS429-15_updated.xlsx`
- A summary sheet showing:
  - Per-group average coursework, exam, and total marks
  - Count of valid entries per metric
  - Overall grade distribution (1st, 2.1, 2.2, 3rd, Fail)

---

## 👨‍💼 Author

**Fayyaz Minhas**  
University of Warwick

---

## 📜 License

MIT License
