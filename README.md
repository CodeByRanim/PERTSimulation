# PERTSimulation
# 📊 PERT Simulation in Excel + VBA

This repository contains a real-world example of project planning using the PERT (Program Evaluation and Review Technique) method. 

## 📁 Contents

- `PERT_real_project.xlsx` — Excel file with realistic task data for a system deployment project.
- `CalculPERT.bas` — VBA macro module to automatically calculate:
  - Expected Duration (TE)
  - Variance

## 🚀 How to Use

1. Open `PERT_real_project.xlsx`
2. Press `Alt + F11` to open the VBA editor.
3. Import the module `CalculPERT.bas` (`File > Import File...`)
4. Go back to Excel and run the macro `CalculPERT` to compute TE and Variance.
5. Optionally, insert a button in the sheet linked to the macro for ease of use.

## 📈 Example Task

| Task | Optimistic | Most Likely | Pessimistic | Expected (TE) | Variance |
|------|------------|-------------|-------------|---------------|----------|
| D    | 7          | 10          | 15          | 10.33         | 1.78     |

## 🔗 Use Case

Project: **Deployment of a Logistics Information System**

Tasks: Needs analysis → Software choice → Configuration → Development → Testing → Integration → Training → Go-live

## 🛠️ Author

Created by BEDOUI Ranim, Supply Chain Engineer & Excel/VBA Enthusiast
