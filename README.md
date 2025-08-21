README.md  
```
# Flowing Red Flag Evaluation System

A lightweight, **offline-first** desktop application built with Python/Tkinter to automate weekly “Red Flag” evaluations for school classes.

- **10 evaluation criteria**: tardiness, morning reading, energy saving, dress code, exercises, nap, hygiene, patrol, paperwork, dormitory.  
- **AM/PM split** for items like exercises and hygiene.  
- **Weighted scoring & bonus/penalty** system.  
- **Live ranking** and one-click export to **Excel/CSV**.  
- **JSON save/load** with unlimited undo/redo.  

## Quick Start
```bash
git clone https://github.com/<your-username>/flowing-red-flag-evaluation.git
cd flowing-red-flag-evaluation
python main.py
```

## Requirements
- Python ≥ 3.7  
- Tkinter (bundled with standard CPython)  
- Optional: `pandas`, `openpyxl` for Excel export (`pip install pandas openpyxl`)

## Usage
1. Launch the app.  
2. Double-click any cell to edit scores.  
3. Press **Calculate Totals** to see the ranking.  
4. **Export** results via the “Tools” menu.

## License
Licensed under [CC BY-SA 3.0](LICENSE).

## Contributing
Pull requests welcome! For bugs or features, open an issue.
```
