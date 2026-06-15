# 🎲 Board Game Success Analysis
**University of Michigan — MADS Milestone 1 (Summer 2026)**  

This repository contains our group project for the University of Michigan's Master of Applied Data Science (MADS) program. The project focuses on analyzing board game success using multiple datasets, including BoardGameGeek (BGG), Google Trends, Wikipedia, and sales data.

> ⚠️ Note: Due to file size limitations, the raw datasets are not included in this repository. However, all data collection, caching, and processing pipelines are preserved so results can be reproduced.

---

## 👥 Team Acknowledgment

This was a fully collaborative team effort. Special thanks to:

- **Chi-Hsun Fan** — API development (BGG, BGG resale lisings, Wikipedia), documentation, and analytical contributions  
- **Sam Meisel** — Regression modeling, statistical analysis, and data processing  

Their work was instrumental, particularly in handling complex API integrations and advanced modeling.

- [Chi-Hsun Fan — LinkedIn](https://www.linkedin.com/in/joye-fan-038b27216/)
- [Sam Meisel — LinkedIn](https://www.linkedin.com/in/meiselsa/)

---

## 🧪 Data Collection & Infrastructure

The following notebooks were used to pull, cache, and preprocess raw data:

- `Google Trends.ipynb`
- `Wikipedia Data.ipynb`
- `BGG API Data Pull.ipynb`
- `fetching_raw_data.ipynb`

These scripts demonstrate how data was collected from external APIs and prepared for analysis.  
Some API pulls (especially Google Trends and BGG “things” endpoints) required careful handling due to rate limits and response complexity.

> 💡 Future improvement: We recommend using serialized formats (e.g., `.pkl` files) instead of `.csv` for more efficient storage and faster I/O.

---

## 📊 Analysis Workflow

After completing data collection and aggregation, run the following notebooks in order:

1. `Init.ipynb`  
2. `Joins and Cleanup.ipynb`  
3. `Binarize.ipynb`  
4. `Award vs Marketplace.ipynb`  
5. `Trends, Awards, and Pricing EDA.ipynb`  
6. `BGG EDA.ipynb`  
7. `Sales EDA.ipynb`  
8. `Success Metric.ipynb`  
9. `Regression Analysis.ipynb`  

These notebooks contain the core analysis, including exploratory data analysis (EDA), feature engineering, and statistical modeling.

---

## ⚙️ Requirements

All required Python packages are listed in `requirements.txt`.

We recommend using a virtual environment
