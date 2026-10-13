# E-commerce recommendation benchmark

This project compares eleven regression-based scoring models for implicit-feedback recommendations on Kazakhstan's `kz.csv` electronics-store extract.

## Dataset

- Kaggle source: https://www.kaggle.com/datasets/mkechinov/ecommerce-purchase-history-from-electronics-store
- File: `dataset/kz.csv.gz` in the repository; the original uncompressed `dataset/kz.csv` is also kept locally. The gzip copy is lossless and pandas reads it directly.
- The raw file has 2,633,521 rows and eight columns. It includes event time, order, product, category, category code, brand, price, and user ID.
- The compressed dataset is 65 MiB, below GitHub's 100 MB per-file limit. The source record does not document the price currency, so the study treats price as a relative recorded value.
- The dataset includes a large number of missing user IDs and a small set of 1970 timestamps. The benchmark excludes records without usable user ID, product, positive price, or an in-range timestamp. The paper reports this limitation.

## Reproduce the model comparison

Python 3.12 is recommended. Install dependencies in a fresh environment with `python -m pip install -r requirements.txt`, then run:

```bash
python recommendation_benchmark.py --data dataset/kz.csv.gz --max-positive-pairs 6000 --negatives 4 --seed 42 --output results/model_results.json
```

The benchmark uses events before 2020-09-01 to build historical user/item features. Later purchases make positive pairs. For each positive pair, four products the user did not buy in the future period are sampled as assumed negatives. Ten-fold `GroupKFold` holds out whole users. Results contain RMSE, R², NDCG@10, and HitRate@10. RMSE and R² answer the assignment's regression comparison; the ranking metrics provide recommendation-specific context.

The paper and slides use the run recorded in `results/model_results.json` and `.csv`. Re-running the script recreates both result files. It does not perform hyperparameter tuning.

## Deliverables

- `dataset/kz.csv.gz` — the real Kaggle source dataset, compressed losslessly for GitHub. The original `dataset/kz.csv` remains in the local working copy.
- `recommendation_benchmark.py` — source code for data preparation, features, folds, and all eleven estimators.
- `recommendation_benchmark.ipynb` — notebook walkthrough using the same real CSV and saved benchmark outputs; it can optionally rerun all 11 models.
- `results/model_results.csv` and `results/model_results.json` — fold summary and machine-readable results.
- `research_paper.docx` — editable research paper with the results table and analysis.
- `deliverables/recommendation_presentation_final.pptx` — editable 14-slide presentation.
- `article_body.md` and `build_docx.py` — paper text and its document builder.

The paper's 3,500-word count applies to Introduction through Results and Discussion, excluding headings, abstract, keywords, conclusion, and references.

## Academic integrity

The assignment brief requires an originality score above 97% and says AI was not used. This draft was prepared with AI assistance, and no similarity checker was run. Do not claim otherwise. Review and rewrite the report in your own words, verify the instructor's dataset uniqueness rule, and follow the course's AI policy before submission.
# ecom-recommendation
