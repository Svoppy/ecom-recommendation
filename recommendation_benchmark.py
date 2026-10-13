#!/usr/bin/env python3
"""Reproduce the Kazakhstan e-commerce implicit-feedback regression benchmark.

The source data contains purchases but no impressions or explicit non-purchases.
This script therefore samples unobserved user-item pairs as *assumed* negatives.
All recommendation features are calculated from the earlier history window only.
"""

from __future__ import annotations

import argparse
import json
import math
import time
from pathlib import Path

import numpy as np
import pandas as pd
from catboost import CatBoostRegressor
from lightgbm import LGBMRegressor
from sklearn.compose import ColumnTransformer
from sklearn.ensemble import AdaBoostRegressor, ExtraTreesRegressor, GradientBoostingRegressor, HistGradientBoostingRegressor
from sklearn.impute import SimpleImputer
from sklearn.linear_model import ElasticNet, Lasso, Ridge
from sklearn.metrics import mean_squared_error, r2_score
from sklearn.model_selection import GroupKFold
from sklearn.neighbors import KNeighborsRegressor
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder, OrdinalEncoder, StandardScaler
from xgboost import XGBRegressor


CAT = ["category_code", "category_id", "brand"]
NUM = [
    "log_price", "log_user_events", "user_unique_products", "user_mean_log_price",
    "user_days_since_last", "log_item_events", "log_item_users", "user_category_share",
    "user_brand_share", "item_category_share", "item_brand_share",
]
FEATURES = CAT + NUM
SEED = 42


def load_and_split(path: Path, cutoff: str) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    use = ["event_time", "product_id", "category_id", "category_code", "brand", "price", "user_id"]
    df = pd.read_csv(path, usecols=use, dtype={c: "string" for c in use if c != "price"})
    df["event_time"] = pd.to_datetime(df["event_time"], utc=True, errors="coerce")
    df["price"] = pd.to_numeric(df["price"], errors="coerce")
    before = len(df)
    df = df.dropna(subset=["event_time", "product_id", "user_id", "price"])
    df = df[(df.event_time >= "2020-04-01") & (df.event_time < "2021-01-01") & (df.price > 0)].copy()
    split = pd.Timestamp(cutoff, tz="UTC")
    hist = df[df.event_time < split].copy()
    future = df[df.event_time >= split].copy()
    summary = {
        "rows_source": before,
        "rows_retained_user_item_time_price": len(df),
        "rows_history": len(hist),
        "rows_future": len(future),
        "cutoff_utc": split.isoformat(),
        "history_start": hist.event_time.min().isoformat(),
        "future_end": future.event_time.max().isoformat(),
        "users_history": int(hist.user_id.nunique()),
        "products_history": int(hist.product_id.nunique()),
    }
    return hist, future, summary


def choose_pairs(future: pd.DataFrame, history: pd.DataFrame, limit: int, rng: np.random.Generator) -> pd.DataFrame:
    # Evaluate users/items with usable historic profile features and consolidate
    # repeated purchases to one user-product relevance label.
    pairs = future.groupby(["user_id", "product_id"], as_index=False).agg(
        category_id=("category_id", "first"), category_code=("category_code", "first"), brand=("brand", "first")
    )
    active_users = set(history.user_id.unique())
    pairs = pairs[pairs.user_id.isin(active_users)].copy()
    if len(pairs) > limit:
        # A uniform deterministic sample over unique positive pairs.
        idx = rng.choice(pairs.index.to_numpy(), size=limit, replace=False)
        pairs = pairs.loc[idx].copy()
    return pairs.reset_index(drop=True)


def feature_maps(hist: pd.DataFrame, future: pd.DataFrame):
    end = hist.event_time.max()
    h = hist.copy()
    h["log_price"] = np.log1p(h.price.clip(lower=0))
    h["category_code"] = h.category_code.fillna("unknown")
    h["category_id"] = h.category_id.fillna("unknown")
    h["brand"] = h.brand.fillna("unknown")
    user = h.groupby("user_id").agg(
        log_user_events=("product_id", "size"), user_unique_products=("product_id", "nunique"),
        user_mean_log_price=("log_price", "mean"), last_event=("event_time", "max"),
    )
    user["user_days_since_last"] = (end - user.pop("last_event")).dt.total_seconds() / 86400
    item = h.groupby("product_id").agg(
        log_item_events=("product_id", "size"), log_item_users=("user_id", "nunique"),
        log_price=("price", "median"), category_id=("category_id", "first"),
        category_code=("category_code", "first"), brand=("brand", "first"),
    )
    item["log_price"] = np.log1p(item.log_price.clip(lower=0))
    item["category_code"] = item.category_code.fillna("unknown")
    item["category_id"] = item.category_id.fillna("unknown")
    item["brand"] = item.brand.fillna("unknown")
    user_cat = h.groupby(["user_id", "category_code"]).size()
    user_brand = h.groupby(["user_id", "brand"]).size()
    user_total = h.groupby("user_id").size()
    cat_total = h.groupby("category_code").size()
    brand_total = h.groupby("brand").size()
    item_cat = h.groupby(["product_id", "category_code"]).size()
    item_brand = h.groupby(["product_id", "brand"]).size()
    return user, item, user_cat, user_brand, user_total, cat_total, brand_total, item_cat, item_brand


def make_samples(pairs: pd.DataFrame, history: pd.DataFrame, future: pd.DataFrame, neg_per_pos: int, rng: np.random.Generator):
    (user, item, user_cat, user_brand, user_total, cat_total, brand_total,
     item_cat, item_brand) = feature_maps(history, future)
    # Negative candidates are drawn from historical products in the same category
    # when possible, excluding both prior user purchases and all observed future
    # positives. Absence of a purchase is a weak negative, not a true rejection.
    hist_pairs = history[["user_id", "product_id"]].drop_duplicates()
    seen_hist = hist_pairs.groupby("user_id").product_id.apply(set).to_dict()
    seen_future = future.groupby("user_id").product_id.apply(set).to_dict()
    catalog = item.reset_index()
    eligible = catalog[catalog.product_id.notna()]
    by_cat = {k: g.product_id.to_numpy() for k, g in eligible.groupby("category_code")}
    all_products = eligible.product_id.to_numpy()
    rows: list[dict] = []
    group_ids: list[str] = []

    def add(u, pid, label):
        if u not in user.index or pid not in item.index:
            return
        ir = item.loc[pid]
        ur = user.loc[u]
        cat, brand = str(ir.category_code), str(ir.brand)
        total = max(int(user_total.get(u, 0)), 1)
        c_count = int(user_cat.get((u, cat), 0))
        b_count = int(user_brand.get((u, brand), 0))
        i_count = int(item_cat.get((pid, cat), 0))
        ib_count = int(item_brand.get((pid, brand), 0))
        rows.append({
            "category_code": cat, "category_id": str(ir.category_id), "brand": brand,
            "log_price": float(ir.log_price), "log_user_events": math.log1p(float(ur.log_user_events)),
            "user_unique_products": float(ur.user_unique_products),
            "user_mean_log_price": float(ur.user_mean_log_price),
            "user_days_since_last": float(ur.user_days_since_last),
            "log_item_events": math.log1p(float(ir.log_item_events)),
            "log_item_users": math.log1p(float(ir.log_item_users)),
            "user_category_share": c_count / total, "user_brand_share": b_count / total,
            "item_category_share": i_count / max(int(cat_total.get(cat, 0)), 1),
            "item_brand_share": ib_count / max(int(brand_total.get(brand, 0)), 1),
            "label": int(label), "user_id": str(u), "product_id": str(pid),
        })
        group_ids.append(str(u))

    for row in pairs.itertuples(index=False):
        u, pid = str(row.user_id), str(row.product_id)
        add(u, pid, 1)
        category = str(item.loc[pid].category_code) if pid in item.index else str(row.category_code)
        pool = by_cat.get(category, all_products)
        forbidden = seen_hist.get(u, set()) | seen_future.get(u, set()) | {pid}
        choices = np.asarray([x for x in pool if x not in forbidden], dtype=object)
        if len(choices) < neg_per_pos:
            choices = np.asarray([x for x in all_products if x not in forbidden], dtype=object)
        if len(choices):
            picks = rng.choice(choices, size=min(neg_per_pos, len(choices)), replace=len(choices) < neg_per_pos)
            for neg in picks:
                add(u, str(neg), 0)
    data = pd.DataFrame(rows)
    groups = np.asarray(group_ids)
    return data, groups, len(pairs), len(user), len(item)


def pipeline(model, scale=False):
    pre = ColumnTransformer([
        ("cat", OneHotEncoder(handle_unknown="ignore", min_frequency=20), CAT),
        ("num", Pipeline([("impute", SimpleImputer(strategy="median")),
                           ("scale", StandardScaler() if scale else "passthrough")]), NUM),
    ], remainder="drop", sparse_threshold=1.0)
    return Pipeline([("prep", pre), ("model", model)])


def hist_pipeline(model):
    # HistGradientBoosting currently requires a dense array. Ordinal encoding is
    # used only for this tree baseline to bound memory on high-cardinality text.
    pre = ColumnTransformer([
        ("cat", Pipeline([("impute", SimpleImputer(strategy="most_frequent")),
                           ("ordinal", OrdinalEncoder(handle_unknown="use_encoded_value", unknown_value=-1))]), CAT),
        ("num", SimpleImputer(strategy="median"), NUM),
    ], sparse_threshold=0.0)
    return Pipeline([("prep", pre), ("model", model)])


def ndcg_at_10(frame: pd.DataFrame, pred: np.ndarray) -> tuple[float, float]:
    tmp = frame[["user_id", "label"]].copy()
    tmp["score"] = pred
    ndcgs, hits = [], []
    for _, g in tmp.groupby("user_id", sort=False):
        y = g.label.to_numpy(dtype=float)
        s = g.score.to_numpy(dtype=float)
        order = np.argsort(-s)[:10]
        ideal = np.argsort(-y)[:10]
        discounts = 1 / np.log2(np.arange(2, len(order) + 2))
        dcg = float((y[order] * discounts).sum())
        idcg = float((y[ideal] * (1 / np.log2(np.arange(2, len(ideal) + 2)))).sum())
        ndcgs.append(dcg / idcg if idcg else 0.0)
        hits.append(float(y[order].sum() > 0))
    return float(np.mean(ndcgs)), float(np.mean(hits))


def models(seed: int, quick: bool = False):
    # Fixed classroom comparison, not a hyperparameter-optimization contest.
    n_trees = 80 if quick else 160
    common = {"random_state": seed}
    return [
        ("Ridge", pipeline(Ridge(alpha=1.0), scale=True)),
        ("Lasso", pipeline(Lasso(alpha=0.0005, max_iter=3000), scale=True)),
        ("Elastic Net", pipeline(ElasticNet(alpha=0.0005, l1_ratio=0.5, max_iter=3000), scale=True)),
        ("KNN Regression", pipeline(KNeighborsRegressor(n_neighbors=9, weights="distance", n_jobs=1), scale=True)),
        ("Extra Trees Regression", pipeline(ExtraTreesRegressor(n_estimators=n_trees, min_samples_leaf=2, max_features=0.9, n_jobs=4, **common))),
        ("Adaptive Boosting", pipeline(AdaBoostRegressor(n_estimators=n_trees, learning_rate=0.04, **common))),
        ("Gradient Boosting Regression", pipeline(GradientBoostingRegressor(n_estimators=n_trees, max_depth=2, learning_rate=0.04, loss="huber", **common))),
        ("XGBoost", pipeline(XGBRegressor(n_estimators=n_trees, max_depth=5, learning_rate=0.05, subsample=0.85, colsample_bytree=0.9, reg_lambda=2, n_jobs=4, objective="reg:squarederror", tree_method="hist", verbosity=0, **common))),
        ("LightGBM", pipeline(LGBMRegressor(n_estimators=n_trees, max_depth=-1, num_leaves=31, learning_rate=0.05, subsample=0.85, colsample_bytree=0.9, reg_lambda=2, n_jobs=4, verbosity=-1, **common))),
        ("CatBoost", (CatBoostRegressor(iterations=n_trees, depth=6, learning_rate=0.05, loss_function="RMSE", verbose=False, random_seed=seed, thread_count=4), "native")),
        ("HistGradientBoosting", hist_pipeline(HistGradientBoostingRegressor(max_iter=n_trees, max_leaf_nodes=31, learning_rate=0.06, l2_regularization=1.0, **common))),
    ]


def run(args):
    rng = np.random.default_rng(args.seed)
    hist, future, data_summary = load_and_split(Path(args.data), args.cutoff)
    pairs = choose_pairs(future, hist, args.max_positive_pairs, rng)
    frame, groups, pair_count, user_count, item_count = make_samples(pairs, hist, future, args.negatives, rng)
    if frame.empty or frame.label.nunique() < 2:
        raise RuntimeError("The chosen time window did not yield both positive and negative examples.")
    X, y = frame[FEATURES], frame.label.to_numpy(dtype=float)
    cv = GroupKFold(n_splits=10)
    output = []
    for name, model_spec in models(args.seed, args.quick):
        fold_rmse, fold_r2, fold_ndcg, fold_hit = [], [], [], []
        started = time.time()
        for fold, (tr, te) in enumerate(cv.split(X, y, groups), 1):
            model, native = model_spec if isinstance(model_spec, tuple) else (model_spec, "pipe")
            if native == "native":
                model.fit(X.iloc[tr], y[tr], cat_features=CAT)
                pred = model.predict(X.iloc[te])
            else:
                model.fit(X.iloc[tr], y[tr])
                pred = model.predict(X.iloc[te])
            fold_rmse.append(float(mean_squared_error(y[te], pred) ** 0.5))
            fold_r2.append(float(r2_score(y[te], pred)))
            n, h = ndcg_at_10(frame.iloc[te], pred)
            fold_ndcg.append(n)
            fold_hit.append(h)
            print(f"{name}: fold {fold}/10 RMSE={fold_rmse[-1]:.5f} R2={fold_r2[-1]:.5f}", flush=True)
        output.append({
            "algorithm": name, "features": len(FEATURES), "targets": 1, "folds": 10,
            "rmse_mean": float(np.mean(fold_rmse)), "rmse_std": float(np.std(fold_rmse, ddof=1)),
            "r2_mean": float(np.mean(fold_r2)), "r2_std": float(np.std(fold_r2, ddof=1)),
            "ndcg_at_10_mean": float(np.mean(fold_ndcg)), "hit_rate_at_10_mean": float(np.mean(fold_hit)),
            "elapsed_seconds": round(time.time() - started, 2),
        })
        print(f"COMPLETED {name} in {output[-1]['elapsed_seconds']} sec", flush=True)
    output.sort(key=lambda x: x["rmse_mean"])
    result = {"dataset": data_summary, "evaluation": {
        "target_definition": "1=observed later purchase; 0=sampled unseen pair (assumed negative)",
        "split": "earlier history features, later-period interaction pairs, 10-fold GroupKFold by user",
        "positive_user_item_pairs": pair_count, "sampled_candidates": len(frame),
        "users_in_sample": int(frame.user_id.nunique()), "historical_products": item_count,
        "negative_samples_per_positive": args.negatives, "feature_names": FEATURES,
        "folds": 10, "seed": args.seed,
    }, "results": output}
    Path(args.output).parent.mkdir(parents=True, exist_ok=True)
    Path(args.output).write_text(json.dumps(result, indent=2), encoding="utf-8")
    pd.DataFrame(output).to_csv(Path(args.output).with_suffix(".csv"), index=False)
    print(json.dumps(result, indent=2))


def main():
    p = argparse.ArgumentParser()
    default_data = Path("dataset/kz.csv")
    if not default_data.exists() and Path("dataset/kz.csv.gz").exists():
        default_data = Path("dataset/kz.csv.gz")
    p.add_argument("--data", default=str(default_data))
    p.add_argument("--cutoff", default="2020-09-01")
    p.add_argument("--max-positive-pairs", type=int, default=6000)
    p.add_argument("--negatives", type=int, default=4)
    p.add_argument("--seed", type=int, default=SEED)
    p.add_argument("--quick", action="store_true", help="Use fewer boosting rounds for a smoke run")
    p.add_argument("--output", default="results/model_results.json")
    run(p.parse_args())


if __name__ == "__main__":
    main()
