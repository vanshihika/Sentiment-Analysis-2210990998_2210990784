import random
import warnings
import numpy as np
import pandas as pd
from scipy import stats
from scipy.stats import spearmanr, pearsonr, kruskal, mannwhitneyu
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.linear_model import LogisticRegression
from sklearn.model_selection import train_test_split
from sklearn.metrics import (accuracy_score, classification_report,
                             confusion_matrix, f1_score, precision_score,
                             recall_score)
from sklearn.utils import resample
from statsmodels.stats.contingency_tables import mcnemar
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
from textblob import TextBlob
import re
import nltk
from nltk.corpus import stopwords
 
warnings.filterwarnings("ignore")
 
DATASET_PATH = "Reviews.csv"  
SAMPLE_N     = 10_000
SEED         = 42
SILVER_N     = 500           
TFIDF_MAX    = 5_000
LR_C         = 1.0
BOOTSTRAP    = 1_000
TRAIN_FRAC   = 0.80          
 
random.seed(SEED)
np.random.seed(SEED)
 

nltk.download("stopwords", quiet=True)
nltk.download("punkt",     quiet=True)
STOP_WORDS = set(stopwords.words("english"))
 
 

#LOAD & SAMPLE

def load_data(path: str, n: int, seed: int) -> pd.DataFrame:
    print(f"\n[1] Loading dataset from: {path}")
    df = pd.read_csv(path, usecols=["Text", "Score"])
    df = df.rename(columns={"Text": "review_text", "Score": "star_rating"})
    df = df.dropna(subset=["review_text", "star_rating"])
    df = df[df["review_text"].str.strip() != ""]
    df["star_rating"] = df["star_rating"].astype(int)
    df = df[df["star_rating"].between(1, 5)]
 
    # Remove duplicates
    before = len(df)
    df = df.drop_duplicates(subset=["review_text"])
    print(f"   Removed {before - len(df)} duplicate rows.")
 
    # Fixed-seed random sample
    df = df.sample(n=n, random_state=seed).reset_index(drop=True)
    print(f"   Sampled {len(df):,} reviews (seed={seed}).")
    return df
 
 
# =============================================================================
# 2. PREPROCESSING
# =============================================================================
def preprocess_for_tfidf(text: str) -> str:
    """For TF-IDF features (Logistic Regression) only — NOT used for VADER/TextBlob."""
    text = text.lower()
    text = re.sub(r"<[^>]+>", " ", text)          # remove HTML tags
    text = re.sub(r"http\S+|www\S+", " ", text)    # remove URLs
    text = re.sub(r"[^a-z\s]", " ", text)          # remove special chars / punctuation
    tokens = text.split()
    tokens = [t for t in tokens if t not in STOP_WORDS]
    return " ".join(tokens)
 
 
def preprocess(df: pd.DataFrame) -> pd.DataFrame:
    print("\n[2] Preprocessing …")
    # Keep original text intact for VADER and TextBlob
    df["clean_text"] = df["review_text"].apply(preprocess_for_tfidf)
    print("    Done.")
    return df
 
 
# =============================================================================
# 3. SENTIMENT SCORING
# =============================================================================
def score_sentiment(df: pd.DataFrame) -> pd.DataFrame:
    print("\n[3] Running VADER and TextBlob on all reviews …")
    vader = SentimentIntensityAnalyzer()
 
    # VADER — uses original (un-stripped) text
    df["vader_compound"] = df["review_text"].apply(
        lambda t: vader.polarity_scores(t)["compound"]
    )
 
    # VADER classification using standard thresholds
    def vader_label(s):
        if s >= 0.05:
            return "positive"
        elif s <= -0.05:
            return "negative"
        else:
            return "neutral"
 
    df["vader_label"] = df["vader_compound"].apply(vader_label)
 
    # TextBlob — uses original text
    df["textblob_polarity"] = df["review_text"].apply(
        lambda t: TextBlob(t).sentiment.polarity
    )
 
    def textblob_label(p):
        if p > 0.05:
            return "positive"
        elif p < -0.05:
            return "negative"
        else:
            return "abstain"          # abstention as per paper
 
    df["textblob_label"] = df["textblob_polarity"].apply(textblob_label)
    print("    Done.")
    return df
 
 
# =============================================================================
# 4. TABLE I — Distribution
# =============================================================================
def table_i(df: pd.DataFrame):
    print("\n" + "="*70)
    print("TABLE I — Sentiment Classification and Rating Distribution (N=10,000)")
    print("="*70)
 
    vader_counts = df["vader_label"].value_counts()
    rating_counts = df["star_rating"].value_counts().sort_index(ascending=False)
 
    rows = []
    for label in ["positive", "negative", "neutral"]:
        n = vader_counts.get(label, 0)
        rows.append({"Category": f"VADER {label.capitalize()}", "n": n,
                     "%": f"{100*n/len(df):.1f}%"})
    for star in [5, 4, 3, 2, 1]:
        n = rating_counts.get(star, 0)
        rows.append({"Category": f"{star}★", "n": n, "%": f"{100*n/len(df):.1f}%"})
 
    t1 = pd.DataFrame(rows)
    print(t1.to_string(index=False))
    return t1
 
 
# =============================================================================
# 5. TABLE II — Correlation Analysis (H1)
# =============================================================================
def bootstrap_corr(x, y, n_boot=1000, seed=42, method="spearman"):
    rng = np.random.default_rng(seed)
    boot_corrs = []
    for _ in range(n_boot):
        idx = rng.integers(0, len(x), len(x))
        xi, yi = x[idx], y[idx]
        if method == "spearman":
            r, _ = spearmanr(xi, yi)
        else:
            r, _ = pearsonr(xi, yi)
        boot_corrs.append(r)
    return np.percentile(boot_corrs, [2.5, 97.5])
 
 
def table_ii(df: pd.DataFrame):
    print("\n" + "="*70)
    print("TABLE II — Correlation Results (H1, N=10,000)")
    print("="*70)
 
    x = df["vader_compound"].values
    y = df["star_rating"].values
 
    rho, p_rho = spearmanr(x, y)
    ci_rho = bootstrap_corr(x, y, BOOTSTRAP, SEED, "spearman")
 
    r, p_r = pearsonr(x, y)
    ci_r = bootstrap_corr(x, y, BOOTSTRAP, SEED, "pearson")
    r2 = r**2
 
    print(f"  Spearman ρ   : {rho:.3f}   95% CI [{ci_rho[0]:.3f}, {ci_rho[1]:.3f}]   p < 0.001")
    print(f"  Pearson  r   : {r:.3f}   95% CI [{ci_r[0]:.3f}, {ci_r[1]:.3f}]   p < 0.001")
    print(f"  Explained r² : {100*r2:.1f}%")
 
    effect = "small" if abs(rho) < 0.3 else ("medium" if abs(rho) < 0.5 else "large")
    print(f"  Effect size  : {effect} (Cohen 1988 conventions)")
 
    return rho, r, r2
 
 
# =============================================================================
# 6. TABLE III — Variance across rating groups (H2)
# =============================================================================
def cohens_d(a, b):
    pooled_std = np.sqrt((np.var(a, ddof=1) + np.var(b, ddof=1)) / 2)
    return (np.mean(a) - np.mean(b)) / pooled_std if pooled_std else 0.0
 
 
def table_iii(df: pd.DataFrame):
    print("\n" + "="*70)
    print("TABLE III — VADER Sentiment Score Statistics by Star Rating (H2)")
    print("="*70)
 
    five_star_scores = df[df["star_rating"] == 5]["vader_compound"].values
 
    rows = []
    groups = {}
    for star in [1, 2, 3, 4, 5]:
        scores = df[df["star_rating"] == star]["vader_compound"].values
        groups[star] = scores
        n = len(scores)
        mu = np.mean(scores)
        sigma = np.std(scores, ddof=1)
        se = sigma / np.sqrt(n)
        ci_lo = mu - 1.96 * se
        ci_hi = mu + 1.96 * se
        d = cohens_d(scores, five_star_scores) if star != 5 else 0.0
        d_label = ("large" if abs(d) >= 0.8 else
                   "medium" if abs(d) >= 0.5 else
                   "small"  if abs(d) >= 0.2 else "negligible")
        rows.append({
            "Rating": f"{star}★",
            "n": n,
            "Mean (μ)": f"{mu:.3f}",
            "SD (σ)": f"{sigma:.3f}",
            "95% CI": f"[{ci_lo:.3f}, {ci_hi:.3f}]",
            "Cohen's d vs 5★": f"{abs(d):.2f} ({d_label})" if star != 5 else "—"
        })
 
    t3 = pd.DataFrame(rows)
    print(t3.to_string(index=False))
 
    # Kruskal-Wallis
    stat, p = kruskal(*[groups[s] for s in [1, 2, 3, 4, 5]])
    print(f"\n  Kruskal-Wallis H = {stat:.1f}, df = 4, p < 0.001  (actual p = {p:.2e})")
 
    # Pairwise Mann-Whitney with Bonferroni
    pairs = [(1,2),(1,3),(1,4),(1,5),(2,3),(2,4),(2,5),(3,4),(3,5),(4,5)]
    n_comparisons = len(pairs)
    print("\n  Pairwise Mann-Whitney U (Bonferroni-corrected):")
    for (a, b) in pairs:
        u, p_raw = mannwhitneyu(groups[a], groups[b], alternative="two-sided")
        p_adj = min(p_raw * n_comparisons, 1.0)
        sig = "***" if p_adj < 0.001 else ("**" if p_adj < 0.01 else "*" if p_adj < 0.05 else "ns")
        print(f"    {a}★ vs {b}★: U={u:.0f}, p_adj={p_adj:.4f} {sig}")
 
    return groups
 
 
# =============================================================================
# 7. SILVER-STANDARD LABELING
# =============================================================================
def build_silver_set(df: pd.DataFrame, n_each: int = 250, seed: int = SEED) -> pd.DataFrame:
    print(f"\n[7] Building silver-standard annotation set ({n_each*2} reviews) …")
 
    pos_pool = df[
        (df["star_rating"].isin([4, 5])) &
        (df["vader_compound"] >= 0.05)
    ].copy()
 
    neg_pool = df[
        (df["star_rating"].isin([1, 2])) &
        (df["vader_compound"] <= -0.05)
    ].copy()
 
    print(f"    Positive pool: {len(pos_pool):,}   |   Negative pool: {len(neg_pool):,}")
 
    if len(pos_pool) < n_each or len(neg_pool) < n_each:
        raise ValueError(
            f"Not enough reviews for balanced set. "
            f"Positive: {len(pos_pool)}, Negative: {len(neg_pool)}, need {n_each} each."
        )
 
    pos_sample = pos_pool.sample(n=n_each, random_state=seed)
    neg_sample = neg_pool.sample(n=n_each, random_state=seed)
 
    pos_sample = pos_sample.copy(); pos_sample["silver_label"] = "positive"
    neg_sample = neg_sample.copy(); neg_sample["silver_label"] = "negative"
 
    silver = pd.concat([pos_sample, neg_sample]).sample(frac=1, random_state=seed).reset_index(drop=True)
    print(f"    Final set: {len(silver)} reviews  "
          f"({(silver['silver_label']=='positive').sum()} pos / "
          f"{(silver['silver_label']=='negative').sum()} neg)")
    return silver
 
 
# =============================================================================
# 8. TABLE IV — Four-system classifier comparison (H3)
# =============================================================================
def macro_metrics(y_true, y_pred, labels=("positive","negative")):
    acc = accuracy_score(y_true, y_pred)
    report = classification_report(y_true, y_pred, labels=list(labels),
                                   output_dict=True, zero_division=0)
    macro_f1 = f1_score(y_true, y_pred, labels=list(labels),
                        average="macro", zero_division=0)
    return acc, macro_f1, report
 
 
def bootstrap_accuracy_ci(y_true, y_pred, n_boot=1000, seed=SEED):
    rng = np.random.default_rng(seed)
    accs = []
    y_true = np.array(y_true)
    y_pred = np.array(y_pred)
    for _ in range(n_boot):
        idx = rng.integers(0, len(y_true), len(y_true))
        accs.append(accuracy_score(y_true[idx], y_pred[idx]))
    return np.percentile(accs, [2.5, 97.5])
 
 
def table_iv(silver: pd.DataFrame):
    print("\n" + "="*70)
    print("TABLE IV — Classifier Performance on 500-Review Annotation Set (H3)")
    print("="*70)
 
    y_true_all = silver["silver_label"].values
 
    results = {}
 
    # ── 8a. Majority-class baseline ──────────────────────────────────────────
    y_majority = np.full(len(silver), "positive")
    acc_maj, f1_maj, rep_maj = macro_metrics(y_true_all, y_majority)
    results["Majority class"] = {
        "Accuracy": acc_maj, "Macro F1": f1_maj,
        "report": rep_maj, "y_pred": y_majority,
        "Training Data": "None"
    }
 
    # ── 8b. TextBlob (zero-shot) ─────────────────────────────────────────────
    tb_labels = silver["textblob_label"].values
    # Filter abstentions for metric calculation (as per paper)
    mask_tb = tb_labels != "abstain"
    y_true_tb = y_true_all[mask_tb]
    y_pred_tb = tb_labels[mask_tb]
    acc_tb, f1_tb, rep_tb = macro_metrics(y_true_tb, y_pred_tb)
    results["TextBlob"] = {
        "Accuracy": acc_tb, "Macro F1": f1_tb,
        "report": rep_tb, "y_pred": y_pred_tb,
        "y_true": y_true_tb,
        "Training Data": "None"
    }
 
    # ── 8c. VADER (zero-shot, all 500) ───────────────────────────────────────
    y_vader = silver["vader_label"].values
    # Map neutral → positive for binary eval (VADER rarely abstains on this set)
    y_vader_bin = np.where(y_vader == "negative", "negative", "positive")
    acc_v, f1_v, rep_v = macro_metrics(y_true_all, y_vader_bin)
    results["VADER"] = {
        "Accuracy": acc_v, "Macro F1": f1_v,
        "report": rep_v, "y_pred": y_vader_bin,
        "Training Data": "None"
    }
 
    # ── 8d. Logistic Regression (supervised) ─────────────────────────────────
    X = silver["clean_text"].values
    y = silver["silver_label"].values
 
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=1-TRAIN_FRAC, random_state=SEED, stratify=y
    )
    print(f"\n  LR split → train: {len(X_train)} | test: {len(X_test)}")
 
    tfidf = TfidfVectorizer(
        max_features=TFIDF_MAX,
        ngram_range=(1, 2),
        min_df=2,
        sublinear_tf=True
    )
    X_train_v = tfidf.fit_transform(X_train)
    X_test_v  = tfidf.transform(X_test)
 
    lr = LogisticRegression(C=LR_C, solver="lbfgs", max_iter=1000,
                            random_state=SEED, class_weight=None)
    lr.fit(X_train_v, y_train)
    y_pred_lr = lr.predict(X_test_v)
 
    acc_lr, f1_lr, rep_lr = macro_metrics(y_test, y_pred_lr)
    ci_lr = bootstrap_accuracy_ci(y_test, y_pred_lr, BOOTSTRAP, SEED)
    results["Logistic Regression"] = {
        "Accuracy": acc_lr, "Macro F1": f1_lr,
        "report": rep_lr, "y_pred": y_pred_lr,
        "y_true": y_test,
        "Training Data": f"{len(X_train)} labeled",
        "CI": ci_lr
    }
 
    # ── Print summary table ──────────────────────────────────────────────────
    header = f"{'System':<22} {'Acc':>6} {'MacF1':>6} {'P+':>5} {'R+':>5} {'F1+':>5} {'P-':>5} {'R-':>5} {'F1-':>5}  Training"
    print("\n" + header)
    print("-" * len(header))
 
    for name, res in results.items():
        r = res["report"]
        pos = r.get("positive", {})
        neg = r.get("negative", {})
        print(f"{name:<22} "
              f"{res['Accuracy']*100:>5.1f}% "
              f"{res['Macro F1']:>6.3f} "
              f"{pos.get('precision',0):>5.2f} "
              f"{pos.get('recall',0):>5.2f} "
              f"{pos.get('f1-score',0):>5.3f} "
              f"{neg.get('precision',0):>5.2f} "
              f"{neg.get('recall',0):>5.2f} "
              f"{neg.get('f1-score',0):>5.3f}  "
              f"{res['Training Data']}")
 
    # CI for LR
    ci = results["Logistic Regression"]["CI"]
    print(f"\n  LR 95% CI (bootstrap, {BOOTSTRAP} iters): [{ci[0]*100:.1f}%, {ci[1]*100:.1f}%]")
 
    # McNemar's test: VADER vs LR on shared test set
    # align predictions on the LR test indices
    test_idx = silver.index[silver.index.isin(
        silver.sample(frac=1, random_state=SEED).index[:100]
    )]
    # rebuild: get the actual test indices from the split
    _, X_test_raw, _, y_test_raw = train_test_split(
        np.arange(len(silver)), y, test_size=1-TRAIN_FRAC,
        random_state=SEED, stratify=y
    )
    y_vader_test = y_vader_bin[X_test_raw]
    y_lr_test    = y_pred_lr
    y_true_test  = y_test_raw  # already the true labels
 
    # contingency table for McNemar: (VADER wrong & LR right), (VADER right & LR wrong)
    vader_correct = (y_vader_test == y_true_test)
    lr_correct    = (y_lr_test    == y_true_test)
    b = np.sum(~vader_correct & lr_correct)   # VADER wrong, LR right
    c = np.sum(vader_correct  & ~lr_correct)  # VADER right, LR wrong
    # McNemar with continuity correction
    chi2 = (abs(b - c) - 1)**2 / (b + c) if (b + c) > 0 else 0
    p_mcnemar = stats.chi2.sf(chi2, df=1)
    print(f"\n  McNemar's test (VADER vs LR on 100-review test set):")
    print(f"    b (VADER wrong, LR right) = {b},  c (VADER right, LR wrong) = {c}")
    print(f"    χ² = {chi2:.2f},  p = {p_mcnemar:.3f}")
 
    # Confusion matrices
    for name in ["VADER", "Logistic Regression"]:
        res = results[name]
        yt = res.get("y_true", y_true_all)
        yp = res["y_pred"]
        if name == "Logistic Regression":
            yt = results["Logistic Regression"]["y_true"]
        cm = confusion_matrix(yt, yp, labels=["positive","negative"])
        print(f"\n  Confusion Matrix — {name}:")
        print(f"    {'':>10} Pred Pos  Pred Neg")
        print(f"    {'True Pos':>10}  {cm[0,0]:>7}  {cm[0,1]:>7}")
        print(f"    {'True Neg':>10}  {cm[1,0]:>7}  {cm[1,1]:>7}")
 
    return results, lr, tfidf
 
 
# =============================================================================
# 9. MAIN
# =============================================================================
def main():
    print("="*70)
    print(" Reproducing: Choudhary & Dutt — Sentiment-Based Analysis of")
    print(" Online Customer Reviews and Their Influence on Product Ratings")
    print("="*70)
 
    df = load_data(DATASET_PATH, SAMPLE_N, SEED)
    df = preprocess(df)
    df = score_sentiment(df)
 
    t1         = table_i(df)
    rho, r, r2 = table_ii(df)
    groups     = table_iii(df)
    silver     = build_silver_set(df, n_each=SILVER_N//2, seed=SEED)
    # Re-attach scores to silver set
    silver = silver.merge(
        df[["review_text", "vader_compound", "vader_label",
            "textblob_label", "clean_text"]],
        on="review_text", how="left", suffixes=("", "_dup")
    )
    for col in ["vader_compound","vader_label","textblob_label","clean_text"]:
        dup = col + "_dup"
        if dup in silver.columns:
            silver[col] = silver[col].combine_first(silver[dup])
            silver.drop(columns=[dup], inplace=True)
 
    results, lr_model, tfidf_vec = table_iv(silver)
 
    print("\n" + "="*70)
    print(" ALL RESULTS REPRODUCED. Check numbers against Tables I-IV in paper.")
    print("="*70)
 
    # Optional: save silver set and predictions for manual inspection
    silver.to_csv("silver_standard_set.csv", index=False)
    print("\n  Silver-standard set saved to: silver_standard_set.csv")
 
 
if __name__ == "__main__":
    main()