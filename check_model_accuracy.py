import json
import numpy as np
from sklearn.metrics import mean_absolute_error, r2_score
from scipy.stats import pearsonr

def load_data():
    # Load the reports from your project's data.json file
    with open("data.json", "r") as f:
        data = json.load(f)
    return data.get("answer_reports", [])

def evaluate():
    reports = load_data()
    
    ai_scores = []
    human_scores = []

    for r in reports:
        # manual_score must be added by you to data.json for comparison
        if "manual_score" in r:
            # AI score converted to percentage
            ai_val = (r['scored'] / r['total_marks']) * 100
            ai_scores.append(ai_val)
            human_scores.append(r["manual_score"])

    if not ai_scores:
        print("Error: No 'manual_score' found in data.json for comparison.")
        return

    # Statistical Metrics
    mae = mean_absolute_error(human_scores, ai_scores)
    correlation, _ = pearsonr(human_scores, ai_scores)
    accuracy = 100 - mae

    print(f"--- ML Model Evaluation ---")
    print(f"Overall Accuracy: {accuracy:.2f}%")
    print(f"Mean Absolute Error: {mae:.2f}%")
    print(f"Pearson Correlation: {correlation:.2f}")

    if correlation > 0.8:
        print("Status: Excellent Performance")
    elif correlation > 0.5:
        print("Status: Fair - Needs Prompt Refinement")
    else:
        print("Status: Poor - Check OCR Extraction Quality")

if __name__ == "__main__":
    evaluate()