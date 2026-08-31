# QuickCart Project Dark Matter — submission package

## Main deliverables
- `final_report.docx` — recommendation memo with assumptions, maps, actions, economics and stage gates.
- `executive_deck.pptx` — 9-slide executive summary for the COO.
- `final_analysis.ipynb` — reproducible analysis companion.
- `dark_matter_analysis.py` — source script used to produce the outputs.
- `optimization_audit.py` — exact order-level network subset search.

## Data / evidence files
- `order_level_cleaned.csv` — all 186,156 orders with distance, breach, month-level fixed allocation, cost-to-serve and catchment flags.
- `after_order_assignment.csv` — exact order-level recommended-network assignment and estimated after metrics.
- `before_after_city_metrics.csv` / `before_after_portfolio_metrics.csv` — headline impact tables.
- `serviceable_radii.csv`, `candidate_site_scorecard.csv`, `recommended_facility_loads.csv`, `pincode_before_after_impact.csv`.
- `images/` — static maps, optimization frontier and charts required by the brief.
- `optimization_summary.csv` / `exhaustive_exact_scenarios.csv` — optimum-strategy audit.
- `fixed_cost_sensitivity.csv` — sensitivity of the selected network to candidate fixed-cost imputation.
- `cost_floor_before_after.csv` / `cost_floor_order_assignment.csv` — cost-first alternative retained as a transparent trade-off.

## Recommendation
Open CS-01 in Pune, CS-07 in Hyderabad and CS-11 in Jaipur; close PUN-02; relocate PUN-07 into CS-01. Candidate fixed cost is imputed at city median existing fixed cost because the candidate file does not provide it.

## Headline portfolio estimate
C2S ₹71.74 -> ₹69.78; breach 24.7% -> 18.2%; coverage 72.7% -> 80.6%; run-rate saving ₹183,006/month; simple fit-out payback 24.4 months.
