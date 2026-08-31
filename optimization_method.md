# Exact optimization audit

Every non-empty subset of the existing and candidate facilities was evaluated separately by city: 4,095 Pune subsets, 2,047 Hyderabad subsets and 1,023 Jaipur subsets. Order locations were used directly, not only pincode centroids. Orders were assigned to the nearest active facility. Scenarios with an active facility that received no orders in either month were excluded.

Because the brief does not provide a single numeric objective, the selected balanced strategy applies a transparent stability guardrail: expected breach improves by at least 2 percentage points and catchment coverage improves by at least 2 percentage points versus the current baseline in each city. Among those scenarios, minimise expected fully loaded cost-to-serve; use 36-month fit-out amortisation only as a tie-breaker.

The chosen network ranks first by both cost and capex-amortised cost among the guardrail-feasible scenarios in each city. It is also non-dominated if cost, expected breach and coverage are considered together.

A sensitivity audit at 0.5x, 1.0x, 1.5x and 2.0x of city-median candidate fixed cost leaves the selected network first under the guardrail in all three cities. Only an implausible 0x fixed-cost assumption changes the Pune choice.

A separate cost-floor option (CS-01 and CS-07 only; close PUN-02, PUN-07 and HYD-01; leave Jaipur unchanged) reaches ₹69.63/order with 75.01% coverage and 21.88% expected breach. It is retained as a sensitivity, not the base recommendation, because the balanced recommendation costs only ₹0.15/order more while improving coverage by 5.6 points and breach by 3.7 points.
