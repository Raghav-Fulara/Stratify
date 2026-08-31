# Project Dark Matter — reproducible analysis

This folder contains the engineered outputs used in the report. The analysis was run from `/home/user/uploads` and writes outputs to `/home/user/submission`.

Core definitions:
- distance_km: haversine distance from store to delivery point
- breached: actual_delivery_min > promised_delivery_min
- cost_to_serve_inr: 22 + 7*distance_km + 8 + 30*breached + monthly store fixed allocation
- serviceable_radius_km: first 0.5-km ring with breach rate >20%, upper ring edge
- after-network: nearest active facility; breach probability estimated from city-specific logistic breach-on-distance model
