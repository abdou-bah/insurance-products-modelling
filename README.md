# Insurance Products Modelling (Solvency II)

Academic project in insurance modelling (Master Actuariat & Finance, University of Rouen Normandie, 2026). Excel models built from course templates and completed with my own calculations and VBA macros. The workbooks and the code comments are in French.

## Products
| Workbook | Product | Line of business |
|---|---|---|
| `models/Modele_Non_Vie_Sante.xlsx` | Health insurance, run-off of the existing portfolio, with quota-share reinsurance | Non-life health |
| `models/Modele_Temporaire_Deces.xlsx` | Annual renewable term life insurance (two model points, French TH/TF 00-02 mortality tables) | Life protection |
| `models/Modele_Rente_Conjoint.xlsx` | Temporary spouse's annuity paid after the death of the insured | Life annuity |

## What each model computes
- **Cash flow projection** over 50 years: premiums, claims (with settlement patterns), lapses, expenses, taxes, financial income.
- **Best Estimate**: cash flows within the contract boundary, discounted year by year with one-year forward rates derived from the zero-coupon curve.
- **PVFP** (present value of future profits): projected technical and financial results, discounted, before and after corporate tax.
- **SCR**:
  - health: premium and reserve risk with the Solvency II standard formula (medical expense, σ = 5% and 5.7%, correlation 0.5, SCR = 3σV);
  - term life: life underwriting risk with standard formula shocks (mortality +15%, lapse up ×1.5 and down ×0.5, expenses +10%, catastrophe +0.15 point of mortality), increases in Best Estimate aggregated with the life correlation matrix.
- **Stochastic interest rates**: Vasicek short-rate paths (2,000 for health and annuity, 100 for term life), each used to recompute the Best Estimate and the PVFP.

## VBA macros (`vba/`)
| Module | Macros |
|---|---|
| `Tempo_Deces.bas` | path simulation, Best Estimate and PVFP for each path, `Calcul_SCR_Vie` (shock loop for the life SCR) |
| `Non_Vie_Sante.bas` | path simulation, Best Estimate and PVFP for each path |
| `Rente_Conjoint.bas` | path simulation, Best Estimate and PVFP for each path |

PS / The workbooks are published as `.xlsx` with all results already computed.
| Product | Best Estimate | PVFP before tax | PVFP after tax | SCR |
|---|---|---|---|---|
| Health | 1,474.9 | 1,604.8 | 1,052.2 (664.3 net of reinsurance) | 415.0 |
| Term life | −23,883.3 | 276,266.8 | 181,148.1 | 1,166.5 |
| Spouse's annuity | 24,462.8 | 17,406.2 | 11,413.2 | not computed |

The term life Best Estimate is negative: within the one-year contract boundary, expected premiums exceed expected claims and expenses. The PVFP covers future renewals, which is why it is much larger. For term life, the catastrophe shock is the largest contributor to the SCR, ahead of lapse, expenses and mortality.

**Best Estimate under Vasicek scenarios**

| Product | Deterministic | Mean of scenarios | 5th to 95th percentile |
|---|---|---|---|
| Health | 1,474.9 | 1,456.0 | 1,359.9 to 1,563.2 |
| Term life | −23,883.3 | −23,883.8 | −23,888.6 to −23,878.9 |
| Spouse's annuity | 24,462.8 | 22,272.5 | 14,801.1 to 30,937.9 |

The longer the liabilities, the wider the distribution: the annuity is the most rate-sensitive product, term life (one-year boundary) almost not at all.

## Limits
- The Vasicek parameters (long-term mean 3.5%) are set by hand, not calibrated to the yield curve: the scenarios are not market-consistent, which is why the mean of the scenarios differs from the deterministic Best Estimate.
- The SCR is simplified: no mass lapse or expense inflation for term life, only premium and reserve risk for health, no SCR for the annuity, and capital is measured on the Best Estimate (no risk margin, no tax effect).
