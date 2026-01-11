# Optimal Portfolio Allocation (ECO358)

An Excel-based portfolio construction workflow that:
- Converts **daily stock returns → weekly returns** (by compounding within each week),
- Computes **expected returns**, **volatility**, and the **variance–covariance matrix**,
- Uses Excel **Solver** to find:
  - the **Optimal Risky Portfolio (ORP)** by **maximizing the Sharpe ratio**,
  - the **Minimum Variance Portfolio (MVP)** by **minimizing portfolio variance**,
  - a **no-short-selling ORP** (weights constrained to be nonnegative),
- Builds an **optimal capital allocation** between a risky portfolio and **3-month T-bills** using mean–variance utility.

## Repository Contents

- `Optimal-Portfolio-Allocation-Sheet.xlsx`  
  The main workbook (data aggregation, stats tables, covariance matrix, Solver setups, and charts).

- `Optimal-Portfolio-Allocation-Write-Up.docx`  
  The written explanation of methodology, results, and interpretation.

## Assets / Universe

This assignment uses six Canadian equities:
- `CNR`, `FTS`, `RY`, `SU`, `T`, `TRI`

## Method Overview

### 1) Daily → Weekly Returns
1. Map each trading day’s date to the **Monday of that week**.
2. For each ticker-week, compute the **weekly return by compounding daily gross returns**:

\[
R_{week} = \prod_{d \in week} (1 + R_d) - 1
\]

### 2) Summary Statistics
For each ticker (using weekly returns):
- Expected return: average of weekly returns  
- Volatility: standard deviation of weekly returns  
- Variance: volatility²  
- Covariances: pairwise covariances across tickers  
- Assemble the **variance–covariance matrix**.

### 3) Risk-Free Rate (Weekly)
Convert the 3-month T-bill rate to a weekly-equivalent rate (example conversion used in the write-up):

\[
r_f = \left(1 + (0.01 \cdot 90/365)\right)^{1/12} - 1
\]

*(Keep your workbook consistent: if your data frequency is weekly, your risk-free rate must be weekly too.)*

## Solver Workflows (Excel)

> **Enable Solver:** `File → Options → Add-ins → Excel Add-ins → Solver Add-in`

### A) Optimal Risky Portfolio (ORP): Max Sharpe Ratio
**Objective:** maximize Sharpe ratio  
\[
S = \frac{\mathbb{E}[R_p] - r_f}{\sigma_p}
\]

**Decision variables:** weights \( w_1,\dots,w_6 \)

**Constraints:**
- \( \sum_i w_i = 1 \)
- (Optional) allow short selling (weights can be negative)
- (Optional) no short selling: \( w_i \ge 0 \)

**Typical Solver settings:** GRG Nonlinear (works fine for this workbook structure)

#### ORP Weights (short selling allowed)
From the write-up’s Solver output:
| Ticker | Weight |
|---|---:|
| CNR | 0.173017 |
| FTS | -0.045621 |
| RY  | 0.301055 |
| SU  | -0.070839 |
| T   | 0.209139 |
| TRI | 0.433250 |

### B) Minimum Variance Portfolio (MVP)
**Objective:** minimize portfolio variance  
\[
\mathrm{Var}(R_p) = w^\top \Sigma w
\]

**Decision variables:** weights \( w \)  
**Constraint:** \( \sum_i w_i = 1 \)  
(Optional) add \( w_i \ge 0 \) for no-short MVP.

### C) No-Short-Selling ORP
Repeat the ORP setup, but add:
- \( w_i \ge 0 \) for all tickers.

## Capital Allocation (Risky Portfolio + T-Bills)

After you have a chosen risky portfolio (e.g., the no-short risky portfolio `NORP`), compute the optimal share in the risky asset:

\[
y^* = \frac{\mathbb{E}[R_{risky}] - r_f}{A \cdot \mathrm{Var}(R_{risky})}
\]

- \( A \) = risk aversion coefficient (example used: **A = 15**)
- Allocation:
  - Risky portfolio: \( y^* \)
  - T-bills: \( 1 - y^* \)

### Mean–Variance Utility
\[
U = \mathbb{E}[R] - \frac{1}{2}A \cdot \mathrm{Var}(R)
\]

You can compare utilities between:
- The optimal allocation using **short-selling allowed** risky portfolio, and
- The optimal allocation using **no-short-selling** risky portfolio.

A simple way to interpret the “value” of short selling is the **difference in utilities** (how much expected return you’d give up to be indifferent).

## Interpretation Notes (What to Expect)

- The ORP should typically sit **leftward** (lower risk) than many individual equities due to diversification.
- Lower covariances generally **increase diversification benefits** (lower portfolio risk, higher Sharpe; efficient frontier shifts outward).
- A higher risk-free rate typically **flattens the capital allocation line** and reduces the incentive to hold risky portfolios.

## How to Reproduce Results Quickly

1. Open `Optimal-Portfolio-Allocation-Sheet.xlsx`.
2. Ensure the weekly returns are computed (daily → Monday mapping + compounding).
3. Confirm the variance–covariance matrix is populated from weekly returns.
4. Set your weekly `r_f`.
5. Run Solver for:
   - ORP (maximize Sharpe),
   - MVP (minimize variance),
   - ORP with `w_i ≥ 0` (no short selling).
6. Use the risky portfolio stats to compute \( y^* \) and utility.

## Requirements

- Microsoft Excel with the **Solver Add-in** enabled.

## License

For course/academic use. If you plan to publish or reuse beyond coursework, add an explicit license (e.g., MIT).
