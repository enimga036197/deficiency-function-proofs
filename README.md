# Deficiency Function Research Instrument

A computational mathematics workbook for studying **d(n) = n − φ(n) − τ(n)**, where φ is Euler's totient and τ is the divisor-counting function.

## What this is

An Excel-based research instrument where **every claim is computed by Excel formulas, not pre-stated**. Charts discover structure. Conditional formatting flags anomalies. If a theorem fails, the spreadsheet shows it.

## The main file

**`Deficiency_Research_Instrument.xlsx`** — 8 sheets, ~3000 live formulas:

| Sheet | What it computes | Key result |
|-------|-----------------|------------|
| Computation Engine | d(n) for n=1..1000 via `=A-B-C`, τ(n) via SUMPRODUCT | 1000 rows, all predictions verified |
| Zero Set Proof | Exhaustive case analysis + 100K scan | d(n)=0 iff n ∈ {6,8,9} |
| Consecutive Coincidences | All pairs where d(n)=d(n+1) up to 100K | Only 7 pairs, all Fermat-related |
| Fiber Concentration | Top 200 fiber sizes, mod 30 analysis | Hardy-Littlewood concentration at k≡25(30) |
| Ray Structure | Multi-series scatter: primes/semiprimes/p²/composites | Visible ray separation by number type |
| Phase Transition | r(n) = (n−φ(n))/τ(n) for n=2..500 | r(n)=1/2 iff n is prime — sharp floor |
| GCD-Projection | cos²(θ\*)=1/φ bridge identity + 200 GCD homomorphism tests | Identity proven to machine precision |
| Parity Theorem | d(n) ≡ n (mod 2) iff n is not a perfect square | 498 values verified |

## Proven theorems (standalone, no external references)

1. **Zero Set**: d(n) = 0 ⟺ n ∈ {6, 8, 9}
2. **Spectrum Gap**: d(n) ≠ 1 for all n
3. **Prime Formula**: d(p) = −1 for all primes p
4. **Prime Power**: d(p^k) = p^(k−1) − k − 1
5. **Semiprime**: d(pq) = p + q − 5 for distinct primes p < q
6. **Isospectral Identity**: d(2p) = d(p²) = p − 3 for all odd primes p
7. **Phase Transition**: The parametric zero set Z(α) transitions sharply at α = 1/2
8. **Fiber Finiteness**: All fibers F_k are finite for k ≥ 0
9. **Parity**: d(n) ≡ n (mod 2) ⟺ n is not a perfect square
10. **Fermat Coincidence**: d(2^(2^m)−1) = d(2^(2^m)) for m = 1,2,3,4
11. **Squarefree Formula**: d(p₁⋯pₖ) = ∏pᵢ − ∏(pᵢ−1) − 2^k
12. **Average Order**: (1/N)∑ d(n)/n → 1 − 6/π² ≈ 0.3921
13. **Dirichlet Series**: ∑ d(n)/n^s = ζ(s−1) − ζ(s−1)/ζ(s) − ζ(s)²
14. **GCD Bridge Identity**: cos²(θ\*) = 1/φ where θ\* = arccos(1/√φ)
15. **Modular Reduction**: gcd(ab, n) = gcd(gcd(a,n)·gcd(b,n), n) for all positive a,b,n

## Gate for entry

Every claim in this workbook passes one test: **can it stand under its own mathematical justification?**

- If it's a theorem, it has a proof
- If it's a computation, Excel verifies it independently
- If it's a conjecture, it's clearly marked as such
- If it can be falsified, the workbook tries

## Build scripts

The Python scripts (`develop_v2.py` etc.) generate the Excel workbooks using `openpyxl`. The workbooks contain live Excel formulas — they are not static reports.

```bash
pip install openpyxl
python develop_v2.py
```

## Requirements

- Python 3.10+ with openpyxl
- Microsoft Excel (for viewing charts and live formulas)
