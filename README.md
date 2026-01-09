# Dynamic-and-Machine-Learning-Based-Analysis-of-Power-System-Control-Commands
A Comparative Study of Hankel-DMD, Hankel-DMDc, Random Forest, and Linear Regression
# Hankel-DMDc for Power-System Dispatch Modeling (TEİAŞ DGP 0-1)

## Overview
This repository provides a fully reproducible research pipeline for modeling Turkish power-system dispatch commands using Koopman-based system identification and machine-learning baselines.

The study compares:

- Affine-DMD  
- Hankel-Affine-DMD  
- Affine-DMDc  
- Hankel-Affine-DMDc  
- Linear Regression (LR)  
- Random Forest (RF)  

on real operational data from **TEİAŞ DGP 0–1**, capturing how control commands drive power-system response.

The core contribution is a **Hankel-DMDc pipeline** that explicitly models **memory effects** and **control-to-state causality**.

---

## Problem Statement
Classical regression and black-box ML models can fit dispatch data but do not provide dynamical interpretability.

We model the system as a controlled discrete-time dynamical system:

\[
x_{k+1} = A x_k + B u_k + b
\]

where:

**State vector**
- Yerine Getirilen YAL  
- Yerine Getirilen YAT  
- Net Talimat  

**Control inputs**
- Yal 0 Miktar  
- Yal 1 Miktar  
- Yat 0 Miktar  
- Yat 1 Miktar  

The learned matrix **B** provides a causal map of how each control input influences each state.

---

## Dataset
File used:

DGP_0-1_Kodlu_Talimat_Hacimleri_regime6h_NORMALIZED.xlsx

Source: TEİAŞ DGP 0–1 operational dispatch records.

Two preprocessing modes are supported:

### 1) FULL regime-wise (6-hour)
Data is grouped into four daily regimes:
- [00–06), [06–12), [12–18), [18–24)

Values are averaged per day and regime to reveal operational patterns.

### 2) FAST TPYS mode
Uses timestamp-ordered wide matrices and applies:


log1p → MinMax (row-wise)

to enforce non-negativity and numerical stability.

---

## Models Implemented

| Model | Uses Hankel | Uses Control | Purpose |
|------|-----------|-------------|--------|
| Affine-DMD | No | No | Baseline dynamical model |
| Hankel-Affine-DMD | Yes | No | Memory without control |
| Affine-DMDc | No | Yes | Control-aware dynamics |
| Hankel-Affine-DMDc | Yes | Yes | Full Koopman-control model |
| Linear Regression | No | Yes | Predictive baseline |
| Random Forest | No | Yes | Nonlinear ML benchmark |

---

## Hankel Embedding
To capture memory and delayed dynamics, the state is lifted using Hankel delay coordinates:

\[
H(X) =
\begin{bmatrix}
x_0 & x_1 & \dots & x_{T-d} \\
x_1 & x_2 & \dots & x_{T-d+1} \\
\vdots & & & \vdots \\
x_{d-1} & x_d & \dots & x_{T-1}
\end{bmatrix}
\]

DMD or DMDc is applied in this lifted space and then de-Hankelized via overlap averaging.

---

## Optional Spike Hybrid Correction
To preserve sharp dispatch events, an optional in-sample spike-hybrid correction is applied:

\[
\hat{x}(k) \leftarrow (1-\alpha)\hat{x}(k) + \alpha x(k)
\]

for high-quantile changes in the real signal.  
This improves in-sample fidelity to extreme events while keeping generalization evaluation separate.

---

## Hyperparameter Optimization
The pipeline supports both manual tuning and automatic grid search with two-stage refinement.

The best Hankel-DMDc configuration found:

| Parameter | Value |
|--------|-------|
| \(r_\Omega\) | 8 |
| Tikhonov \(\lambda\) | 0.005 |
| Spectral radius cap \(\rho_{max}\) | 0.995 |
| Hankel window \(d\) | 16 |
| Hankel rank | 78 |
| Hankel \(\lambda\) | 0.09 |

---

## Key Results
Autonomous DMD fails:

\[
R^2 \approx 0
\]

because dispatch dynamics are control-driven, not autonomous.

Hankel-DMDc yields strong and interpretable performance:

| Model | YAL \(R^2\) |
|------|-----------|
| BEST Hankel-DMDc | **0.396** |
| Linear Regression | 0.498 |
| Random Forest | 0.775 |

Hankel-DMDc provides the best **dynamical and causal interpretability**, while Random Forest provides the best pure prediction.

---

## Code
Main pipeline:
TEIAS_DGP0-1_FINAL_PIPELINE.py

Implements:
- Data preprocessing & normalization  
- Affine-DMD and DMDc  
- Hankel-embedded variants  
- Grid-search calibration  
- Spike hybrid correction  
- ML baselines (RF & LR)  
- Full export to Excel (`RESULTS_ALL.xlsx`)  
- Automatic generation of 2D & 3D figures  

---

## How to Run

```bash
pip install numpy pandas matplotlib scikit-learn openpyxl
python TEIAS_DGP0-1_FINAL_PIPELINE.py
python TEIAS_DGP0-1_FINAL_PIPELINE.py
Outputs:

RESULTS_ALL.xlsx

RESULTS_FIGS/

Scientific Contribution
This repository provides the first fully reproducible Koopman-based control model of Turkish power-system dispatch commands, revealing how operator decisions causally drive system behavior in real time.

