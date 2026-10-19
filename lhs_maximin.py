"""
LHS Maximin — Antena Mikrostrip U-Slot + DGS Dumbbell
Baseline: 3,5 GHz, FR-4 (εr = 4,3, h = 1,6 mm)
Kompatibel Python 3.12+ (scipy.stats.qmc)

Output:
- lhs_samples.csv     : 60 kombinasi parameter siap simulasi CST
- lhs_config.json     : metadata rentang & baseline
- lhs_pairplot.png    : visualisasi sebaran sampel
"""

import json
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.stats.qmc import LatinHypercube
from scipy.spatial.distance import pdist

# ============================================================
# 1. PARAMETER & RENTANG
# ============================================================
param_names = ['Lp','Wp','Lu','Wu','gU','PU','LI','LD','WD','gD','PD']

baseline = {
    'Lp': 20.22, 'Wp': 26.33, 'Lu': 10.0, 'Wu': 10.5, 'gU': 0.6,
    'PU': 0.0,   'LI': 6.15,  'LD': 3.0,  'WD': 3.0,  'gD': 0.3,
    'PD': 0.0
}

lower = np.array([18.0, 24.0,  7.0, 7.5, 0.3, -2.0, 4.0, 1.5, 1.5, 0.2, -3.0])
upper = np.array([22.0, 28.5, 13.0, 13.5, 1.2, 2.0, 8.5, 4.5, 4.5, 0.8, 3.0])

N_SAMPLES    = 60
N_CANDIDATES = 50
SEED         = 42

# ============================================================
# 2. GENERATE KANDIDAT LHS & SELEKSI MAXIMIN
# ============================================================
print(f"Mencari Maximin LHS dari {N_CANDIDATES} kandidat...")
rng = np.random.default_rng(SEED)

best_X       = None
best_min_dist = -np.inf

for i in range(N_CANDIDATES):
    sampler = LatinHypercube(d=len(param_names), seed=rng.integers(0, 10**6))
    X_unit  = sampler.random(n=N_SAMPLES)
    X_scaled = lower + X_unit * (upper - lower)
    # Normalisasi 0-1 agar jarak antar dimensi adil
    X_norm  = (X_scaled - lower) / (upper - lower)
    min_dist = pdist(X_norm).min()
    if min_dist > best_min_dist:
        best_min_dist = min_dist
        best_X        = X_scaled

print(f"  Jarak minimum terbaik: {best_min_dist:.4f}")

df = pd.DataFrame(best_X, columns=param_names)
df.insert(0, 'sample_id', [f"S{i+1:03d}" for i in range(N_SAMPLES)])

# ============================================================
# 3. VALIDASI GEOMETRI
# ============================================================
print("\nValidasi geometri:")
warnings = []

# U-Slot harus muat di dalam patch (margin 0.5 mm)
mask_slot = (df['Lu'] < df['Lp'] - 0.5) & (df['Wu'] < df['Wp'] - 0.5)
# DGS harus muat di ground (margin 1 mm)
mask_dgs  = (2*df['LD'] < 29.82 - 1) & (2*df['WD'] < 35.93 - 1)
# Penghubung DGS harus lebih kecil dari kepala
mask_gD   = df['gD'] < df[['LD','WD']].min(axis=1)

for name, mask in [('U-Slot keluar patch', mask_slot),
                   ('DGS keluar ground',  mask_dgs),
                   ('gD >= kepala DGS',   mask_gD)]:
    n_bad = int((~mask).sum())
    if n_bad > 0:
        warnings.append(f"{n_bad} sampel: {name}")

if warnings:
    print("  PERINGATAN:")
    for w in warnings:
        print(f"   - {w}")
else:
    print("  Semua sampel valid.")

# ============================================================
# 4. SIMPAN CSV & CONFIG
# ============================================================
df.to_csv('lhs_samples.csv', index=False)
print(f"\nTersimpan: lhs_samples.csv ({df.shape[0]} baris)")

config = {
    'method': 'Maximin LHS (scipy.stats.qmc)',
    'n_samples': N_SAMPLES,
    'n_candidates': N_CANDIDATES,
    'seed': SEED,
    'min_distance': float(best_min_dist),
    'baseline': baseline,
    'parameters': {
        name: {
            'baseline': baseline[name],
            'lower': float(lower[i]),
            'upper': float(upper[i]),
            'unit': 'mm'
        } for i, name in enumerate(param_names)
    }
}
with open('lhs_config.json', 'w') as f:
    json.dump(config, f, indent=2)
print("Tersimpan: lhs_config.json")

# ============================================================
# 5. STATISTIK & KORELASI
# ============================================================
print("\nStatistik parameter:")
print(df[param_names].describe().round(3))

# Korelasi antar parameter (fix: copy dulu agar writable)
corr     = df[param_names].corr().abs()
corr_arr = corr.to_numpy(copy=True)
np.fill_diagonal(corr_arr, 0)
max_corr = corr_arr.max()
print(f"\nKorelasi maks |r|: {max_corr:.4f} (target < 0.3)")

# Sampel terdekat dengan baseline
dist = ((df[param_names] - pd.Series(baseline))**2).sum(axis=1)
idx  = dist.idxmin()
print(f"\nSampel terdekat baseline: {df.loc[idx, 'sample_id']}")
print(df.loc[idx, param_names].round(3).to_string())

# ============================================================
# 6. PAIRPLOT
# ============================================================
n = len(param_names)
fig, axes = plt.subplots(n, n, figsize=(18, 18))
for i, p1 in enumerate(param_names):
    for j, p2 in enumerate(param_names):
        ax = axes[i, j]
        if i == j:
            ax.hist(df[p1], bins=12, color='steelblue', edgecolor='k')
            ax.axvline(baseline[p1], color='red', ls='--', lw=1)
        else:
            ax.scatter(df[p2], df[p1], s=12, alpha=0.7, color='coral')
            ax.axvline(baseline[p2], color='gray', ls=':', lw=0.5)
            ax.axhline(baseline[p1], color='gray', ls=':', lw=0.5)
        ax.tick_params(labelsize=6)
        if i == n - 1: ax.set_xlabel(p2, fontsize=8)
        if j == 0:     ax.set_ylabel(p1, fontsize=8)

plt.suptitle(f'Maximin LHS — {N_SAMPLES} sampel, baseline merah', fontsize=13)
plt.tight_layout(rect=[0, 0, 1, 0.98])
plt.savefig('lhs_pairplot.png', dpi=120)
print("\nTersimpan: lhs_pairplot.png")
print("Selesai.")