"""Generate supplementary visuals for the defense deck."""
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np
import os

os.makedirs("ppt_assets/gen", exist_ok=True)

# Matched palette
NAVY     = "#0F2540"
TEAL     = "#1D7E7E"
AMBER    = "#E0B85F"
CORAL    = "#E56945"
INK      = "#1A1A1A"
GRAY     = "#7A7A7A"
GRAY_LT  = "#E5E5E0"

# ─── 1. LIF dynamics — for slide 3 ──────────────────────────────────
fig, ax = plt.subplots(figsize=(6.5, 3.0), dpi=200)
np.random.seed(42)
T = 80
inputs = np.zeros(T)
spike_t = [10, 25, 36, 48, 56, 65, 73]
for s in spike_t:
    inputs[s] = 1.2 * (0.8 + 0.4 * np.random.rand())

V_thr = 1.0
tau = 5
M = np.zeros(T); U = np.zeros(T); spikes = []
for t in range(1, T):
    M[t] = (1 - 1/tau) * U[t-1] + (1/tau) * inputs[t] * tau
    if M[t] > V_thr:
        spikes.append(t)
        U[t] = M[t] - V_thr
    else:
        U[t] = M[t]

ax.plot(np.arange(T), M, color=NAVY, lw=1.8, label="$M[t]$")
ax.axhline(y=V_thr, color=CORAL, ls="--", lw=1.5, label="$V_{\\rm thr}$")
for sp in spikes:
    ax.annotate("", xy=(sp, V_thr + 0.4), xytext=(sp, V_thr),
                arrowprops=dict(arrowstyle='->', color=AMBER, lw=2))
ax.set_xlim(0, T-1)
ax.set_ylim(-0.1, 1.7)
ax.set_xlabel("time t", fontsize=10)
ax.set_ylabel("membrane potential", fontsize=10)
ax.legend(loc="upper left", fontsize=10, frameon=False)
ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
ax.tick_params(axis="both", labelsize=9, colors=GRAY)
ax.text(spikes[0] - 1, V_thr + 0.55, "spike",
        fontsize=9, color=AMBER, ha="center")
plt.tight_layout()
plt.savefig("ppt_assets/gen/lif_dynamics.png", bbox_inches="tight",
            facecolor="white")
plt.close()
print("Saved lif_dynamics.png")


# ─── 2. SG window comparison (AS-SG vs RS-SG) — for slide 9 ────────
fig, axes = plt.subplots(1, 2, figsize=(8, 3.0), dpi=200)
xs = np.linspace(-2, 4, 400)

def density(x, mu, s=0.5):
    return np.exp(-(x - mu)**2 / (2*s**2)) / (s * np.sqrt(2*np.pi))

# Left: AS-SG  (window width = γ regardless of V_thr)
ax = axes[0]
mu = 1.5; gamma = 1.0
ax.fill_between(xs, density(xs, mu), color=NAVY, alpha=0.18,
                label="$M[t]$ density")
ax.plot(xs, density(xs, mu), color=NAVY, lw=1.5)
# AS-SG window: V_thr ± γ/2  with V_thr at mu
V_thr = mu
ax.axvline(V_thr, color=CORAL, ls=":", lw=1.2, label="$V_{\\rm thr}$")
rect = patches.Rectangle((V_thr - gamma/2, 0), gamma, 0.85,
                          linewidth=0, color=CORAL, alpha=0.30)
ax.add_patch(rect)
ax.text(V_thr, 0.95, "$\\gamma$ (fixed)", fontsize=9, ha="center", color=CORAL)
ax.set_title("AS-SG:  $x = M - V_{\\rm thr}$", fontsize=11, color=NAVY)
ax.set_xlim(-1.5, 4); ax.set_ylim(0, 1.15)
ax.set_xticks([]); ax.set_yticks([])
ax.spines[:].set_visible(False)
ax.legend(loc="upper right", fontsize=8, frameon=False)

# Right: RS-SG  (window scales with V_thr)
ax = axes[1]
mu = 1.5; gamma = 1.0
V_thr = mu
ax.fill_between(xs, density(xs, mu), color=NAVY, alpha=0.18)
ax.plot(xs, density(xs, mu), color=NAVY, lw=1.5)
ax.axvline(V_thr, color=AMBER, ls=":", lw=1.2)
# RS-SG window: V_thr · (1 ± γ/2)  - scales with V_thr
window_w = gamma * V_thr
rect = patches.Rectangle((V_thr - window_w/2, 0), window_w, 0.85,
                          linewidth=0, color=AMBER, alpha=0.45)
ax.add_patch(rect)
ax.text(V_thr, 0.95, "$\\gamma \\cdot V_{\\rm thr}$ (scales)",
        fontsize=9, ha="center", color="#A07020")
ax.set_title("RS-SG:  $x = M / V_{\\rm thr} - 1$", fontsize=11, color=NAVY)
ax.set_xlim(-1.5, 4); ax.set_ylim(0, 1.15)
ax.set_xticks([]); ax.set_yticks([])
ax.spines[:].set_visible(False)

plt.tight_layout()
plt.savefig("ppt_assets/gen/sg_windows.png", bbox_inches="tight",
            facecolor="white")
plt.close()
print("Saved sg_windows.png")


# ─── 3. Distribution convergence — for slide 6 ─────────────────────
fig, ax = plt.subplots(figsize=(6.5, 3.5), dpi=200)
xs = np.linspace(-2, 4, 400)
# Plot snapshots from U[0]=0 to π (centered around 1.5)
ts = [0, 1, 2, 3, 4, 5, 8, 12]
colors_step = plt.cm.RdYlBu(np.linspace(0.05, 0.95, len(ts)))
for i, t in enumerate(ts):
    # interpolate from delta-at-0 toward target Gaussian
    if t == 0:
        # start: very narrow at 0
        d = density(xs, 0.0, s=0.1) * 0.15
    else:
        # converge towards mu=1.5, s=0.5 over time
        progress = 1 - (1 - 0.5) ** t  # exponential
        cur_mu = 0.0 + (1.5 - 0.0) * progress
        cur_s  = 0.1 + (0.5 - 0.1) * progress
        d = density(xs, cur_mu, s=cur_s)
    label = f"$U[{t}]$" if t in [0, 1, 4, 12] else None
    ax.plot(xs, d, color=colors_step[i], lw=1.6, label=label)

# Target
d_target = density(xs, 1.5, 0.5)
ax.plot(xs, d_target, color=NAVY, lw=2.5, ls="--", label="$\\pi$ (stationary)")

ax.set_xlabel("membrane potential", fontsize=10)
ax.set_ylabel("density", fontsize=10)
ax.set_xlim(-1.5, 4); ax.set_ylim(0, 4.5)
ax.set_yticks([])
ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
ax.tick_params(axis="both", labelsize=9, colors=GRAY)
ax.legend(loc="upper right", fontsize=9, frameon=False)
ax.text(-1.0, 3.5, "drift!", fontsize=11, color=CORAL, fontweight="bold")
ax.annotate("", xy=(0.8, 2.8), xytext=(-0.5, 3.3),
            arrowprops=dict(arrowstyle="->", color=CORAL, lw=1.5))

plt.tight_layout()
plt.savefig("ppt_assets/gen/convergence.png", bbox_inches="tight",
            facecolor="white")
plt.close()
print("Saved convergence.png")


# ─── 4. Ablation bar chart — for slide 15 ──────────────────────────
fig, axes = plt.subplots(1, 2, figsize=(8, 3.4), dpi=200)
labels = ["baseline", "MP-Init", "TrSG", "both"]
colors = [GRAY, TEAL, AMBER, "#0E5E5E"]

# CIFAR-100
cifar = [75.55, 76.09, 77.15, 77.66]
ax = axes[0]
bars = ax.bar(labels, cifar, color=colors, width=0.65)
for b, v in zip(bars, cifar):
    ax.text(b.get_x() + b.get_width()/2, v + 0.05, f"{v:.2f}",
            ha="center", fontsize=10, color=NAVY,
            fontweight="bold" if v == max(cifar) else "normal")
ax.set_title("CIFAR-100  (T = 4)", fontsize=12, color=NAVY)
ax.set_ylim(74.5, 78.5)
ax.set_ylabel("Accuracy (%)", fontsize=10)
ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
ax.tick_params(axis="both", labelsize=9, colors=GRAY)

# DVS-CIFAR10
dvs = [76.60, 77.37, 80.83, 81.43]
ax = axes[1]
bars = ax.bar(labels, dvs, color=colors, width=0.65)
for b, v in zip(bars, dvs):
    ax.text(b.get_x() + b.get_width()/2, v + 0.07, f"{v:.2f}",
            ha="center", fontsize=10, color=NAVY,
            fontweight="bold" if v == max(dvs) else "normal")
ax.set_title("DVS-CIFAR10  (T = 10)", fontsize=12, color=NAVY)
ax.set_ylim(75.5, 82.5)
ax.set_ylabel("Accuracy (%)", fontsize=10)
ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
ax.tick_params(axis="both", labelsize=9, colors=GRAY)

plt.tight_layout()
plt.savefig("ppt_assets/gen/ablation_bars.png", bbox_inches="tight",
            facecolor="white")
plt.close()
print("Saved ablation_bars.png")


# ─── 5. Main results bar chart — for slide 14 ──────────────────────
fig, ax = plt.subplots(figsize=(8, 3.4), dpi=200)
datasets = ["CIFAR-10\n(R-19, T=6)", "CIFAR-100\n(R-19, T=6)",
            "ImageNet\n(SEW-R34, T=4)", "DVS-CIFAR10\n(R-19, T=10)"]
prev_best = [94.81, 76.82, 69.03, 76.20]
ours      = [95.50, 77.91, 69.67, 81.43]
prev_label = ["TAB", "TAB", "MPS", "RMP"]

x = np.arange(len(datasets))
w = 0.32
b1 = ax.bar(x - w/2, prev_best, w, color=GRAY,
            label="Previous best", edgecolor="white")
b2 = ax.bar(x + w/2, ours, w, color=TEAL,
            label="Ours (MP-Init + TrSG)", edgecolor="white")

for bar, v, lab in zip(b1, prev_best, prev_label):
    ax.text(bar.get_x() + bar.get_width()/2, v + 0.4, f"{lab}\n{v:.1f}",
            ha="center", va="bottom", fontsize=8.5, color=GRAY)
for bar, v, p in zip(b2, ours, prev_best):
    ax.text(bar.get_x() + bar.get_width()/2, v + 0.4,
            f"{v:.1f}\n+{v-p:.2f}",
            ha="center", va="bottom", fontsize=8.5,
            color=TEAL, fontweight="bold")

ax.set_xticks(x)
ax.set_xticklabels(datasets, fontsize=9, color=NAVY)
ax.set_ylabel("Top-1 Accuracy (%)", fontsize=10)
ax.set_ylim(60, 100)
ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
ax.tick_params(axis="y", labelsize=9, colors=GRAY)
ax.legend(loc="lower right", fontsize=10, frameon=False)

plt.tight_layout()
plt.savefig("ppt_assets/gen/main_results_bars.png", bbox_inches="tight",
            facecolor="white")
plt.close()
print("Saved main_results_bars.png")


# ─── 6. Chain-rule cancellation diagram — for slide 11 ─────────────
fig, ax = plt.subplots(figsize=(8, 2.2), dpi=200)
ax.set_xlim(0, 10); ax.set_ylim(0, 3)
ax.axis("off")

# Three boxes flowing
boxes = [
    (0.5, 1.0, "$\\frac{\\partial S}{\\partial M}$\n$\\approx \\frac{1}{V_{\\rm thr}} f'(x)$", CORAL),
    (4.0, 1.0, "$\\times V_{\\rm thr}$\n(forward)", AMBER),
    (7.5, 1.0, "$\\frac{\\partial O}{\\partial M}$\n$= f'(x)$", TEAL),
]
for (x, y, lab, color) in boxes:
    rect = patches.FancyBboxPatch((x, y), 2.0, 1.2,
                                    boxstyle="round,pad=0.05",
                                    linewidth=1.5, edgecolor=color,
                                    facecolor="white")
    ax.add_patch(rect)
    ax.text(x + 1.0, y + 0.6, lab, fontsize=12,
            ha="center", va="center", color=color)

# Arrows
ax.annotate("", xy=(4.0, 1.6), xytext=(2.5, 1.6),
            arrowprops=dict(arrowstyle="->", color=NAVY, lw=2.5))
ax.annotate("", xy=(7.5, 1.6), xytext=(6.0, 1.6),
            arrowprops=dict(arrowstyle="->", color=NAVY, lw=2.5))

# Labels under arrows
ax.text(3.25, 1.2, "RS-SG\nproblem", fontsize=8, ha="center", color=GRAY)
ax.text(6.75, 1.2, "factor cancels", fontsize=8, ha="center", color=TEAL)

# Title above
ax.text(5, 2.7, "Chain-rule cancellation",
        fontsize=12, ha="center", color=NAVY, fontweight="bold")

plt.tight_layout()
plt.savefig("ppt_assets/gen/chain_rule.png", bbox_inches="tight",
            facecolor="white")
plt.close()
print("Saved chain_rule.png")

print("\nAll generated:")
for f in os.listdir("ppt_assets/gen"):
    print(f"  ppt_assets/gen/{f}")


# ─── 7. No-training V_thr accuracy — for "Why TrSG Works" S14 ──────
fig, ax = plt.subplots(figsize=(7.5, 3.0), dpi=200)
v_thrs = ["0.1", "0.5", "1.0", "1.5", "2.0"]
ax_vals = [61.59, 75.96, np.nan, 69.97, 18.09]   # AS-SG
rs_vals = [np.nan, 76.12, np.nan, 71.43,  5.43]   # RS-SG (div. at 0.1)
tr_vals = [73.99, 76.82, 75.56, 71.97, 64.27]     # TrSG

x = np.arange(len(v_thrs))
w = 0.27
b1 = ax.bar(x - w, ax_vals, w, color=CORAL, label="AS-SG", edgecolor="white")
b2 = ax.bar(x,     rs_vals, w, color=AMBER, label="RS-SG", edgecolor="white")
b3 = ax.bar(x + w, tr_vals, w, color=TEAL,  label="TrSG",  edgecolor="white")

# Mark NaNs / divergences
for xi, val in zip(x, ax_vals):
    if np.isnan(val):
        ax.text(xi - w, 5, "—", ha="center", color=GRAY, fontsize=11)
for xi, val in zip(x, rs_vals):
    if np.isnan(val):
        ax.text(xi, 5, "div.", ha="center", color=CORAL, fontsize=9, fontweight="bold")
for bar, val in zip(b3, tr_vals):
    ax.text(bar.get_x() + bar.get_width()/2, val + 1, f"{val:.1f}",
            ha="center", color=TEAL, fontsize=8.5, fontweight="bold")
for bar, val in zip(b1, ax_vals):
    if not np.isnan(val):
        ax.text(bar.get_x() + bar.get_width()/2, val + 1, f"{val:.0f}",
                ha="center", color=GRAY, fontsize=8)
for bar, val in zip(b2, rs_vals):
    if not np.isnan(val):
        ax.text(bar.get_x() + bar.get_width()/2, val + 1, f"{val:.0f}",
                ha="center", color=GRAY, fontsize=8)

ax.set_xticks(x); ax.set_xticklabels([f"$V_{{\\rm thr}}={v}$" for v in v_thrs], fontsize=10)
ax.set_ylabel("Accuracy (%)", fontsize=10)
ax.set_ylim(0, 90)
ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
ax.tick_params(axis="y", labelsize=9, colors=GRAY)
ax.legend(loc="upper right", fontsize=10, frameon=False)
ax.set_title("CIFAR-100, ResNet-19, $\\tau=2.0$ frozen — $V_{\\rm thr}$ frozen at varying values",
             fontsize=11, color=NAVY)

plt.tight_layout()
plt.savefig("ppt_assets/gen/no_training_acc.png", bbox_inches="tight",
            facecolor="white")
plt.close()
print("Saved no_training_acc.png")
