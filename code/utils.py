"""Plotting utilities for the rule-taking vs rule-making model."""
from __future__ import annotations

from pathlib import Path
from typing import Iterable, Tuple

import matplotlib.pyplot as plt
import numpy as np

FIG_DIR = Path(__file__).resolve().parent.parent / "paper" / "figures"
FIG_DIR.mkdir(parents=True, exist_ok=True)


def save_timeseries(
    t: np.ndarray,
    trajectories: np.ndarray,
    p: float,
    p_c: float,
    y0_values: Tuple[float, float],
    path: Path | None = None,
) -> Path:
    """Plot and save time-series trajectories for two initial conditions.

    Args:
        t: Time grid.
        trajectories: Array of shape (2, len(t)) containing y(t) for each initial condition.
        p: Reinforcement strength used in the simulation.
        p_c: Critical reinforcement threshold.
        y0_values: Initial conditions used for the two trajectories.
        path: Optional output path.

    Returns:
        Path to the saved figure.
    """

    if path is None:
        path = FIG_DIR / "timeseries.png"

    fig, ax = plt.subplots(figsize=(7, 4))
    colors = ["#1b9e77", "#d95f02"]
    for i, y0 in enumerate(y0_values):
        ax.plot(t, trajectories[i], label=f"y0={y0}", color=colors[i], lw=2)
    ax.axhline(0, color="k", lw=0.8, alpha=0.5)
    ax.axhline(1, color="k", lw=0.8, alpha=0.5)
    ax.set_xlabel("Time")
    ax.set_ylabel("Modifier frequency y")
    ax.set_title(f"Time evolution (p={p:.3f}, p_c={p_c:.3f})")
    ax.legend(frameon=False)
    ax.grid(True, alpha=0.3)
    fig.tight_layout()
    fig.savefig(path, dpi=300)
    plt.close(fig)
    return path


def save_bifurcation(
    p_values: np.ndarray,
    asym_low: np.ndarray,
    asym_high: np.ndarray,
    p_c: float,
    path: Path | None = None,
) -> Path:
    """Plot and save the bifurcation diagram from two initial conditions.

    Args:
        p_values: Array of reinforcement strengths.
        asym_low: Asymptotic y for the low initial condition.
        asym_high: Asymptotic y for the high initial condition.
        p_c: Critical reinforcement threshold.
        path: Optional output path.

    Returns:
        Path to the saved figure.
    """

    if path is None:
        path = FIG_DIR / "bifurcation.png"

    fig, ax = plt.subplots(figsize=(7, 4))
    ax.plot(p_values, asym_low, "o-", label="low initial y", color="#7570b3")
    ax.plot(p_values, asym_high, "s-", label="high initial y", color="#e7298a")
    ax.axvline(p_c, color="k", linestyle="--", lw=1, label="$p_c$")
    ax.set_xlabel("Reinforcement strength p")
    ax.set_ylabel("Asymptotic modifier frequency")
    ax.set_title("Bifurcation scan")
    ax.set_ylim(-0.05, 1.05)
    ax.legend(frameon=False)
    ax.grid(True, alpha=0.3)
    fig.tight_layout()
    fig.savefig(path, dpi=300)
    plt.close(fig)
    return path


__all__ = ["save_timeseries", "save_bifurcation", "FIG_DIR"]
