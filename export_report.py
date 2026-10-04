"""Write a browser report of the current IBKR analysis. No Streamlit required to open it."""

from __future__ import annotations

import html
from datetime import datetime
from pathlib import Path

import pandas as pd

from app import (
    DEFAULT_CAD_PER_USD,
    aggregate_fills,
    convert_executions,
    cumulative_figure,
    distribution_figure,
    format_factor,
    format_money,
    format_percent,
    hold_bucket_figure,
    hold_bucket_stats,
    hold_outcome_figure,
    leveraged_net_pnl,
    load_saved_flex,
    match_fifo,
    summarize,
    ticker_totals,
    top_drivers_figure,
    _mean_hold_seconds,
    format_duration,
)

REPORT_PATH = Path(__file__).resolve().parent / "trade_analytics_report.html"


def prepare_view(executions: pd.DataFrame, currency: str, cad_per_usd: float):
    aggregated = aggregate_fills(executions)
    converted, _notes = convert_executions(aggregated, currency, cad_per_usd)
    matched = match_fifo(converted)
    closed = matched.closed
    commissions = float(converted["commission"].sum()) if converted is not None and not converted.empty else 0.0
    summary = summarize(closed, commissions)
    return closed, summary


def _chart(figure, div_id: str, include_js: bool) -> str:
    if figure is None:
        return ""
    return figure.to_html(
        full_html=False,
        include_plotlyjs="cdn" if include_js else False,
        div_id=div_id,
        config={"displayModeBar": False, "responsive": True},
    )


def _card(title: str, value: str, tone: str = "") -> str:
    color = ""
    if tone == "good":
        color = "color:#1B8A5A;"
    elif tone == "bad":
        color = "color:#D64545;"
    return (
        f'<div class="card"><div class="card-title">{html.escape(title)}</div>'
        f'<div class="card-value" style="{color}">{html.escape(value)}</div></div>'
    )


def _ticker_table(totals: pd.DataFrame, currency: str) -> str:
    if totals is None or totals.empty:
        return "<p>No closed trades.</p>"
    rows = []
    for record in totals.itertuples(index=False):
        win_rate = "—" if pd.isna(record[3]) else f"{float(record[3]):.1f}%"
        tone = "good" if record[1] > 0 else "bad" if record[1] < 0 else ""
        rows.append(
            "<tr>"
            f"<td>{html.escape(str(record[0]))}</td>"
            f'<td class="{tone}">{html.escape(format_money(float(record[1]), currency))}</td>'
            f"<td>{int(record[2])}</td>"
            f"<td>{win_rate}</td>"
            f"<td>{html.escape(format_money(float(record[4]), currency))}</td>"
            "</tr>"
        )
    return (
        "<table><thead><tr><th>Symbol</th><th>Net PnL</th><th>Trades</th>"
        "<th>Win rate</th><th>Commissions</th></tr></thead><tbody>"
        + "".join(rows)
        + "</tbody></table>"
    )


def render_html(closed: pd.DataFrame, summary, currency: str, cad_per_usd: float, source_label: str) -> str:
    totals = ticker_totals(closed)
    leveraged = leveraged_net_pnl(closed)
    pnl_tone = "good" if summary.net_pnl >= 0 else "bad"
    lev_tone = "good" if leveraged >= 0 else "bad"
    cards = "".join(
        [
            _card("Net realized PnL", format_money(summary.net_pnl, currency), pnl_tone),
            _card("Commissions & fees", format_money(summary.commissions, currency)),
            _card("Win rate", format_percent(summary.win_rate)),
            _card("Profit factor", format_factor(summary.profit_factor)),
            _card("Expectancy", format_money(summary.expectancy, currency)),
            _card("Average winning trade", format_money(summary.avg_win, currency)),
            _card("Average losing trade", format_money(summary.avg_loss, currency)),
            _card("Max drawdown", format_money(summary.max_drawdown, currency), "bad"),
            _card("Payoff ratio", format_factor(summary.payoff_ratio)),
            _card("Return / max drawdown", format_factor(summary.return_on_drawdown)),
            _card("Commission drag", format_percent(summary.commission_drag)),
            _card("Max consecutive losses", f"{summary.max_consecutive_losses:,}"),
            _card("Leveraged net PnL", format_money(leveraged, currency), lev_tone),
            _card("Closed trades", f"{summary.closed_count:,}"),
            _card("Avg winner hold", format_duration(_mean_hold_seconds(closed, True))),
            _card("Avg loser hold", format_duration(_mean_hold_seconds(closed, False))),
        ]
    )
    charts = [
        _chart(cumulative_figure(closed, currency), "cumulative", True),
        _chart(top_drivers_figure(totals, 5, currency), "drivers", False),
        _chart(distribution_figure(summary), "distribution", False),
        _chart(hold_outcome_figure(closed), "hold-outcome", False),
        _chart(hold_bucket_figure(hold_bucket_stats(closed), currency), "hold-buckets", False),
    ]
    chart_html = "".join(f'<div class="card chart">{block}</div>' for block in charts if block)
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    rate = f"{float(cad_per_usd):.2f}"
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>IBKR Trade Analysis</title>
<style>
body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; margin: 0; padding: 24px; color: #1c2430; background: #f6f7f9; }}
h1 {{ margin: 0 0 4px; font-size: 1.6rem; }}
.sub {{ color: #5c6770; margin: 0 0 20px; }}
.grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; margin-bottom: 20px; }}
.charts {{ display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }}
.card {{ background: #fff; border: 1px solid #e2e6ea; border-radius: 12px; padding: 14px 16px; }}
.chart {{ min-height: 280px; }}
.card-title {{ font-size: 0.8rem; color: #5c6770; margin-bottom: 6px; }}
.card-value {{ font-size: 1.35rem; font-weight: 650; font-variant-numeric: tabular-nums; }}
.good {{ color: #1B8A5A; }}
.bad {{ color: #D64545; }}
table {{ width: 100%; border-collapse: collapse; font-size: 0.92rem; }}
th, td {{ text-align: left; padding: 6px 8px; border-bottom: 1px solid #e2e6ea; }}
th {{ color: #5c6770; font-weight: 600; }}
@media (max-width: 800px) {{ .charts {{ grid-template-columns: 1fr; }} }}
</style>
</head>
<body>
<h1>IBKR Trade Analysis</h1>
<p class="sub">{html.escape(source_label)} · shown in {html.escape(currency)} at {rate} CAD per 1 USD · snapshot {stamp}. Charts need an internet connection. This file is not connected to the broker.</p>
<div class="grid">{cards}</div>
<div class="charts">{chart_html}</div>
<div class="card" style="margin-top:12px">
<div class="card-title">Tickers</div>
{_ticker_table(totals, currency)}
</div>
</body>
</html>
"""


def write_friend_report(
    closed: pd.DataFrame,
    summary,
    currency: str,
    cad_per_usd: float,
    source_label: str,
    path: Path = REPORT_PATH,
) -> Path:
    path.write_text(render_html(closed, summary, currency, cad_per_usd, source_label), encoding="utf-8")
    return path


def main() -> None:
    loaded = load_saved_flex()
    if loaded is None:
        raise SystemExit("No saved sync in data/flex. Sync in the app first.")
    executions, _notes = loaded
    closed, summary = prepare_view(executions, "CAD", DEFAULT_CAD_PER_USD)
    path = write_friend_report(closed, summary, "CAD", DEFAULT_CAD_PER_USD, "IBKR Flex Query")
    print(f"Wrote {path}")


if __name__ == "__main__":
    main()
