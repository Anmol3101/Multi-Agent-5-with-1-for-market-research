#!/usr/bin/env python3
"""
HYPERMARKET 5+1 — General Market Research Agent

Usage:
    python agent.py
    python agent.py --product "Wireless earbuds" --market "Japan, South Korea" --depth standard
    python agent.py --product "Oat milk" --market "Germany" --output oat_milk_germany.xlsx

Requirements:
    pip install -r requirements.txt
    export ANTHROPIC_API_KEY=your_key_here
"""

import argparse
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

import anthropic

from excel_builder import build_excel

# ── HYPERMARKET 5+1 System Prompt ────────────────────────────────────────────

SYSTEM_PROMPT = r"""
You are a SINGLE model simulating a 5+1 multi-agent system for market research and go-to-market analysis.

The user will provide:
- PRODUCT_OR_TOPIC (string)
- TARGET_MARKET (string)
- DEPTH (optional: "auto" | "simple" | "standard" | "complex")

Your job in EVERY run:
- Internally simulate 5 research/analysis agents + 1 master writer.
- Enforce strict structured outputs.
- Return ONE JSON envelope that can drive analysis, Excel, and slides.

You MUST return EXACTLY ONE top-level JSON object shaped like:

{
  "status": "ok",
  "mode_used": "simple | standard | complex",
  "final_json": { ... },
  "xlsx_spec": { ... },
  "pptx_spec": { ... },
  "markdown_report": "string",
  "self_eval": {
    "run_quality": "low | medium | high",
    "bottlenecks": ["string"],
    "prompt_fix_suggestions": ["string"]
  }
}

No prose before or after this JSON.

====================================================
SHARED CONTRACT (ALL INTERNAL AGENTS)
====================================================

JSON RULES:
- Use null for unknown numeric values.
- When approximating, set "estimate": true.
- Values are short phrases, not long paragraphs.
- Hard caps: max_competitors=8, max_trends=8, max_risks=8, max_segments=6

CORE SCHEMA for final_json:

{
  "product": "string",
  "target_market": "string",
  "timestamp_utc": "string",
  "market_overview": {
    "segment_description": "string",
    "customer_segments": ["string"],
    "estimated_market_size": {
      "value": "number|null", "unit": "string", "currency": "USD",
      "year": "number|null", "estimate": true, "source": "string"
    },
    "estimated_growth_rate": {
      "cagr_percent": "number|null", "period": "string",
      "estimate": true, "source": "string"
    },
    "regional_breakdown": {
      "<market_name>": {
        "market_size_usd_million": "number|null",
        "cagr_percent": "number|null",
        "maturity": "emerging|growing|mature",
        "estimate": true
      }
    }
  },
  "competitors": [
    {
      "name": "string", "type": "brand|marketplace|aggregator|substitute",
      "positioning": "string", "price_band": "string",
      "channels": ["string"], "website": "string",
      "markets": ["string"]
    }
  ],
  "analysis": {
    "overall_demand_score": {
      "score_1_to_10": "number|null", "rationale": "string"
    },
    "segment_opportunities": [
      {
        "segment_name": "string",
        "need_or_pain_point": "string",
        "suggested_positioning": "string",
        "willingness_to_pay_1_to_10": "number|null",
        "target_market": "string"
      }
    ],
    "pricing_recommendation": {
      "<market_name>": {
        "band_label": "string", "currency": "string",
        "min": "number|null", "max": "number|null",
        "unit": "string", "sweet_spot": "string"
      }
    },
    "key_trends": [
      {
        "trend": "string",
        "relevance": "high|medium|low",
        "markets": ["string"]
      }
    ],
    "key_risks_summary": [
      {
        "description": "string",
        "level": "low|medium|high",
        "markets": ["string"]
      }
    ],
    "gtm_recommendations": {
      "<market_name>": {
        "priority": "number",
        "rationale": "string",
        "channels": ["string"],
        "strategy": "string",
        "timeline": "string"
      }
    }
  },
  "meta": {
    "notes": "string",
    "confidence": {
      "data_quality": "low|medium|high",
      "numeric_accuracy": "low|medium|high",
      "strategic_assessment": "low|medium|high"
    },
    "open_questions": ["string"]
  }
}

====================================================
ORCHESTRATOR LOGIC
====================================================

STEP 1 — Read PRODUCT_OR_TOPIC, TARGET_MARKET, DEPTH (default "auto").

STEP 2 — Classify complexity:
- If DEPTH != "auto": MODE = DEPTH
- Else:
  - "quick overview" or single narrow Q → MODE = "simple"
  - One product + one geography + GTM/pricing Q → MODE = "standard"
  - Multi-country, multi-segment, full strategy → MODE = "complex"

STEP 3 — Run all relevant agents internally and synthesize into the final JSON.

STEP 4 — Return ONLY the JSON object. No markdown fences. No prose.
""".strip()


# ── Input collection ──────────────────────────────────────────────────────────

def collect_inputs(args):
    print("\n╔══════════════════════════════════════════════════╗")
    print("║   HYPERMARKET 5+1 — Market Research Engine       ║")
    print("╚══════════════════════════════════════════════════╝\n")

    product = args.product or input("Product or Topic: ").strip()
    if not product:
        print("Error: product is required.")
        sys.exit(1)

    market = args.market or input("Target Market(s) (e.g. USA, India, Europe): ").strip()
    if not market:
        print("Error: target market is required.")
        sys.exit(1)

    depth = args.depth
    if not depth:
        depth = input("Depth [auto/simple/standard/complex] (default: auto): ").strip() or "auto"

    default_name = "{}_report.xlsx".format(
        re.sub(r"[^\w]", "_", product.lower())[:30]
    )
    output = args.output or input(f"Output file (default: {default_name}): ").strip() or default_name

    return product, market, depth, output


# ── Claude API call ───────────────────────────────────────────────────────────

def run_analysis(product: str, market: str, depth: str) -> dict:
    client = anthropic.Anthropic()

    user_message = json.dumps({
        "PRODUCT_OR_TOPIC": product,
        "TARGET_MARKET": market,
        "DEPTH": depth,
    }, indent=2)

    print(f"\n[Analyzing: {product} → {market} | depth={depth}]\n")
    print("─" * 60)

    full_text = ""
    with client.messages.stream(
        model="claude-opus-4-6",
        max_tokens=8000,
        thinking={"type": "adaptive"},
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_message}],
    ) as stream:
        for text in stream.text_stream:
            print(text, end="", flush=True)
            full_text += text

    print("\n" + "─" * 60)
    return full_text


def parse_response(raw: str) -> dict:
    raw = raw.strip()
    # Strip markdown fences if present
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        # Try to find the outermost JSON object
        match = re.search(r"\{[\s\S]*\}", raw)
        if match:
            try:
                return json.loads(match.group())
            except json.JSONDecodeError:
                pass
    print("\n[ERROR] Could not parse JSON from response. Raw output saved to debug_output.txt")
    Path("debug_output.txt").write_text(raw)
    sys.exit(1)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="HYPERMARKET 5+1 — Market Research Agent"
    )
    parser.add_argument("--product", "-p", help="Product or topic to research")
    parser.add_argument("--market",  "-m", help="Target market(s), e.g. 'USA, India'")
    parser.add_argument("--depth",   "-d",
                        choices=["auto", "simple", "standard", "complex"],
                        help="Analysis depth (default: auto)")
    parser.add_argument("--output",  "-o", help="Output .xlsx filename")
    args = parser.parse_args()

    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("Error: ANTHROPIC_API_KEY environment variable not set.")
        sys.exit(1)

    product, market, depth, output_file = collect_inputs(args)

    raw_response = run_analysis(product, market, depth)
    data = parse_response(raw_response)

    if data.get("status") != "ok":
        print(f"\n[ERROR] Analysis returned status: {data.get('status')}")
        sys.exit(1)

    mode = data.get("mode_used", "unknown")
    quality = data.get("self_eval", {}).get("run_quality", "unknown")
    print(f"\n[Mode: {mode} | Quality: {quality}]")

    final_json = data.get("final_json", {})
    if not final_json:
        print("[ERROR] No final_json in response.")
        sys.exit(1)

    build_excel(final_json, output_file)
    print(f"\n✓ Excel saved: {output_file}")

    # Print open questions if any
    oq = final_json.get("meta", {}).get("open_questions", [])
    if oq:
        print("\n── Open Questions ────────────────────────────────────")
        for i, q in enumerate(oq, 1):
            print(f"  {i}. {q}")

    # Print markdown report preview
    report = data.get("markdown_report", "")
    if report:
        preview = report[:500].strip()
        print(f"\n── Report Preview ────────────────────────────────────\n{preview}...\n")


if __name__ == "__main__":
    main()
