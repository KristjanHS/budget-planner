import { useState } from "react";

const phases = [
  {
    id: "proposal",
    label: "PROPOSAL PHASE",
    color: "#2D5A27",
    bgColor: "#E8F0E6",
    borderColor: "#2D5A27",
    icon: "📐",
    summary: "Bottom-up estimation → budget tables → submission",
    steps: [
      {
        title: "Estimate effort per partner",
        detail: "Each partner estimates PMs per work package, per staff category (senior researcher, engineer, dev, admin)",
        type: "action",
      },
      {
        title: "Calculate monthly rates",
        detail: "Annual gross employment cost ÷ 12 per staff category. Includes salary, social security, employer taxes.",
        type: "action",
      },
      {
        title: "Derive personnel EUR",
        detail: "PMs × monthly rate = personnel costs per partner",
        type: "calculation",
        formula: "personnel_eur = Σ(PM[i] × rate[i])",
      },
      {
        title: "Add other direct costs",
        detail: "Travel, equipment, subcontracting, other (software, dissemination, IP, audit cert)",
        type: "action",
      },
      {
        title: "Apply 25% indirect flat rate",
        detail: "On eligible direct costs excluding subcontracting",
        type: "calculation",
        formula: "indirect = 0.25 × (personnel + travel + equip + other)",
      },
      {
        title: "Apply programme funding rate",
        detail: "Innovation actions: up to 70%. The rest is partner co-financing.",
        type: "calculation",
        formula: "eu_grant = eligible_costs × funding_rate",
      },
      {
        title: "Consolidate into Part A + Part B",
        detail: "Part A: EUR bottom lines per partner. Part B: PM tables per WP per partner + narrative justification.",
        type: "output",
      },
    ],
  },
  {
    id: "grant",
    label: "GRANT AGREEMENT",
    color: "#8B4513",
    bgColor: "#FDF0E2",
    borderColor: "#8B4513",
    icon: "🔒",
    summary: "EUR ceilings per partner become legally binding",
    steps: [
      {
        title: "EUR ceiling locked per partner",
        detail: "Total eligible costs and EU grant amount per partner are fixed. This is the hard constraint for the entire project.",
        type: "constraint",
      },
      {
        title: "PM distribution recorded",
        detail: "PMs per WP per partner as stated in Part B. Serves as the baseline plan — flexible within EUR ceiling.",
        type: "output",
      },
      {
        title: "Funding rate confirmed",
        detail: "Per-partner funding rate is locked.",
        type: "constraint",
      },
    ],
  },
  {
    id: "kickoff",
    label: "PROJECT KICK-OFF",
    color: "#1A4A6E",
    bgColor: "#E4EDF4",
    borderColor: "#1A4A6E",
    icon: "🚀",
    summary: "Detailed planning within the locked EUR envelopes",
    steps: [
      {
        title: "Detailed resource scheduling",
        detail: "Map named people to tasks within WPs. Account for allocation % (e.g., 50% on this project, 50% on other projects).",
        type: "action",
      },
      {
        title: "Build internal timeline",
        detail: "Break WP-level PMs into task-level effort estimates with start/end dates and dependencies.",
        type: "action",
      },
      {
        title: "Validate against ceiling",
        detail: "Check: will planned PMs × actual rates fit within the partner EUR envelope? Leave margin for rate increases.",
        type: "validation",
      },
      {
        title: "Set up cost tracking",
        detail: "Time sheets, cost recording system. Each person logs hours/days to specific WPs.",
        type: "action",
      },
    ],
  },
  {
    id: "execution",
    label: "EXECUTION CYCLE",
    color: "#5B2C6F",
    bgColor: "#F0E6F6",
    borderColor: "#5B2C6F",
    icon: "🔄",
    summary: "Repeats per reporting period (typically 12–18 months)",
    steps: [
      {
        title: "People work, log time to WPs",
        detail: "Actual person-days recorded per WP. Monthly declarations or timesheets.",
        type: "action",
      },
      {
        title: "Incur and record other costs",
        detail: "Travel, equipment purchases, subcontracts — all attributed to specific WPs.",
        type: "action",
      },
      {
        title: "Monitor burn rate vs ceiling",
        detail: "Compare actual cumulative costs to partner EUR ceiling. Flag if on track to overspend or underspend.",
        type: "validation",
      },
      {
        title: "Reallocate PMs between WPs",
        detail: "Shift effort between work packages as needed. No amendment needed if within partner ceiling and work plan still delivered.",
        type: "flexible",
      },
      {
        title: "Adjust resource scheduling",
        detail: "React to delays, new hires, changed allocation %. Rebalance the plan within the ceiling.",
        type: "flexible",
      },
    ],
  },
  {
    id: "reporting",
    label: "PERIODIC REPORTING",
    color: "#8B0000",
    bgColor: "#FBE8E8",
    borderColor: "#8B0000",
    icon: "📊",
    summary: "Formal cost claim + technical progress report",
    steps: [
      {
        title: "Compile financial statements",
        detail: "Actual costs per partner, per category, with supporting docs. Costs must match time records.",
        type: "output",
      },
      {
        title: "Calculate indirect costs",
        detail: "25% flat rate on eligible direct costs (excl. subcontracting). Automatic — no actuals tracked.",
        type: "calculation",
        formula: "indirect_claim = 0.25 × (personnel + travel + equip + other)",
      },
      {
        title: "Technical report",
        detail: "Deliverables achieved, deviations from plan explained, PM consumption per WP reported.",
        type: "output",
      },
      {
        title: "Audit certificate if > €430k",
        detail: "Partners claiming over €430k total must submit an external CFS (Certificate on Financial Statements).",
        type: "constraint",
      },
      {
        title: "Review & payment",
        detail: "EC project officer reviews. May question costs. Approved amounts are reimbursed at the funding rate.",
        type: "action",
      },
    ],
  },
  {
    id: "amendment",
    label: "AMENDMENT (if needed)",
    color: "#666",
    bgColor: "#F2F2F2",
    borderColor: "#999",
    icon: "⚠️",
    summary: "Required only for changes that cross partner boundaries",
    steps: [
      {
        title: "Budget transfer between partners",
        detail: "Moving EUR from one partner to another requires formal grant agreement amendment.",
        type: "constraint",
      },
      {
        title: "Major scope changes",
        detail: "Adding/removing WPs, changing core deliverables, adding new partners — all need amendment.",
        type: "constraint",
      },
      {
        title: "NOT needed for",
        detail: "Shifting PMs between WPs within same partner. Moving budget between cost categories within same partner (unless structural).",
        type: "flexible",
      },
    ],
  },
  {
    id: "final",
    label: "FINAL REPORTING & CLOSE-OUT",
    color: "#1B4332",
    bgColor: "#D8F3DC",
    borderColor: "#1B4332",
    icon: "✅",
    summary: "Final cost claim, reconciliation, audit",
    steps: [
      {
        title: "Final financial statements",
        detail: "All remaining costs claimed. Total per partner must not exceed EUR ceiling from grant agreement.",
        type: "output",
      },
      {
        title: "Reconciliation",
        detail: "Compare total actual costs to budget. Underspend stays with EC. Overspend is absorbed by partner.",
        type: "validation",
      },
      {
        title: "Final technical report",
        detail: "All deliverables accounted for. PM actuals vs plan comparison.",
        type: "output",
      },
      {
        title: "Final payment",
        detail: "EC releases remaining grant after review. Possible ex-post audit within 5 years.",
        type: "action",
      },
    ],
  },
];

const typeStyles = {
  action: { bg: "#f8f9fa", border: "#dee2e6", badge: "#6c757d", badgeText: "#fff", label: "ACTION" },
  calculation: { bg: "#fff8e1", border: "#ffe082", badge: "#f57f17", badgeText: "#fff", label: "CALC" },
  output: { bg: "#e3f2fd", border: "#90caf9", badge: "#1565c0", badgeText: "#fff", label: "OUTPUT" },
  constraint: { bg: "#fce4ec", border: "#ef9a9a", badge: "#c62828", badgeText: "#fff", label: "HARD LIMIT" },
  validation: { bg: "#fff3e0", border: "#ffcc80", badge: "#e65100", badgeText: "#fff", label: "CHECK" },
  flexible: { bg: "#e8f5e9", border: "#a5d6a7", badge: "#2e7d32", badgeText: "#fff", label: "FLEXIBLE" },
};

function StepCard({ step, index }) {
  const style = typeStyles[step.type] || typeStyles.action;
  return (
    <div
      style={{
        background: style.bg,
        border: `1px solid ${style.border}`,
        borderRadius: 8,
        padding: "12px 14px",
        display: "flex",
        gap: 12,
        alignItems: "flex-start",
      }}
    >
      <div
        style={{
          minWidth: 26,
          height: 26,
          borderRadius: "50%",
          background: "#fff",
          border: `2px solid ${style.border}`,
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          fontFamily: "'Source Code Pro', 'Fira Code', monospace",
          fontSize: 12,
          fontWeight: 700,
          color: style.badge,
          flexShrink: 0,
          marginTop: 1,
        }}
      >
        {index + 1}
      </div>
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", marginBottom: 4 }}>
          <span style={{ fontWeight: 700, fontSize: 14, color: "#1a1a1a", fontFamily: "'DM Sans', sans-serif" }}>
            {step.title}
          </span>
          <span
            style={{
              fontSize: 9,
              fontWeight: 700,
              letterSpacing: "0.08em",
              background: style.badge,
              color: style.badgeText,
              padding: "2px 6px",
              borderRadius: 3,
              fontFamily: "'Source Code Pro', monospace",
            }}
          >
            {style.label}
          </span>
        </div>
        <p style={{ margin: 0, fontSize: 13, color: "#4a4a4a", lineHeight: 1.45, fontFamily: "'DM Sans', sans-serif" }}>
          {step.detail}
        </p>
        {step.formula && (
          <code
            style={{
              display: "inline-block",
              marginTop: 6,
              padding: "4px 10px",
              background: "rgba(0,0,0,0.06)",
              borderRadius: 4,
              fontSize: 12,
              fontFamily: "'Source Code Pro', 'Fira Code', monospace",
              color: "#333",
              letterSpacing: "0.02em",
            }}
          >
            {step.formula}
          </code>
        )}
      </div>
    </div>
  );
}

function PhaseCard({ phase, isOpen, onToggle }) {
  return (
    <div style={{ position: "relative" }}>
      {/* Connector line */}
      <div
        style={{
          position: "absolute",
          left: 24,
          top: -16,
          width: 2,
          height: 16,
          background: "#ccc",
        }}
      />
      <div
        style={{
          border: `2px solid ${phase.borderColor}`,
          borderRadius: 12,
          overflow: "hidden",
          background: "#fff",
          boxShadow: "0 2px 8px rgba(0,0,0,0.06)",
        }}
      >
        {/* Phase header */}
        <button
          onClick={onToggle}
          style={{
            width: "100%",
            display: "flex",
            alignItems: "center",
            gap: 12,
            padding: "14px 16px",
            background: phase.bgColor,
            border: "none",
            cursor: "pointer",
            textAlign: "left",
          }}
        >
          <span style={{ fontSize: 22 }}>{phase.icon}</span>
          <div style={{ flex: 1 }}>
            <div
              style={{
                fontFamily: "'Source Code Pro', 'Fira Code', monospace",
                fontSize: 11,
                fontWeight: 700,
                letterSpacing: "0.12em",
                color: phase.color,
                marginBottom: 2,
              }}
            >
              {phase.label}
            </div>
            <div style={{ fontSize: 13, color: "#555", fontFamily: "'DM Sans', sans-serif" }}>
              {phase.summary}
            </div>
          </div>
          <span
            style={{
              fontSize: 18,
              color: phase.color,
              transform: isOpen ? "rotate(180deg)" : "rotate(0deg)",
              transition: "transform 0.2s ease",
              fontWeight: 700,
            }}
          >
            ▾
          </span>
        </button>

        {/* Steps */}
        {isOpen && (
          <div style={{ padding: "12px 16px 16px", display: "flex", flexDirection: "column", gap: 8 }}>
            {phase.steps.map((step, i) => (
              <StepCard key={i} step={step} index={i} />
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

export default function GrantBudgetFlow() {
  const [openPhases, setOpenPhases] = useState({ proposal: true, grant: false, kickoff: false, execution: false, reporting: false, amendment: false, final: false });

  const toggle = (id) => setOpenPhases((prev) => ({ ...prev, [id]: !prev[id] }));
  const expandAll = () => {
    const all = {};
    phases.forEach((p) => (all[p.id] = true));
    setOpenPhases(all);
  };
  const collapseAll = () => {
    const all = {};
    phases.forEach((p) => (all[p.id] = false));
    setOpenPhases(all);
  };

  return (
    <div style={{ minHeight: "100vh", background: "#FAFAF8", padding: "24px 16px" }}>
      <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&family=Source+Code+Pro:wght@400;600;700&display=swap" rel="stylesheet" />

      <div style={{ maxWidth: 640, margin: "0 auto" }}>
        {/* Title */}
        <div style={{ marginBottom: 24 }}>
          <h1
            style={{
              fontFamily: "'Source Code Pro', monospace",
              fontSize: 15,
              fontWeight: 700,
              letterSpacing: "0.1em",
              color: "#888",
              margin: "0 0 4px",
              textTransform: "uppercase",
            }}
          >
            Programme Actual-Cost Model
          </h1>
          <h2
            style={{
              fontFamily: "'DM Sans', sans-serif",
              fontSize: 24,
              fontWeight: 700,
              color: "#1a1a1a",
              margin: "0 0 12px",
              lineHeight: 1.2,
            }}
          >
            Budget & Resource Scheduling Flow
          </h2>

          {/* Legend */}
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 12 }}>
            {Object.entries(typeStyles).map(([key, val]) => (
              <span
                key={key}
                style={{
                  fontSize: 9,
                  fontWeight: 700,
                  letterSpacing: "0.08em",
                  background: val.badge,
                  color: val.badgeText,
                  padding: "2px 7px",
                  borderRadius: 3,
                  fontFamily: "'Source Code Pro', monospace",
                }}
              >
                {val.label}
              </span>
            ))}
          </div>

          <div style={{ display: "flex", gap: 8 }}>
            <button
              onClick={expandAll}
              style={{
                padding: "6px 14px",
                fontSize: 12,
                fontFamily: "'DM Sans', sans-serif",
                fontWeight: 600,
                border: "1px solid #ccc",
                borderRadius: 6,
                background: "#fff",
                cursor: "pointer",
                color: "#333",
              }}
            >
              Expand all
            </button>
            <button
              onClick={collapseAll}
              style={{
                padding: "6px 14px",
                fontSize: 12,
                fontFamily: "'DM Sans', sans-serif",
                fontWeight: 600,
                border: "1px solid #ccc",
                borderRadius: 6,
                background: "#fff",
                cursor: "pointer",
                color: "#333",
              }}
            >
              Collapse all
            </button>
          </div>
        </div>

        {/* Flow */}
        <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
          {phases.map((phase, idx) => (
            <PhaseCard
              key={phase.id}
              phase={phase}
              isOpen={!!openPhases[phase.id]}
              onToggle={() => toggle(phase.id)}
            />
          ))}
        </div>

        {/* Key insight box */}
        <div
          style={{
            marginTop: 24,
            padding: "16px 18px",
            background: "#FFF8E1",
            border: "2px solid #FFD54F",
            borderRadius: 10,
          }}
        >
          <div
            style={{
              fontFamily: "'Source Code Pro', monospace",
              fontSize: 10,
              fontWeight: 700,
              letterSpacing: "0.1em",
              color: "#F57F17",
              marginBottom: 6,
            }}
          >
            KEY INSIGHT
          </div>
          <p style={{ margin: 0, fontSize: 14, color: "#333", lineHeight: 1.5, fontFamily: "'DM Sans', sans-serif" }}>
            <strong>EUR per partner is the ceiling. PMs are the lever.</strong> During execution, you're free to reallocate PMs across work packages within your partner envelope. The only hard triggers for a grant amendment are moving money between partners or changing total project scope.
          </p>
        </div>

        <p style={{ marginTop: 16, fontSize: 11, color: "#999", fontFamily: "'Source Code Pro', monospace", textAlign: "center" }}>
          Innovation Action · Budget & Resource Scheduling
        </p>
      </div>
    </div>
  );
}
