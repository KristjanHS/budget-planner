// Sheet Flow Diagram — Budget Planner
// Render with: npx create-react-app tmp && copy this into App.js, or paste into https://playcode.io/react

import React from "react";

const colors = {
  red: { bg: "#FDEDEC", border: "#E74C3C", text: "#922B21" },
  default: { bg: "#F2F3F4", border: "#95A5A6", text: "#2C3E50" },
  orange: { bg: "#FEF5E7", border: "#F5A623", text: "#7E5109" },
  purple: { bg: "#F4ECF7", border: "#9B59B6", text: "#6C3483" },
  blue: { bg: "#D6EAF8", border: "#4A90D9", text: "#1B4F72" },
  green: { bg: "#EAFAF1", border: "#7ED321", text: "#1E8449" },
};

const Arrow = ({ from, to, label, dashed = false }) => (
  <div
    style={{
      display: "flex",
      alignItems: "center",
      gap: 8,
      margin: "4px 0",
      fontSize: 12,
      color: "#555",
    }}
  >
    <span style={{ fontWeight: 600, minWidth: 70, textAlign: "right" }}>
      {from}
    </span>
    <span
      style={{
        flex: 1,
        borderTop: dashed ? "2px dashed #aaa" : "2px solid #555",
        position: "relative",
        height: 0,
      }}
    >
      <span
        style={{
          position: "absolute",
          right: -4,
          top: -5,
          fontSize: 10,
        }}
      >
        ▶
      </span>
    </span>
    <span style={{ fontWeight: 600, minWidth: 70 }}>{to}</span>
    <span style={{ color: "#888", fontStyle: "italic", minWidth: 160 }}>
      {label}
    </span>
  </div>
);

const Sheet = ({ name, tabColor, children, phase }) => {
  const c = colors[tabColor] || colors.default;
  return (
    <div
      style={{
        border: `2px solid ${c.border}`,
        borderRadius: 8,
        background: c.bg,
        padding: 16,
        margin: 8,
        minWidth: 200,
        maxWidth: 280,
        boxShadow: "0 2px 4px rgba(0,0,0,0.1)",
      }}
    >
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          marginBottom: 8,
        }}
      >
        <h3 style={{ margin: 0, color: c.text, fontSize: 16 }}>{name}</h3>
        {phase && (
          <span
            style={{
              fontSize: 10,
              background: c.border,
              color: "#fff",
              padding: "2px 6px",
              borderRadius: 4,
            }}
          >
            {phase}
          </span>
        )}
      </div>
      <div
        style={{
          height: 3,
          background: c.border,
          borderRadius: 2,
          marginBottom: 8,
        }}
      />
      <div style={{ fontSize: 13, color: c.text, lineHeight: 1.5 }}>
        {children}
      </div>
    </div>
  );
};

const Section = ({ title, children }) => (
  <div style={{ marginBottom: 24 }}>
    <h4
      style={{
        color: "#555",
        fontSize: 12,
        textTransform: "uppercase",
        letterSpacing: 1,
        marginBottom: 8,
      }}
    >
      {title}
    </h4>
    <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
      {children}
    </div>
  </div>
);

export default function SheetFlowDiagram() {
  return (
    <div
      style={{
        fontFamily: '-apple-system, "Segoe UI", Roboto, sans-serif',
        maxWidth: 900,
        margin: "0 auto",
        padding: 24,
      }}
    >
      <h1 style={{ fontSize: 20, color: "#2C3E50", marginBottom: 4 }}>
        Budget Planner — Sheet Flow
      </h1>
      <p style={{ color: "#888", fontSize: 13, marginBottom: 24 }}>
        Programme Phases 2–7 · Single partner · Personnel costs +
        25% indirect
      </p>

      {/* Sheet cards */}
      <Section title="Documentation">
        <Sheet name="_META" tabColor="red">
          AI-agent documentation
          <br />
          Structure, layout coords, modification guide
        </Sheet>
      </Section>

      <Section title="Configuration & Planning">
        <Sheet name="Plan" tabColor="default" phase="Ph 2–3">
          <b>Global config:</b> title, start, rate, personnel ceiling
          <br />
          <b>WP config:</b> ID, name, PM budget, start/end, per-WP ceiling
          <br />
          <b>Planned FTE grid:</b> WP × month (from Capacity)
        </Sheet>
        <Sheet name="Team" tabColor="orange" phase="Ph 3">
          <b>Roster:</b> name, max FTE
          <br />
          <b>WP seeds:</b> person × WP share %
        </Sheet>
        <Sheet name="Capacity" tabColor="purple" phase="Ph 3–4">
          <b>Per-person grids:</b> monthly FTE per WP
          <br />
          <b>External commitments:</b> other project FTE
          <br />
          <b>Remaining FTE:</b> max − allocated
        </Sheet>
      </Section>

      <Section title="Execution & Tracking">
        <Sheet name="Actuals" tabColor="blue" phase="Ph 4–5">
          <b>EUR input grid:</b> WP × month
          <br />
          <i>Blue input — from accountant</i>
          <br />
          <b>Validation:</b> sum, ceiling, variance per WP
        </Sheet>
        <Sheet name="Summary" tabColor="green" phase="Ph 4–7">
          <b>Annual table:</b> Plan FTE, Plan EUR, Actual EUR, Variance
          <br />
          <b>Budget overview:</b>
          <br />
          &nbsp;&nbsp;Personnel Actual
          <br />
          &nbsp;&nbsp;Total Claimable (+25%)
          <br />
          &nbsp;&nbsp;Ceiling · Remaining · % Consumed
        </Sheet>
      </Section>

      {/* Data flow arrows */}
      <div
        style={{
          marginTop: 24,
          padding: 16,
          background: "#FAFAFA",
          borderRadius: 8,
          border: "1px solid #eee",
        }}
      >
        <h4
          style={{
            color: "#555",
            fontSize: 12,
            textTransform: "uppercase",
            letterSpacing: 1,
            marginBottom: 12,
          }}
        >
          Data Flow
        </h4>
        <Arrow from="Team" to="Capacity" label="person names, max FTE" />
        <Arrow
          from="Team"
          to="Capacity"
          label="WP share seeds (initial only)"
          dashed
        />
        <Arrow from="Plan" to="Capacity" label="WP names" />
        <Arrow
          from="Capacity"
          to="Plan"
          label="SUM per WP across people → Planned FTE"
        />
        <Arrow from="Plan" to="Actuals" label="WP structure, month range, ceilings" />
        <Arrow from="Plan" to="Summary" label="Plan FTE, Plan EUR, partner_ceiling" />
        <Arrow from="Actuals" to="Summary" label="Actual EUR per WP, total actual" />
      </div>

      {/* Legend */}
      <div
        style={{
          marginTop: 16,
          display: "flex",
          gap: 16,
          fontSize: 12,
          color: "#888",
        }}
      >
        <span>
          <span style={{ color: "#4A90D9" }}>■</span> Blue = user input
        </span>
        <span>
          <span style={{ color: "#2C3E50" }}>■</span> Black = formula
        </span>
        <span>
          <span style={{ color: "#27AE60" }}>■</span> Green = cross-sheet ref
        </span>
      </div>
    </div>
  );
}
