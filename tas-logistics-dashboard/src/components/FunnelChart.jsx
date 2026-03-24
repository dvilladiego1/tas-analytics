export default function FunnelChart({ data }) {
  if (!data || data.length === 0) return null;
  const maxVal = data[0]?.value || 1;

  return (
    <div className="chart-card">
      <h3>Funnel General</h3>
      <div className="funnel-container">
        {data.map((item, i) => {
          const pct = maxVal > 0 ? (item.value / maxVal) * 100 : 0;
          const conversionRate = i > 0 && maxVal > 0
            ? ((item.value / maxVal) * 100).toFixed(1)
            : null;
          return (
            <div key={item.label} className="funnel-row">
              <span className="funnel-label">{item.label}</span>
              <div className="funnel-bar-track">
                <div
                  className="funnel-bar-fill"
                  style={{
                    width: `${Math.max(pct, 2)}%`,
                    background: item.color,
                  }}
                >
                  <span className="funnel-bar-value">
                    {item.value.toLocaleString()}
                  </span>
                </div>
              </div>
              {conversionRate && (
                <span className="funnel-conversion">{conversionRate}%</span>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}
