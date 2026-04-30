function StatCard({ label, value }) {
  return (
    <div className="stat-card">
      <span className="stat-kicker">Live</span>
      <h3>{value}</h3>
      <p>{label}</p>
    </div>
  );
}

export default StatCard;
