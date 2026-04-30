function Navbar({ user, onLogout, activeTab, setActiveTab }) {
  const baseTabs = [
    "overview",
    "appointments",
    "cart",
    "orders",
    "upload",
    "map",
    "chat",
    "notifications"
  ];

  const tabs = user?.role === "admin" ? ["admin", ...baseTabs] : baseTabs;

  return (
    <div className="navbar">
      <div>
        <h2>Smart Hospital</h2>
        <p>
          {user?.name} ({user?.role})
        </p>
      </div>

      <div className="nav-tabs">
        {tabs.map((tab) => (
          <button
            key={tab}
            className={activeTab === tab ? "tab-button active" : "tab-button"}
            onClick={() => setActiveTab(tab)}
          >
            {tab}
          </button>
        ))}
      </div>

      <button className="danger-btn small-btn" onClick={onLogout}>
        Logout
      </button>
    </div>
  );
}

export default Navbar;
