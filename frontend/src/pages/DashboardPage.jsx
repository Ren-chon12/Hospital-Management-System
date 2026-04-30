import { useEffect, useMemo, useState } from "react";
import { useLocation, useNavigate } from "react-router-dom";
import api from "../api/axios";
import { useAuth } from "../context/AuthContext";
import Navbar from "../components/Navbar";
import StatCard from "../components/StatCard";
import SectionCard from "../components/SectionCard";
import NotificationBell from "../components/NotificationBell";
import MapView from "../components/MapView";
import ChatBox from "../components/ChatBox";

const storeItems = [
  { name: "Vitamin Tablets", price: 250 },
  { name: "Digital Thermometer", price: 180 },
  { name: "BP Monitor", price: 1450 }
];

function DashboardPage() {
  const { user, logout } = useAuth();
  const navigate = useNavigate();
  const location = useLocation();
  const [activeTab, setActiveTab] = useState(
    location.state?.activeTab || (user.role === "admin" ? "admin" : "overview")
  );
  const [dashboard, setDashboard] = useState({
    users: 0,
    patients: 0,
    appointments: 0,
    orders: 0
  });
  const [users, setUsers] = useState([]);
  const [patients, setPatients] = useState([]);
  const [appointments, setAppointments] = useState([]);
  const [orders, setOrders] = useState([]);
  const [cart, setCart] = useState([]);
  const [notifications, setNotifications] = useState([]);
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [uploadMessage, setUploadMessage] = useState("");
  const [forms, setForms] = useState({
    patient: { patientName: "", age: "", gender: "Male", bloodGroup: "", disease: "" },
    appointment: {
      patient: "",
      doctor: "",
      appointmentDate: "",
      appointmentTime: "",
      reason: ""
    }
  });

  const doctors = useMemo(
    () => users.filter((item) => item.role === "doctor"),
    [users]
  );
  const patientsUsers = useMemo(
    () => users.filter((item) => item.role === "patient"),
    [users]
  );

  const loadData = async () => {
    const requests = [
      api.get("/appointments", { params: { search, status: statusFilter } }),
      api.get("/orders", { params: { search } }),
      api.get("/notifications"),
      api.get("/orders/cart")
    ];

    if (user.role === "admin") {
      requests.unshift(
        api.get("/admin/dashboard"),
        api.get("/users", { params: { search } }),
        api.get("/patients", { params: { search } })
      );
    } else {
      requests.unshift(
        Promise.resolve({ data: dashboard }),
        api.get("/users/contacts"),
        api.get("/patients", { params: { search } })
      );
    }

    const [
      dashboardRes,
      usersRes,
      patientsRes,
      appointmentsRes,
      ordersRes,
      notificationsRes,
      cartRes
    ] = await Promise.all(requests);

    setDashboard(dashboardRes.data);
    setUsers(usersRes.data);
    setPatients(patientsRes.data);
    setAppointments(appointmentsRes.data);
    setOrders(ordersRes.data);
    setNotifications(notificationsRes.data);
    setCart(cartRes.data);
  };

  useEffect(() => {
    loadData();
  }, [search, statusFilter]);

  useEffect(() => {
    if (location.state?.activeTab) {
      setActiveTab(location.state.activeTab);
      navigate(location.pathname, { replace: true, state: null });
    }
  }, [location.pathname, location.state, navigate]);

  const handleLogout = () => {
    logout();
  };

  const updateForm = (section, key, value) => {
    setForms((prev) => ({
      ...prev,
      [section]: { ...prev[section], [key]: value }
    }));
  };

  // Create patient
  const createPatient = async (e) => {
    e.preventDefault();
    await api.post("/patients", forms.patient);
    setForms((prev) => ({
      ...prev,
      patient: { patientName: "", age: "", gender: "Male", bloodGroup: "", disease: "" }
    }));
    loadData();
  };

  // Create appointment
  const createAppointment = async (e) => {
    e.preventDefault();
    await api.post("/appointments", forms.appointment);
    setForms((prev) => ({
      ...prev,
      appointment: {
        patient: "",
        doctor: "",
        appointmentDate: "",
        appointmentTime: "",
        reason: ""
      }
    }));
    loadData();
  };

  // Delete patient
  const removePatient = async (id) => {
    await api.delete(`/patients/${id}`);
    loadData();
  };

  // Update patient
  const editPatient = async (item) => {
    const disease = window.prompt("Update disease", item.disease || "");
    const age = window.prompt("Update age", item.age || "");

    if (disease === null || age === null) return;

    await api.put(`/patients/${item._id}`, {
      disease,
      age: Number(age)
    });
    loadData();
  };

  // Update appointment status
  const updateAppointmentStatus = async (id, status) => {
    await api.put(`/appointments/${id}`, { status });
    loadData();
  };

  // Delete appointment
  const removeAppointment = async (id) => {
    await api.delete(`/appointments/${id}`);
    loadData();
  };

  // Add item to cart
  const addItemToCart = async (item) => {
    const { data } = await api.post("/orders/cart", { ...item, quantity: 1 });
    setCart(data);
  };

  // Remove one item from cart
  const removeOneFromCart = async (name) => {
    const { data } = await api.put("/orders/cart", { name });
    setCart(data);
  };

  // Place order
  const placeOrder = async () => {
    const { data } = await api.post("/orders/place");
    await loadData();
    navigate("/order-confirmation", { state: { order: data } });
  };

  // Update order status
  const updateOrderStatus = async (id, orderStatus) => {
    await api.put(`/orders/${id}`, { orderStatus });
    loadData();
  };

  // Delete delivered order
  const deleteOrderHistory = async (id) => {
    await api.delete(`/orders/${id}`);
    loadData();
  };

  // Upload file
  const handleUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const formData = new FormData();
    formData.append("file", file);
    const { data } = await api.post("/uploads", formData, {
      headers: { "Content-Type": "multipart/form-data" }
    });
    setUploadMessage(data.filePath);
  };

  // Mark notification read
  const markRead = async (id) => {
    await api.put(`/notifications/${id}/read`);
    loadData();
  };

  const renderOverview = () => (
    <>
      <div className="page-hero">
        <div className="page-hero-copy">
          <span className="page-hero-tag">Hospital Dashboard</span>
          <h1>Manage patients, appointments, orders, and communication in one place.</h1>
          <p>
            Keep hospital operations organized with a simple digital workflow for admins,
            doctors, and patients.
          </p>
        </div>
        <div className="page-hero-panel">
          <div className="hero-chip">Role: {user.role}</div>
          <div className="hero-chip">Search enabled</div>
          <div className="hero-chip">Notifications active</div>
        </div>
      </div>

      <div className="stats-grid">
        <StatCard label="Users" value={dashboard.users || users.length} />
        <StatCard label="Patients" value={dashboard.patients || patients.length} />
        <StatCard label="Appointments" value={dashboard.appointments || appointments.length} />
        <StatCard label="Orders" value={dashboard.orders || orders.length} />
      </div>

      <SectionCard
        title="Quick Search & Filters"
        subtitle="Find records faster by keyword or appointment status."
      >
        <div className="inline-form">
          <input
            type="text"
            placeholder="Search users, patients, doctors, orders"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
          <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
            <option value="">All Appointment Status</option>
            <option value="scheduled">Scheduled</option>
            <option value="completed">Completed</option>
            <option value="cancelled">Cancelled</option>
          </select>
        </div>
      </SectionCard>

      <SectionCard
        title="Appointments"
        subtitle="Track current bookings and update appointment status when needed."
      >
        <div className="list-grid">
          {appointments.map((item) => (
            <div key={item._id} className="list-card">
              <p>
                <strong>Patient:</strong> {item.patient?.name}
              </p>
              <p>
                <strong>Doctor:</strong> {item.doctor?.name}
              </p>
              <p>
                <strong>Date:</strong> {item.appointmentDate} {item.appointmentTime}
              </p>
              <p>
                <strong>Reason:</strong> {item.reason}
              </p>
              <p>
                <strong>Status:</strong> {item.status}
              </p>
              {(user.role === "admin" || user.role === "doctor") && (
                <div className="row-gap">
                  <button
                    className="secondary-btn small-btn"
                    onClick={() => updateAppointmentStatus(item._id, "completed")}
                  >
                    Mark Completed
                  </button>
                  <button
                    className="danger-btn small-btn"
                    onClick={() => updateAppointmentStatus(item._id, "cancelled")}
                  >
                    Cancel
                  </button>
                  {user.role === "admin" && (
                    <button
                      className="danger-btn small-btn"
                      onClick={() => removeAppointment(item._id)}
                    >
                      Delete
                    </button>
                  )}
                </div>
              )}
            </div>
          ))}
        </div>
      </SectionCard>
    </>
  );

  const renderOrders = () => (
    <>
      <SectionCard
        title="Medical Store"
        subtitle="Add medicines and hospital essentials to the cart."
      >
        <div className="list-grid">
          {storeItems.map((item) => (
            <div key={item.name} className="list-card">
              <p>
                <strong>{item.name}</strong>
              </p>
              <p>Price: Rs. {item.price}</p>
              <button className="primary-btn small-btn" onClick={() => addItemToCart(item)}>
                Add to Cart
              </button>
            </div>
          ))}
        </div>
      </SectionCard>
    </>
  );

  const renderCart = () => (
    <>
      <SectionCard
        title="Cart & COD Checkout"
        subtitle="Review selected items before placing the order."
      >
        <div className="list-grid">
          {cart.map((item, index) => (
            <div key={`${item.name}-${index}`} className="list-card">
              <p>{item.name}</p>
              <p>Qty: {item.quantity}</p>
              <p>Total: Rs. {item.price * item.quantity}</p>
              <button
                className="danger-btn small-btn"
                onClick={() => removeOneFromCart(item.name)}
              >
                Remove One
              </button>
            </div>
          ))}
        </div>
        {!cart.length && <p className="muted-text">Your cart is empty right now.</p>}
        <button className="primary-btn" onClick={placeOrder} disabled={!cart.length}>
          Place Order with COD
        </button>
      </SectionCard>

      <SectionCard
        title="Cart Summary"
        subtitle="Quick view of item count and estimated bill."
      >
        <div className="stats-grid">
          <StatCard label="Items in Cart" value={cart.length} />
          <StatCard
            label="Cart Total"
            value={`Rs. ${cart.reduce(
              (sum, item) => sum + item.price * item.quantity,
              0
            )}`}
          />
        </div>
      </SectionCard>
    </>
  );

  const renderOrderHistory = () => (
    <SectionCard
      title="Orders History"
      subtitle="Review placed orders and manage delivered order records."
    >
        <div className="list-grid">
          {orders.map((item) => (
            <div key={item._id} className="list-card">
              <p>
                <strong>User:</strong> {item.user?.name}
              </p>
              <p>
                <strong>Total:</strong> Rs. {item.totalAmount}
              </p>
              <p>
                <strong>Payment:</strong> {item.paymentMethod} ({item.paymentStatus})
              </p>
              <p>
                <strong>Status:</strong> {item.orderStatus}
              </p>
              {user.role === "admin" && (
                <div className="row-gap">
                  {item.orderStatus !== "delivered" && (
                    <button
                      className="secondary-btn small-btn"
                      onClick={() => updateOrderStatus(item._id, "delivered")}
                    >
                      Mark Delivered
                    </button>
                  )}
                  {item.orderStatus === "delivered" && (
                    <button
                      className="danger-btn small-btn"
                      onClick={() => deleteOrderHistory(item._id)}
                    >
                      Remove from History
                    </button>
                  )}
                </div>
              )}
              {user.role !== "admin" && item.orderStatus === "delivered" && (
                <button
                  className="danger-btn small-btn"
                  onClick={() => deleteOrderHistory(item._id)}
                >
                  Remove from History
                </button>
              )}
            </div>
          ))}
        </div>
        {!orders.length && <p className="muted-text">No orders have been placed yet.</p>}
      </SectionCard>
  );

  const renderUpload = () => (
    <SectionCard
      title="File Uploads"
      subtitle="Upload reports or supporting medical documents."
    >
      <input type="file" onChange={handleUpload} />
      {uploadMessage && <p className="info-text">Uploaded file: {uploadMessage}</p>}
      <p className="muted-text">
        This uses local uploads and can be replaced with cloud storage later.
      </p>
    </SectionCard>
  );

  const renderMap = () => (
    <SectionCard
      title="Hospital Location"
      subtitle="Help patients and visitors locate the hospital branch."
    >
      <MapView />
    </SectionCard>
  );

  const renderNotifications = () => (
    <SectionCard
      title="Notifications"
      subtitle="Stay updated with account, appointment, and order activity."
    >
      <NotificationBell notifications={notifications} onRead={markRead} />
    </SectionCard>
  );

  const renderAdmin = () => (
    <>
      <SectionCard title="Create Patient">
        <form className="grid-form" onSubmit={createPatient}>
          <input
            placeholder="Patient Name"
            value={forms.patient.patientName}
            onChange={(e) => updateForm("patient", "patientName", e.target.value)}
            required
          />
          <input
            type="number"
            placeholder="Age"
            value={forms.patient.age}
            onChange={(e) => updateForm("patient", "age", e.target.value)}
            required
          />
          <select
            value={forms.patient.gender}
            onChange={(e) => updateForm("patient", "gender", e.target.value)}
          >
            <option value="Male">Male</option>
            <option value="Female">Female</option>
            <option value="Other">Other</option>
          </select>
          <input
            placeholder="Blood Group"
            value={forms.patient.bloodGroup}
            onChange={(e) => updateForm("patient", "bloodGroup", e.target.value)}
          />
          <input
            placeholder="Disease"
            value={forms.patient.disease}
            onChange={(e) => updateForm("patient", "disease", e.target.value)}
          />
          <button className="primary-btn" type="submit">
            Add Patient
          </button>
        </form>
      </SectionCard>

      <SectionCard
        title="Create Appointment"
        subtitle="Schedule a doctor visit by selecting a patient, doctor, and time."
      >
        <form className="grid-form" onSubmit={createAppointment}>
          <select
            value={forms.appointment.patient}
            onChange={(e) => updateForm("appointment", "patient", e.target.value)}
            required
          >
            <option value="">Select Patient</option>
            {patientsUsers.map((item) => (
              <option key={item._id} value={item._id}>
                {item.name}
              </option>
            ))}
          </select>
          <select
            value={forms.appointment.doctor}
            onChange={(e) => updateForm("appointment", "doctor", e.target.value)}
            required
          >
            <option value="">Select Doctor</option>
            {doctors.map((item) => (
              <option key={item._id} value={item._id}>
                {item.name}
              </option>
            ))}
          </select>
          <input
            type="date"
            value={forms.appointment.appointmentDate}
            onChange={(e) => updateForm("appointment", "appointmentDate", e.target.value)}
            required
          />
          <input
            placeholder="Time"
            value={forms.appointment.appointmentTime}
            onChange={(e) => updateForm("appointment", "appointmentTime", e.target.value)}
            required
          />
          <input
            placeholder="Reason"
            value={forms.appointment.reason}
            onChange={(e) => updateForm("appointment", "reason", e.target.value)}
            required
          />
          <button className="primary-btn" type="submit">
            Book Appointment
          </button>
        </form>
      </SectionCard>

      <SectionCard
        title="Users"
        subtitle="Overview of all registered users in the system."
      >
        <div className="list-grid">
          {users.map((item) => (
            <div key={item._id} className="list-card">
              <p>
                <strong>{item.name}</strong>
              </p>
              <p>{item.email}</p>
              <p>{item.role}</p>
            </div>
          ))}
        </div>
      </SectionCard>

      <SectionCard
        title="Patients CRUD"
        subtitle="Edit or remove patient records from the admin panel."
      >
        <div className="list-grid">
          {patients.map((item) => (
            <div key={item._id} className="list-card">
              <p>
                <strong>{item.user?.name}</strong>
              </p>
              <p>Age: {item.age}</p>
              <p>Disease: {item.disease}</p>
              <button
                className="secondary-btn small-btn"
                onClick={() => editPatient(item)}
              >
                Edit
              </button>
              <button className="danger-btn small-btn" onClick={() => removePatient(item._id)}>
                Delete
              </button>
            </div>
          ))}
        </div>
      </SectionCard>
    </>
  );

  return (
    <div className="dashboard-page">
      <Navbar
        user={user}
        onLogout={handleLogout}
        activeTab={activeTab}
        setActiveTab={setActiveTab}
      />

      <div className="content-wrap">
        {activeTab === "admin" && user.role === "admin" && renderAdmin()}
        {activeTab === "overview" && renderOverview()}
        {activeTab === "appointments" && renderOverview()}
        {activeTab === "orders" && renderOrders()}
        {activeTab === "cart" && renderCart()}
        {activeTab === "orders" && renderOrderHistory()}
        {activeTab === "upload" && renderUpload()}
        {activeTab === "map" && renderMap()}
        {activeTab === "chat" && (
          <SectionCard title="Real-Time Chat">
            <ChatBox user={user} />
          </SectionCard>
        )}
        {activeTab === "notifications" && renderNotifications()}
      </div>
    </div>
  );
}

export default DashboardPage;
