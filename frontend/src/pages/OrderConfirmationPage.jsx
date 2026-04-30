import { Navigate, useLocation, useNavigate } from "react-router-dom";

function OrderConfirmationPage() {
  const navigate = useNavigate();
  const location = useLocation();
  const order = location.state?.order;

  if (!order) {
    return <Navigate to="/dashboard" replace />;
  }

  return (
    <div className="confirmation-page">
      <div className="confirmation-card">
        <p className="confirmation-tag">Order Confirmed</p>
        <h1>Your order has been placed successfully</h1>
        <p className="muted-text">
          Cash on Delivery has been selected for this order.
        </p>

        <div className="confirmation-summary">
          <div className="confirmation-row">
            <span>Order ID</span>
            <strong>{order._id}</strong>
          </div>
          <div className="confirmation-row">
            <span>Payment Method</span>
            <strong>{order.paymentMethod}</strong>
          </div>
          <div className="confirmation-row">
            <span>Order Status</span>
            <strong>{order.orderStatus}</strong>
          </div>
          <div className="confirmation-row total">
            <span>Order Total</span>
            <strong>Rs. {order.totalAmount}</strong>
          </div>
        </div>

        <div className="confirmation-items">
          <h3>Items Ordered</h3>
          {order.items.map((item, index) => (
            <div key={`${item.name}-${index}`} className="confirmation-item">
              <span>{item.name}</span>
              <span>
                {item.quantity} x Rs. {item.price}
              </span>
            </div>
          ))}
        </div>

        <button
          className="primary-btn confirmation-btn"
          onClick={() => navigate("/dashboard", { state: { activeTab: "orders" } })}
        >
          Continue Buying
        </button>
      </div>
    </div>
  );
}

export default OrderConfirmationPage;
