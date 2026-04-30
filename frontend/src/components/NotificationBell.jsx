function NotificationBell({ notifications, onRead }) {
  return (
    <div className="list-grid">
      {notifications.map((item) => (
        <div key={item._id} className={item.read ? "notice read" : "notice"}>
          <div>
            <strong>{item.title}</strong>
            <p>{item.message}</p>
          </div>
          {!item.read && (
            <button className="primary-btn small-btn" onClick={() => onRead(item._id)}>
              Mark Read
            </button>
          )}
        </div>
      ))}
      {!notifications.length && <p>No notifications yet.</p>}
    </div>
  );
}

export default NotificationBell;
