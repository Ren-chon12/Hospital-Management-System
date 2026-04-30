# Smart Hospital Management System

Simple full stack MERN project using:

- MongoDB
- Express.js
- React with JSX and Vite
- Node.js
- MVC pattern on the backend

## Features

- User authentication
- Email OTP based verification
- Admin panel with CRUD
- Patient and appointment management
- Search and filter system
- Medical order processing with cart and COD
- Admin-user connectivity through shared dashboards and chat
- File uploads
- Geo-location map integration
- Real-time chat with Socket.io
- Notification system
- Dummy seed data

## Project Structure

```text
backend/
frontend/
README.md
```

## Backend Setup

1. Open `backend/.env.example` and copy it to `.env`
2. Update MongoDB and email values if needed
3. Install dependencies

```bash
cd backend
npm install
npm run seed
npm run dev
```

## Frontend Setup

```bash
cd frontend
npm install
npm run dev
```

## Demo Credentials

- Admin: `admin@hospital.com` / `123456`
- Doctor: `doctor@hospital.com` / `123456`
- Patient: `patient@hospital.com` / `123456`

## Notes

- OTP emails fall back to console logging if SMTP is not configured.
- File upload currently uses local storage in `backend/uploads`.
- COD is implemented for order placement, and payment API can be added later.
- Map uses OpenStreetMap tiles through Leaflet.
- Future scope like telemedicine and wearable integration can be added as extra modules.
