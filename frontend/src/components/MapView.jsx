import { MapContainer, Marker, Popup, TileLayer } from "react-leaflet";
import L from "leaflet";

const markerIcon = new L.Icon({
  iconUrl: "https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon.png",
  shadowUrl: "https://unpkg.com/leaflet@1.9.4/dist/images/marker-shadow.png",
  iconSize: [25, 41],
  iconAnchor: [12, 41]
});

function MapView() {
  const hospitalPosition = [28.6139, 77.209];

  return (
    <MapContainer center={hospitalPosition} zoom={13} className="map-box">
      <TileLayer
        attribution="&copy; OpenStreetMap contributors"
        url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
      />
      <Marker position={hospitalPosition} icon={markerIcon}>
        <Popup>Smart Hospital Main Branch</Popup>
      </Marker>
    </MapContainer>
  );
}

export default MapView;
