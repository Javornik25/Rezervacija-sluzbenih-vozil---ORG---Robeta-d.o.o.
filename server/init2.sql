CREATE DATABASE robeta_rezervacije;

USE robeta_rezervacije;

CREATE TABLE users (
  id INT AUTO_INCREMENT PRIMARY KEY,
  first_name VARCHAR(100),
  last_name VARCHAR(100),
  email VARCHAR(150) UNIQUE,
  role VARCHAR(50)
);

CREATE TABLE cars (
  id INT AUTO_INCREMENT PRIMARY KEY,
  car_key VARCHAR(50) UNIQUE,
  name VARCHAR(100),
  registration VARCHAR(20),
  admin_only BOOLEAN DEFAULT FALSE
);

CREATE TABLE reservations (
  id INT AUTO_INCREMENT PRIMARY KEY,
  car_id INT,
  user_id INT,
  user_email VARCHAR(150),
  user_name VARCHAR(150),
  reserved_by_email VARCHAR(150),
  reserved_by_name VARCHAR(150),
  start_time DATETIME,
  end_time DATETIME,
  FOREIGN KEY (car_id) REFERENCES cars(id),
  FOREIGN KEY (user_id) REFERENCES users(id)
);

INSERT INTO cars (car_key, name, registration, admin_only)
VALUES
('audi-a4', 'Audi A4', 'SG RO-001', TRUE),
('audi-a6', 'Audi A6', 'SG RO-002', TRUE),
('bmw-320d', 'BMW 320d', 'SG RO-003', TRUE),
('bmw-520d', 'BMW 520d', 'SG RO-004', TRUE),
('mercedes-c220', 'Mercedes C220', 'SG RO-005', FALSE),
('skoda-superb', 'Škoda Superb', 'SG RO-006', FALSE),
('skoda-octavia', 'Škoda Octavia', 'SG RO-007', FALSE),
('vw-passat', 'Volkswagen Passat', 'SG RO-008', FALSE),
('toyota-corolla', 'Toyota Corolla', 'SG RO-009', FALSE),
('hyundai-i30', 'Hyundai i30', 'SG RO-010', FALSE);

INSERT INTO reservations (car_id, user_email, user_name, reserved_by_email, reserved_by_name, start_time, end_time)
VALUES
(6, 'stas.javornik@robeta.si', 'Staš Javornik', 'stas.javornik@robeta.si', 'Staš Javornik', '2026-06-03 09:00:00', '2026-06-03 11:00:00'),
(8, 'jan.novak@robeta.si', 'Jan Novak', 'jan.novak@robeta.si', 'Jan Novak', '2026-06-03 12:00:00', '2026-06-03 14:00:00');

SELECT * FROM reservations;
SELECT * FROM cars;