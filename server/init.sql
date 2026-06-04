CREATE DATABASE IF NOT EXISTS robeta_rezervacije;
USE robeta_rezervacije;

CREATE TABLE IF NOT EXISTS users (
  id INT AUTO_INCREMENT PRIMARY KEY,
  first_name VARCHAR(100),
  last_name VARCHAR(100),
  email VARCHAR(150) UNIQUE,
  role VARCHAR(50)
);

CREATE TABLE IF NOT EXISTS cars (
  id INT AUTO_INCREMENT PRIMARY KEY,
  car_key VARCHAR(50) UNIQUE,
  name VARCHAR(100),
  registration VARCHAR(30),
  admin_only BOOLEAN DEFAULT FALSE
);

CREATE TABLE IF NOT EXISTS reservations (
  id INT AUTO_INCREMENT PRIMARY KEY,
  car_id INT,
  user_id INT NULL,
  user_email VARCHAR(150),
  user_name VARCHAR(150),
  reserved_by_email VARCHAR(150),
  reserved_by_name VARCHAR(150),
  start_time DATETIME,
  end_time DATETIME,
  FOREIGN KEY (car_id) REFERENCES cars(id),
  FOREIGN KEY (user_id) REFERENCES users(id)
);

INSERT IGNORE INTO cars (car_key, name, registration, admin_only)
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
