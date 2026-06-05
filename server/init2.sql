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
('ford-tourneo-15', 'Ford Tourneo 1.5', 'SG JF-219', FALSE),
('ford-transit-keosn', 'Ford Transit KEOSN', 'SG JF-067', FALSE),
('mercedes-sprinter-delavnica', 'Mercedes Sprinter Delavnica', 'SG SU-599', FALSE),
('citroen-berlingo', 'Citroen Berlingo', 'SG IF-364', FALSE),
('tovorna-prikolica', 'Tovorna Prikolica', 'SG LC-718', FALSE),
('ford-transit-veliki-siv', 'Ford Transit Veliki Siv', 'SG AC-955', FALSE),
('ford-transit-custom-sedezi', 'Ford Transit Custom Sedeži', 'SG FH-961', FALSE);


SELECT * FROM reservations;
SELECT * FROM cars;