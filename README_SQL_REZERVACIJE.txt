ROBETA - SQL REZERVACIJE

Dodano:
- ob zagonu aplikacija prebere rezervacije iz http://localhost:3001/reservations
- ob kliku Rezerviraj vozilo se rezervacija shrani v MySQL prek POST /reservations
- ob brisanju se rezervacija izbriše iz MySQL prek DELETE /reservations/:id

Zagon:
Terminal 1:
node server/server.js

Terminal 2:
npm run dev

Preverjanje v MySQL Workbench:
SELECT * FROM reservations;
