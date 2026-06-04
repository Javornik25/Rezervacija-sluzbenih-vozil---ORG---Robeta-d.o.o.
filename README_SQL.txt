ROBETA - SQL POPRAVEK

Urejeno:
- dodan / urejen server/server.js
- dodani paketi express, mysql2 in cors v package.json
- vozila se v App.jsx nalagajo iz MySQL prek http://localhost:3001/cars
- SQL polja se pretvorijo:
  car_key -> id
  name -> znamka
  registration -> registracija
  admin_only -> adminOnly
- dodan server/init.sql za ustvarjanje baze in vnos vozil

Zagon:
1. V server/server.js preveri geslo:
   password: "TVOJE_GESLO"

2. Terminal 1:
   cd robeta-final-m365
   node server/server.js

3. Terminal 2:
   cd robeta-final-m365
   npm run dev

Test:
- http://localhost:3001/cars mora vrniti seznam vozil iz MySQL.
