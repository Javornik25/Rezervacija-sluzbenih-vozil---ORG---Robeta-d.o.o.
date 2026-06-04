const express = require("express");
const mysql = require("mysql2");
const cors = require("cors");

const app = express();

app.use(cors());
app.use(express.json());

const db = mysql.createConnection({
  host: "localhost",
  user: "root",
  password: "ROBETA_2026!", // tukaj vpiši svoje MySQL root geslo, če še ni pravilno
  database: "robeta_rezervacije",
});

db.connect((err) => {
  if (err) {
    console.error("Napaka pri povezavi z MySQL:", err);
    return;
  }

  console.log("MySQL povezan!");
});

app.get("/cars", (req, res) => {
  db.query(
    "SELECT id, car_key, name, registration, admin_only FROM cars ORDER BY id",
    (err, results) => {
      if (err) {
        console.error("Napaka pri branju vozil:", err);
        res.status(500).json({ error: "Napaka pri branju vozil." });
        return;
      }

      res.json(results);
    }
  );
});

app.get("/reservations", (req, res) => {
  db.query(
    `SELECT 
      r.id,
      c.car_key AS carId,
      c.name AS carName,
      c.registration AS registracija,
      r.user_email AS userEmail,
      r.user_name AS userName,
      r.reserved_by_email AS reservedByEmail,
      r.reserved_by_name AS reservedByName,
      r.start_time AS start,
      r.end_time AS end
    FROM reservations r
    JOIN cars c ON r.car_id = c.id
    ORDER BY r.start_time`,
    (err, results) => {
      if (err) {
        console.error("Napaka pri branju rezervacij:", err);
        res.status(500).json({ error: "Napaka pri branju rezervacij." });
        return;
      }

      res.json(results);
    }
  );
});

app.post("/reservations", (req, res) => {
  const {
    carId,
    userEmail,
    userName,
    reservedByEmail,
    reservedByName,
    start,
    end,
  } = req.body;

  if (!carId || !userEmail || !userName || !start || !end) {
    res.status(400).json({ error: "Manjkajo podatki za rezervacijo." });
    return;
  }

  db.query("SELECT id FROM cars WHERE car_key = ?", [carId], (carErr, carRows) => {
    if (carErr) {
      console.error(carErr);
      res.status(500).json({ error: "Napaka pri iskanju vozila." });
      return;
    }

    if (carRows.length === 0) {
      res.status(404).json({ error: "Vozilo ne obstaja v bazi." });
      return;
    }

    const carDbId = carRows[0].id;

    const overlapSql = `
      SELECT id FROM reservations
      WHERE car_id = ?
      AND start_time < ?
      AND end_time > ?
      LIMIT 1
    `;

    db.query(overlapSql, [carDbId, end, start], (overlapErr, overlapRows) => {
      if (overlapErr) {
        console.error(overlapErr);
        res.status(500).json({ error: "Napaka pri preverjanju termina." });
        return;
      }

      if (overlapRows.length > 0) {
        res.status(409).json({ error: "To vozilo je v tem terminu že rezervirano." });
        return;
      }

      const insertSql = `
        INSERT INTO reservations
        (car_id, user_email, user_name, reserved_by_email, reserved_by_name, start_time, end_time)
        VALUES (?, ?, ?, ?, ?, ?, ?)
      `;

      db.query(
        insertSql,
        [
          carDbId,
          userEmail,
          userName,
          reservedByEmail || userEmail,
          reservedByName || userName,
          start,
          end,
        ],
        (insertErr, result) => {
          if (insertErr) {
            console.error(insertErr);
            res.status(500).json({ error: "Napaka pri shranjevanju rezervacije." });
            return;
          }

          res.status(201).json({ id: result.insertId });
        }
      );
    });
  });
});

app.delete("/reservations/:id", (req, res) => {
  db.query("DELETE FROM reservations WHERE id = ?", [req.params.id], (err) => {
    if (err) {
      console.error("Napaka pri brisanju rezervacije:", err);
      res.status(500).json({ error: "Napaka pri brisanju rezervacije." });
      return;
    }

    res.json({ success: true });
  });
});

app.listen(3001, () => {
  console.log("Server teče na portu 3001");
});
