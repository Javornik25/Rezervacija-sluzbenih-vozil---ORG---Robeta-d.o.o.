import { useEffect, useMemo, useState } from "react";
import FullCalendar from "@fullcalendar/react";
import dayGridPlugin from "@fullcalendar/daygrid";
import timeGridPlugin from "@fullcalendar/timegrid";
import interactionPlugin from "@fullcalendar/interaction";
import logo from "./assets/robeta-logo.png";
import { PublicClientApplication } from "@azure/msal-browser";

const RESERVATIONS_KEY = "robeta_rezervacije_monday_logout";
const USERS_KEY = "robeta_uporabniki_monday_logout";
const SESSION_KEY = "robeta_prijavljen_uporabnik_monday_logout";
const LOGIN_MESSAGE_KEY = "robeta_login_message_monday_logout";

const msalConfig = {
  auth: {
    clientId: "89da1f76-52a6-4278-9349-92ca4cde4a69",
    authority: "https://login.microsoftonline.com/b910a59c-fb77-4328-9e53-593dfdba7617",
    redirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false,
  },
};

const msalInstance = new PublicClientApplication(msalConfig);
let msalReady = msalInstance.initialize();

const microsoftLoginRequest = {
  scopes: ["openid", "profile", "email"],
  prompt: "select_account",
};

msalReady.then(async () => {
  try {
    const result = await msalInstance.handleRedirectPromise();

    if (result?.account) {
      const email = (result.account.username || "").toLowerCase();

      if (!email.endsWith("@robeta.si")) return;

      const savedUsers = JSON.parse(localStorage.getItem("robeta_uporabniki_monday_logout") || "[]");
      let user = savedUsers.find((u) => u.email === email);

      if (!user) {
        const parts = (result.account.name || "").split(" ");
        user = {
          id: crypto.randomUUID(),
          firstName: parts[0] || "Uporabnik",
          lastName: parts.slice(1).join(" ") || "",
          email,
          password: "",
          isAdmin: [
          "stas.javornik@robeta.si",
          "matjaz@robeta.si"
          ].includes(email),
          loginType: "microsoft365",
        };
      }

      localStorage.setItem(
        "robeta_prijavljen_uporabnik_monday_logout",
        JSON.stringify(user)
      );

      localStorage.setItem(
        LOGIN_MESSAGE_KEY,
        `Dobrodošli, ${user.firstName} ${user.lastName}`
      );

      window.location.reload();
    }
  } catch (e) {
    console.error(e);
  }
});
function getCarIconById(carId, carsList = []) {
  const car = carsList.find((item) => item.id === carId);
  return car?.adminOnly ? "🔒🚗" : "🚗";
}

function getCarIcon(carName) {
  const lower = carName.toLowerCase();
  return lower.includes("audi") || lower.includes("bmw") ? "🔒🚗" : "🚗";
}

function getCarDisplayName(carName) {
  return `${getCarIcon(carName)} ${carName}`;
}

function getCarDisplayNameById(carId, carName, carsList = []) {
  return `${getCarIconById(carId, carsList)} ${carName}`;
}

function getEventTitle(reservation, carsList = []) {
  return `${getCarDisplayNameById(
    reservation.carId,
    reservation.carName,
    carsList
  )} - ${reservation.userName}`;
}

function getPlainCarName(carName) {
  return carName.replace(/^🔒🚗\s*/, "").replace(/^🚗\s*/, "");
}

function getTodayDate() {
  return new Date().toISOString().slice(0, 10);
}

function formatDateTime(value) {
  return new Date(value).toLocaleString("sl-SI", {
    day: "numeric",
    month: "numeric",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function getYears() {
  const currentYear = new Date().getFullYear();
  return [currentYear - 1, currentYear, currentYear + 1, currentYear + 2];
}

function getMonthOptions() {
  return [
    { value: "01", label: "Januar" },
    { value: "02", label: "Februar" },
    { value: "03", label: "Marec" },
    { value: "04", label: "April" },
    { value: "05", label: "Maj" },
    { value: "06", label: "Junij" },
    { value: "07", label: "Julij" },
    { value: "08", label: "Avgust" },
    { value: "09", label: "September" },
    { value: "10", label: "Oktober" },
    { value: "11", label: "November" },
    { value: "12", label: "December" },
  ];
}

function getDaysForDate(dateValue) {
  const [year, month] = dateValue.split("-").map(Number);
  const numberOfDays = new Date(year, month, 0).getDate();
  return Array.from({ length: numberOfDays }, (_, index) => String(index + 1).padStart(2, "0"));
}

function updateDatePart(dateValue, part, value) {
  let [year, month, day] = dateValue.split("-");
  if (part === "year") year = value;
  if (part === "month") month = value;
  if (part === "day") day = value;

  const maxDay = new Date(Number(year), Number(month), 0).getDate();
  if (Number(day) > maxDay) day = String(maxDay).padStart(2, "0");

  return `${year}-${month}-${day}`;
}

function updateTimePart(timeValue, part, value) {
  let [hour, minute] = timeValue.split(":");
  if (part === "hour") hour = value;
  if (part === "minute") minute = value;
  return `${hour}:${minute}`;
}

const hours = Array.from({ length: 17 }, (_, index) => String(index + 6).padStart(2, "0"));
const minutes = ["00", "15", "30", "45"];

function formatDisplayDate(dateValue) {
  return new Date(`${dateValue}T12:00`).toLocaleDateString("sl-SI", {
    day: "2-digit",
    month: "long",
    year: "numeric",
  });
}

function getCalendarDays(monthDate) {
  const year = monthDate.getFullYear();
  const month = monthDate.getMonth();

  const firstDay = new Date(year, month, 1);
  const mondayIndex = (firstDay.getDay() + 6) % 7;
  const start = new Date(year, month, 1 - mondayIndex);

  return Array.from({ length: 42 }, (_, index) => {
    const current = new Date(start);
    current.setDate(start.getDate() + index);

    const yyyy = current.getFullYear();
    const mm = String(current.getMonth() + 1).padStart(2, "0");
    const dd = String(current.getDate()).padStart(2, "0");

    return {
      value: `${yyyy}-${mm}-${dd}`,
      day: current.getDate(),
      currentMonth: current.getMonth() === month,
    };
  });
}

function changeCalendarMonth(dateValue, amount) {
  let date;

  if (dateValue instanceof Date) {
    date = new Date(dateValue.getFullYear(), dateValue.getMonth(), 1, 12);
  } else {
    date = new Date(`${dateValue}T12:00`);
  }

  if (Number.isNaN(date.getTime())) {
    date = new Date();
  }

  date.setMonth(date.getMonth() + amount);
  return new Date(date.getFullYear(), date.getMonth(), 1, 12);
}



function loadFromStorage(key, fallback) {
  try {
    const saved = localStorage.getItem(key);
    return saved ? JSON.parse(saved) : fallback;
  } catch {
    return fallback;
  }
}

function saveToStorage(key, value) {
  localStorage.setItem(key, JSON.stringify(value));
}

function createDefaultUsers() {
  return [
    {
      id: crypto.randomUUID(),
      firstName: "Staš",
      lastName: "Javornik",
      email: "stas.javornik@robeta.si",
      password: "admin123",
      isAdmin: true,
    },
    {
      id: crypto.randomUUID(),
      firstName: "Jan",
      lastName: "Novak",
      email: "jan.novak@robeta.si",
      password: "test123",
      isAdmin: false,
    },
  ];
}

function getFullName(user) {
  return `${user.firstName} ${user.lastName}`;
}

function isRobetaEmail(email) {
  return String(email || "").toLowerCase().endsWith("@robeta.si");
}

function getNamesFromEmail(email) {
  const beforeAt = String(email || "").split("@")[0] || "";
  const parts = beforeAt.split(".").filter(Boolean);

  const firstName = parts[0]
    ? parts[0].charAt(0).toUpperCase() + parts[0].slice(1)
    : "Uporabnik";

  const lastName = parts[1]
    ? parts[1].charAt(0).toUpperCase() + parts[1].slice(1)
    : "Robeta";

  return { firstName, lastName };
}

function getRoleLabel(user) {
  if (!user) return "";
  if (user.isAdmin) return "Administrator";
  if (isRobetaEmail(user.email)) return "Pisarna";
  return "uporabnik";
}


function PickerBox({ id, label, value, options, openPicker, setOpenPicker, onChange, wide = false }) {
  const isOpen = openPicker === id;

  return (
    <div className={wide ? "picker-box wide" : "picker-box"}>
      <button
        type="button"
        className={isOpen ? "picker-display active" : "picker-display"}
        onClick={() => setOpenPicker(isOpen ? null : id)}
      >
        <span>{label}</span>
        <strong>{value}</strong>
        <em>⌄</em>
      </button>

      {isOpen && (
        <div className="picker-menu">
          {options.map((option) => (
            <button
              type="button"
              key={option.value}
              className={option.value === value ? "picker-option selected" : "picker-option"}
              onClick={() => {
                onChange(option.value);
                setOpenPicker(null);
              }}
            >
              {option.label}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}


function CalendarPicker({ id, label, value, openPicker, setOpenPicker, onChange }) {
  const isOpen = openPicker === id;
  const [viewMonth, setViewMonth] = useState(() => new Date(`${value}T12:00`));

  useEffect(() => {
    if (isOpen) {
      setViewMonth(new Date(`${value}T12:00`));
    }
  }, [isOpen, value]);

  const calendarDays = getCalendarDays(viewMonth);
  const monthTitle = viewMonth.toLocaleDateString("sl-SI", {
    month: "long",
    year: "numeric",
  });

  return (
    <div className="calendar-picker">
      <button
        type="button"
        className={isOpen ? "calendar-display active" : "calendar-display"}
        onClick={() => setOpenPicker(isOpen ? null : id)}
      >
        <span>{label}</span>
        <strong>{formatDisplayDate(value)}</strong>
        <em>📅</em>
      </button>

      {isOpen && (
        <div className="mini-calendar">
          <div className="mini-calendar-header">
            <button type="button" onClick={() => setViewMonth(changeCalendarMonth(viewMonth, -1))}>
              ‹
            </button>
            <strong>{monthTitle}</strong>
            <button type="button" onClick={() => setViewMonth(changeCalendarMonth(viewMonth, 1))}>
              ›
            </button>
          </div>

          <div className="mini-calendar-weekdays">
            <span>Pon</span>
            <span>Tor</span>
            <span>Sre</span>
            <span>Čet</span>
            <span>Pet</span>
            <span>Sob</span>
            <span>Ned</span>
          </div>

          <div className="mini-calendar-grid">
            {calendarDays.map((day) => (
              <button
                type="button"
                key={day.value}
                className={
                  day.value === value
                    ? "mini-day selected"
                    : day.value === getTodayDate()
                      ? "mini-day today"
                      : day.currentMonth
                        ? "mini-day"
                        : "mini-day muted"
                }
                onClick={() => {
                  onChange(day.value);
                  setOpenPicker(null);
                }}
              >
                {day.day}
              </button>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

function DateTimePickerPanel({
  title,
  dateValue,
  timeValue,
  onDateChange,
  onTimeChange,
  hours,
  minutes,
  openPicker,
  setOpenPicker,
  prefix,
}) {
  return (
    <div className="date-time-card">
      <div className="date-time-title">{title}</div>

      <CalendarPicker
        id={`${prefix}-calendar`}
        label="Datum"
        value={dateValue}
        openPicker={openPicker}
        setOpenPicker={setOpenPicker}
        onChange={onDateChange}
      />

      <div className="fancy-time-row only-time">
        <PickerBox
          id={`${prefix}-hour`}
          label="Ura"
          value={timeValue.slice(0, 2)}
          options={hours.map((hour) => ({ value: hour, label: hour }))}
          openPicker={openPicker}
          setOpenPicker={setOpenPicker}
          onChange={(value) => onTimeChange(updateTimePart(timeValue, "hour", value))}
        />

        <span className="time-divider">:</span>

        <PickerBox
          id={`${prefix}-minute`}
          label="Minute"
          value={timeValue.slice(3, 5)}
          options={minutes.map((minute) => ({ value: minute, label: minute }))}
          openPicker={openPicker}
          setOpenPicker={setOpenPicker}
          onChange={(value) => onTimeChange(updateTimePart(timeValue, "minute", value))}
        />
      </div>
    </div>
  );
}



function getFilterLabel(filterCarId, cars) {
  if (filterCarId === "all") return "Vsa vozila";
  if (filterCarId === "mine") return "Moje rezervacije";

  const car = cars.find((item) => item.id === filterCarId);
  return car ? getCarDisplayNameById(car.id, car.znamka, cars) : "Filter";
}

function getFilterSubLabel(filterCarId, cars) {
  if (filterCarId === "all") return "Celoten vozni park";
  if (filterCarId === "mine") return "Samo moje rezervacije";

  const car = cars.find((item) => item.id === filterCarId);
  return car ? car.registracija : "";
}

function App() {
  const [users, setUsers] = useState(() => {
    const saved = loadFromStorage(USERS_KEY, []);
    return saved.length > 0 ? saved : createDefaultUsers();
  });
useEffect(() => {
    fetch("http://localhost:3001/reservations")
      .then((res) => res.json())
      .then((data) => {
        const mappedReservations = data.map((reservation) => ({
          ...reservation,
          id: String(reservation.id),
          start: reservation.start,
          end: reservation.end,
        }));

        setReservations(mappedReservations);
      })
      .catch((err) => console.error("Napaka pri nalaganju rezervacij:", err));
  }, []);

  const [cars, setCars] = useState([]);

  useEffect(() => {
    fetch("http://localhost:3001/cars")
      .then((res) => res.json())
      .then((data) => {
        const mappedCars = data.map((car) => ({
          id: car.car_key || String(car.id),
          dbId: car.id,
          znamka: car.name,
          registracija: car.registration,
          adminOnly: Boolean(car.admin_only),
        }));

        setCars(mappedCars);
      })
      .catch((err) => console.error("Napaka pri nalaganju vozil:", err));
  }, []);

  const [user, setUser] = useState(() => loadFromStorage(SESSION_KEY, null));
  const [authMode, setAuthMode] = useState("login");
  const [isTransitioning, setIsTransitioning] = useState(false);
  const [loginMessage, setLoginMessage] = useState(() => localStorage.getItem(LOGIN_MESSAGE_KEY) || "");
  const [selectedReservation, setSelectedReservation] = useState(null);
  const [appMessage, setAppMessage] = useState(null);
  const [goToRegisterAfterMessage, setGoToRegisterAfterMessage] = useState(false);
  const [deleteTarget, setDeleteTarget] = useState(null);

  const [microsoftEmail, setMicrosoftEmail] = useState("stas.javornik@robeta.si");
  const [loginEmail, setLoginEmail] = useState("stas.javornik@robeta.si");
  const [loginPassword, setLoginPassword] = useState("admin123");

  const [registerData, setRegisterData] = useState({
    firstName: "",
    lastName: "",
    email: "",
    password: "",
    confirmPassword: "",
  });

  const [selectedCarId, setSelectedCarId] = useState("mercedes-c220");
  const [filterCarId, setFilterCarId] = useState("all");
  const [startDate, setStartDate] = useState(getTodayDate());
  const [startTime, setStartTime] = useState("08:00");
  const [endDate, setEndDate] = useState(getTodayDate());
  const [endTime, setEndTime] = useState("10:00");
  const [isCalendarBig, setIsCalendarBig] = useState(false);
  const [reservationForMode, setReservationForMode] = useState("self");
  const [reservationForName, setReservationForName] = useState("");
  const [isCarDropdownOpen, setIsCarDropdownOpen] = useState(false);
  const [openPicker, setOpenPicker] = useState(null);
  const [isFilterDropdownOpen, setIsFilterDropdownOpen] = useState(false);

  const [reservations, setReservations] = useState(() => {
    const saved = loadFromStorage(RESERVATIONS_KEY, []);
    if (saved.length > 0) return saved;

    return [
      {
        id: crypto.randomUUID(),
        carId: "skoda-superb",
        carName: "Škoda Superb",
        registracija: "SG RO-006",
        userEmail: "stas.javornik@robeta.si",
        userName: "Staš Javornik",
        reservedByEmail: "stas.javornik@robeta.si",
        reservedByName: "Staš Javornik",
        start: `${getTodayDate()}T09:00`,
        end: `${getTodayDate()}T11:00`,
      },
      {
        id: crypto.randomUUID(),
        carId: "vw-passat",
        carName: "Volkswagen Passat",
        registracija: "SG RO-008",
        userEmail: "jan.novak@robeta.si",
        userName: "Jan Novak",
        reservedByEmail: "jan.novak@robeta.si",
        reservedByName: "Jan Novak",
        start: `${getTodayDate()}T12:00`,
        end: `${getTodayDate()}T14:00`,
      },
    ];
  });

  useEffect(() => saveToStorage(USERS_KEY, users), [users]);
  useEffect(() => saveToStorage(RESERVATIONS_KEY, reservations), [reservations]);

  useEffect(() => {
    if (user) saveToStorage(SESSION_KEY, user);
    else localStorage.removeItem(SESSION_KEY);
  }, [user]);

  const availableCars = useMemo(() => {
    return cars;
  }, [cars]);

  useEffect(() => {
    if (!user?.isAdmin) {
      const selected = cars.find((car) => car.id === selectedCarId);
      if (selected?.adminOnly) {
        const firstUnlocked = cars.find((car) => !car.adminOnly);
        setSelectedCarId(firstUnlocked?.id || "");
      }
    }
  }, [selectedCarId, user]);

  const selectedCar = availableCars.find((car) => car.id === selectedCarId) || availableCars[0] || null;

  const dashboardCars = user?.isAdmin ? cars : cars.filter((car) => !car.adminOnly);
  const dashboardCarIds = new Set(dashboardCars.map((car) => car.id));
  const dashboardReservations = user?.isAdmin
    ? reservations
    : reservations.filter((r) => dashboardCarIds.has(r.carId));

  const yearOptions = getYears();
  const monthOptions = getMonthOptions();
  const startDays = getDaysForDate(startDate);
  const endDays = getDaysForDate(endDate);

  const filteredReservations = useMemo(() => {
    let visibleReservations = reservations;

    if (filterCarId === "mine") return visibleReservations.filter((r) => r.userEmail === user?.email);
    if (filterCarId !== "all") return visibleReservations.filter((r) => r.carId === filterCarId);
    return visibleReservations;
  }, [reservations, filterCarId, user, availableCars]);

  const calendarEvents = useMemo(() => {
    return filteredReservations.map((r) => ({
      id: r.id,
      title: getEventTitle(r, cars),
      start: r.start,
      end: r.end,
      className: r.userEmail === user?.email ? "my-event" : "other-event",
    }));
  }, [filteredReservations, user]);

  const todaysReservations = dashboardReservations.filter(
    (r) => r.start.slice(0, 10) <= getTodayDate() && r.end.slice(0, 10) >= getTodayDate()
  );
  const busyCarsToday = new Set(todaysReservations.map((r) => r.carId)).size;
  const freeCarsToday = dashboardCars.length - busyCarsToday;

  function finishLogin(nextUser) {
    const message =
      localStorage.getItem(LOGIN_MESSAGE_KEY) ||
      `Dobrodošli, ${nextUser.firstName} ${nextUser.lastName}`;

    setLoginMessage(message);
    setIsTransitioning(true);

    setTimeout(() => {
      setUser(nextUser);
      setIsTransitioning(false);
      localStorage.removeItem(LOGIN_MESSAGE_KEY);
      setLoginMessage("");
    }, 1800);
  }

  function showMessage(title, text, type = "info") {
    setAppMessage({ title, text, type });
  }

  function requestDeleteReservation(reservation) {
    if (!reservation) return;

    if (!canDeleteReservation(reservation)) {
      showMessage("Ni dovoljeno", "Izbrišeš lahko samo svoje rezervacije.", "warning");
      return;
    }

    setSelectedReservation(null);
    setDeleteTarget(reservation);
  }

function confirmDeleteReservation() {
  if (!deleteTarget) return;

  fetch(`http://localhost:3001/reservations/${deleteTarget.id}`, {
    method: "DELETE",
  })
    .then(async (res) => {
      if (!res.ok) {
        const data = await res.json();

        showMessage(
          "Brisanje ni uspelo",
          data.error || "Rezervacije ni bilo mogoče izbrisati iz SQL baze.",
          "warning"
        );

        return;
      }

      setReservations(reservations.filter((r) => r.id !== deleteTarget.id));
      setSelectedReservation(null);
      setDeleteTarget(null);

      showMessage(
        "Rezervacija izbrisana",
        "Rezervacija je bila odstranjena iz koledarja.",
        "success"
      );
    })
    .catch((err) => {
      console.error(err);

      showMessage(
        "Napaka povezave",
        "Rezervacije trenutno ni bilo mogoče izbrisati iz SQL baze.",
        "warning"
      );
    });
}

  function handleRegisterChange(field, value) {
    setRegisterData((old) => ({ ...old, [field]: value }));
  }

  function register(e) {
    e.preventDefault();

    const firstName = registerData.firstName.trim();
    const lastName = registerData.lastName.trim();
    const email = registerData.email.trim().toLowerCase();
    const password = registerData.password;
    const confirmPassword = registerData.confirmPassword;

    if (!firstName || !lastName || !email || !password || !confirmPassword) {
      showMessage("Manjkajoči podatki", "Izpolni vsa polja.", "warning");
      return;
    }

    if (!email.includes("@") || !email.includes(".")) {
      showMessage("Neveljaven e-mail", "Vpiši veljaven e-mail naslov.", "warning");
      return;
    }

    if (password.length < 6) {
      showMessage("Prekratko geslo", "Geslo mora imeti vsaj 6 znakov.", "warning");
      return;
    }

    if (password !== confirmPassword) {
      showMessage("Gesli se ne ujemata", "Ponovno preveri geslo in potrditev gesla.", "warning");
      return;
    }

    if (users.some((u) => u.email === email)) {
      showMessage("E-mail že obstaja", "Ta e-mail je že registriran. Uporabnik naj se samo prijavi.", "warning");
      setAuthMode("login");
      setLoginEmail(email);
      setLoginPassword("");
      return;
    }

    const newUser = {
      id: crypto.randomUUID(),
      firstName,
      lastName,
      email,
      password,
      isAdmin: [
    "stas.javornik@robeta.si",
    "matjaz@robeta.si",
    ].includes(email),
      loginType: isRobetaEmail(email) ? "microsoft365" : "local",
    };

    setUsers([...users, newUser]);
    setRegisterData({ firstName: "", lastName: "", email: "", password: "", confirmPassword: "" });
    finishLogin(newUser);
  }

  async function loginMicrosoft(e) {
    e.preventDefault();

    try {
      await msalReady;

      await msalInstance.loginRedirect(microsoftLoginRequest); return;
      const account = result.account;

      if (!account) {
        showMessage("Microsoft 365 prijava", "Microsoft prijava ni vrnila uporabniškega računa.", "warning");
        return;
      }

      const email = (account.username || account.idTokenClaims?.preferred_username || "").toLowerCase();

      if (!isRobetaEmail(email)) {
        showMessage(
          "Dostop zavrnjen",
          "Prijava je dovoljena samo z Robeta Microsoft 365 računom.",
          "warning"
        );
        return;
      }

      const existingUser = users.find((u) => u.email === email);

      if (existingUser) {
        localStorage.setItem(
          LOGIN_MESSAGE_KEY,
          `Dobrodošli, ${existingUser.firstName} ${existingUser.lastName}`
        );

        finishLogin({
          ...existingUser,
          loginType: "microsoft365",
        });

        return;
      }

      const displayName = account.name || "";
      const nameParts = displayName.split(" ").filter(Boolean);
      const namesFromEmail = getNamesFromEmail(email);

      const newUser = {
        id: crypto.randomUUID(),
        firstName: nameParts[0] || namesFromEmail.firstName,
        lastName: nameParts.slice(1).join(" ") || namesFromEmail.lastName,
        email,
        password: "",
        isAdmin: [
        "stas.javornik@robeta.si",
        "matjaz@robeta.si",
        ].includes(email),
        loginType: "microsoft365",
      };

      setUsers([...users, newUser]);
      localStorage.setItem(
        LOGIN_MESSAGE_KEY,
        `Dobrodošli, ${newUser.firstName} ${newUser.lastName}`
      );

      finishLogin(newUser);
    } catch (error) {
      console.error(error);
      showMessage(
        "Microsoft 365 prijava ni uspela",
        "Prijava je bila preklicana ali pa aplikacija v Microsoft Entra še nima pravilno nastavljenega Redirect URI-ja.",
        "warning"
      );
    }
  }

  function login(e) {
    e.preventDefault();

    const email = loginEmail.trim().toLowerCase();
    const foundUserByEmail = users.find((u) => u.email === email);

    if (!foundUserByEmail) {
      setGoToRegisterAfterMessage(true);
      showMessage(
        "Račun še ne obstaja",
        "Uporabnik s tem e-mailom še nima ustvarjenega računa. Za uporabo sistema moraš najprej opraviti registracijo.",
        "warning"
      );
      setRegisterData((old) => ({
        ...old,
        email,
        password: "",
        confirmPassword: "",
      }));
      return;
    }

    if (foundUserByEmail.password !== loginPassword) {
      showMessage("Prijava ni uspela", "Napačno geslo za ta e-mail.", "warning");
      return;
    }

    finishLogin(foundUserByEmail);
  }

  function logout() {
    localStorage.removeItem(SESSION_KEY);
    localStorage.removeItem(LOGIN_MESSAGE_KEY);
    setLoginMessage("");
    setIsTransitioning(false);
    setSelectedReservation(null);
    setDeleteTarget(null);
    setAppMessage(null);

    try {
      if (typeof msalInstance !== "undefined") {
        msalInstance.clearCache();
      }
    } catch (error) {
      console.warn("MSAL cache clear skipped", error);
    }

    setUser(null);
  }


  function hasConflict(newReservation) {
    return reservations.some((r) => {
      return (
        r.carId === newReservation.carId &&
        newReservation.start < r.end &&
        newReservation.end > r.start
      );
    });
  }

  function reserveCar(e) {
    e.preventDefault();

    if (!selectedCar) {
      showMessage("Vozilo ni izbrano", "Vozila se še nalagajo ali pa vozilo ni izbrano.", "warning");
      return;
    }

    if (selectedCar?.adminOnly && !user.isAdmin) {
      showMessage("Vozilo je zaklenjeno", "To vozilo lahko rezervira samo administrator.", "warning");
      return;
    }

    const start = `${startDate}T${startTime}`;
    const end = `${endDate}T${endTime}`;

    if (start >= end) {
      showMessage("Nepravilen termin", "Končni datum in ura morata biti kasneje kot začetni datum in ura.", "warning");
      return;
    }

    const reservedFor = reservationForMode === "other" ? reservationForName.trim() : getFullName(user);

    if (reservationForMode === "other" && !reservedFor) {
      showMessage("Manjka ime delavca", "Vpiši ime in priimek delavca, za katerega vnašaš rezervacijo.", "warning");
      return;
    }

    const newReservation = {
      id: crypto.randomUUID(),
      carId: selectedCar?.id,
      carName: selectedCar?.znamka,
      registracija: selectedCar?.registracija,
      userEmail: user.email,
      userName: reservedFor,
      reservedByEmail: user.email,
      reservedByName: getFullName(user),
      start,
      end,
    };

    if (hasConflict(newReservation)) {
      showMessage("Termin je zaseden", "To vozilo je v tem terminu že rezervirano. Izberi drugo vozilo ali drug termin.", "warning");
      return;
    }

    fetch("http://localhost:3001/reservations", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        carId: newReservation.carId,
        userEmail: newReservation.userEmail,
        userName: newReservation.userName,
        reservedByEmail: newReservation.reservedByEmail || newReservation.userEmail,
        reservedByName: newReservation.reservedByName || newReservation.userName,
        start: new Date(newReservation.start).toISOString().slice(0, 19).replace("T", " "),
        end: new Date(newReservation.end).toISOString().slice(0, 19).replace("T", " "),
      }),
    })
      .then(async (res) => {
        const data = await res.json();

        if (!res.ok) {
          showMessage(
            "Rezervacija ni uspela",
            data.error || "Pri shranjevanju rezervacije je prišlo do napake.",
            "warning"
          );
          return;
        }

        setReservations([
          ...reservations,
          {
            ...newReservation,
            id: String(data.id),
          },
        ]);

        showMessage("Rezervacija dodana", "Rezervacija je bila uspešno dodana v koledar.", "success");
      })
      .catch((err) => {
        console.error(err);
        showMessage(
          "Napaka povezave",
          "Rezervacije trenutno ni bilo mogoče shraniti v SQL bazo.",
          "warning"
        );
      });
    setReservationForMode("self");
    setReservationForName("");
  }

  function canDeleteReservation(reservation) {
    return user && (reservation.userEmail === user.email || user.isAdmin);
  }

  function deleteReservation(id) {
    const reservation = reservations.find((r) => r.id === id);
    requestDeleteReservation(reservation);
  }

  function handleDateClick(info) {
    setStartDate(info.dateStr);
    setEndDate(info.dateStr);
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function handleEventClick(info) {
    const r = reservations.find((reservation) => reservation.id === info.event.id);
    if (!r) return;
    setSelectedReservation(r);
  }

  function renderAppMessage() {
    if (!appMessage) return null;

    return (
      <div className="modal-backdrop" onClick={() => setAppMessage(null)}>
        <div className={`message-modal ${appMessage.type}`} onClick={(e) => e.stopPropagation()}>
          <div className="message-icon">
            {appMessage.type === "success" ? "✓" : appMessage.type === "warning" ? "!" : "i"}
          </div>

          <div className="message-content">
            <p className="eyebrow copper">Obvestilo sistema</p>
            <h2>{appMessage.title}</h2>
            <p>{appMessage.text}</p>
          </div>

          <div className="modal-actions message-actions">
            {goToRegisterAfterMessage ? (
              <>
                <button className="ghost-modal-button" onClick={() => {
                  setGoToRegisterAfterMessage(false);
                  setAppMessage(null);
                }}>
                  Zapri
                </button>
                <button className="primary-message-button" onClick={() => {
                  setAuthMode("register");
                  setGoToRegisterAfterMessage(false);
                  setAppMessage(null);
                }}>
                  Registracija
                </button>
              </>
            ) : (
              <button className="primary-message-button" onClick={() => setAppMessage(null)}>
                V redu
              </button>
            )}
          </div>
        </div>
      </div>
    );
  }

  if (isTransitioning) {
    return (
      <div className="transition-page">
        <div className="transition-card">
          <img src={logo} alt="Robeta logo" />

          <div className="loader"></div>

          {loginMessage ? (
            <>
              <h3>✔ Prijava uspešna</h3>
              <p>{loginMessage}</p>
              <p>Nalagam podatke ...</p>
              <p>Pripravljam koledar rezervacij ...</p>
            </>
          ) : (
            <p>Nalaganje sistema rezervacij ...</p>
          )}
        </div>
      </div>
    );
  }

  if (!user) {
    return (
      <div className="login-page">
        <div className="login-card fade-in">
          <img src={logo} alt="Robeta logo" className="login-logo" />
          <p className="eyebrow copper">Interni sistem za uslužbence</p>
          <h1>Rezervacije službenih vozil</h1>
          <p className="login-description">
            Zaposleni v pisarni se prijavijo s službenim Microsoft 365 računom. Rezervacijo lahko vpišejo tudi za delavca brez službenega e-maila.
          </p>

          <div className="login-company">
          <div> Robeta d.o.o.</div>
          </div>

          {authMode === "register" ? (
            <>
              <div className="auth-header-row">
                <button className="back-login-button" onClick={() => setAuthMode("login")}>
                  ← Nazaj na prijavo
                </button>
              </div>

              <form onSubmit={register} className="login-form">
                <div className="two-columns">
                  <label>
                    Ime
                    <input value={registerData.firstName} onChange={(e) => handleRegisterChange("firstName", e.target.value)} />
                  </label>
                  <label>
                    Priimek
                    <input value={registerData.lastName} onChange={(e) => handleRegisterChange("lastName", e.target.value)} />
                  </label>
                </div>
                <label>
                  E-mail
                  <input type="email" value={registerData.email} onChange={(e) => handleRegisterChange("email", e.target.value)} placeholder="ime.priimek@robeta.si" />
                </label>
                <div className="two-columns">
                  <label>
                    Geslo
                    <input type="password" value={registerData.password} onChange={(e) => handleRegisterChange("password", e.target.value)} />
                  </label>
                  <label>
                    Potrdi geslo
                    <input type="password" value={registerData.confirmPassword} onChange={(e) => handleRegisterChange("confirmPassword", e.target.value)} />
                  </label>
                </div>
                <button className="primary-button" type="submit">Ustvari račun</button>
              </form>
            </>
          ) : (
            <>
              <form onSubmit={loginMicrosoft} className="microsoft-login-form">
                <button className="m365-button" type="submit">
                  <span className="m365-icon">▦</span>
                  Prijava z Microsoft 365
                </button>
                <p className="hint">
                  Za dostop uporabite svoj službeni Microsoft 365 račun.
                </p>
              </form>
            </>
          )}
        </div>
        {renderAppMessage()}
      </div>
    );
  }

  return (
    <div className={isCalendarBig ? "app app-calendar-big fade-in" : "app fade-in"}>
      <header className="hero">
        <div className="hero-left">
          <img src={logo} alt="Robeta logo" className="hero-logo" />
          <div>
            <p className="eyebrow copper">Interni sistem za uslužbence Robete</p>
            <h1>Rezervacije službenih vozil</h1>
            <p>
              Aplikacija za rezervacijo službenih vozil in pregled razpoložljivosti voznega parka.
            </p>
          </div>
        </div>

        <div className="hero-actions">
          <div className="user-box">
            <span>Prijavljen:</span>
            <strong>{getFullName(user)}</strong>
            <small>{user.email}</small>
            <em>{getRoleLabel(user)}</em>
          </div>
          <button className="ghost-button" onClick={logout}>Odjava</button>
          <button className="calendar-toggle" onClick={() => setIsCalendarBig(!isCalendarBig)}>
            📅 {isCalendarBig ? "Pomanjšaj koledar" : "Razširi koledar"}
          </button>
        </div>
      </header>

      <section className="stats-grid">
        <div className="stat-card"><span>Vozila</span><strong>{dashboardCars.length}</strong></div>
        <div className="stat-card"><span>Vse rezervacije</span><strong>{dashboardReservations.length}</strong></div>
        <div className="stat-card"><span>Danes zasedena</span><strong>{busyCarsToday}</strong></div>
        <div className="stat-card"><span>Danes prosta</span><strong>{freeCarsToday}</strong></div>
      </section>

      <main className="layout">
        <section className="card form-card">
          <div className="section-title">
            <span>01</span>
            <h2>Nova rezervacija</h2>
          </div>

          <form onSubmit={reserveCar}>
            <div className="field-card">
              <span className="field-title">🚗 Izberi službeno vozilo</span>

              <div className="car-picker">
                <button
                  type="button"
                  className="car-picker-button"
                  onClick={() => setIsCarDropdownOpen(!isCarDropdownOpen)}
                >
                  <span>
                    <strong>{selectedCar ? getCarDisplayNameById(selectedCar?.id, selectedCar?.znamka, cars) : ""}</strong>
                    <small>{selectedCar?.registracija}</small>
                  </span>
                  <em>⌄</em>
                </button>

                {isCarDropdownOpen && (
                  <div className="car-picker-menu">
                    {cars.map((car) => {
                      const locked = car.adminOnly && !user.isAdmin;

                      return (
                        <button
                          type="button"
                          key={car.id}
                          className={
                            selectedCarId === car.id
                              ? "car-picker-option selected"
                              : locked
                                ? "car-picker-option locked"
                                : "car-picker-option"
                          }
                          disabled={locked}
                          onClick={() => {
                            if (locked) return;
                            setSelectedCarId(car.id);
                            setIsCarDropdownOpen(false);
                          }}
                        >
                          <span className="car-name">{getCarDisplayNameById(car.id, car.znamka, cars)}</span>
                          <span className="car-plate">{car.registracija}</span>
                          {locked && <span className="lock-badge">🔒</span>}
                        </button>
                      );
                    })}
                  </div>
                )}
              </div>


            </div>

            <DateTimePickerPanel
              title="📅 Začetek rezervacije"
              dateValue={startDate}
              timeValue={startTime}
              onDateChange={setStartDate}
              onTimeChange={setStartTime}
              hours={hours}
              minutes={minutes}
              openPicker={openPicker}
              setOpenPicker={setOpenPicker}
              prefix="start"
            />

            <DateTimePickerPanel
              title="📅 Konec rezervacije"
              dateValue={endDate}
              timeValue={endTime}
              onDateChange={setEndDate}
              onTimeChange={setEndTime}
              hours={hours}
              minutes={minutes}
              openPicker={openPicker}
              setOpenPicker={setOpenPicker}
              prefix="end"
            />

            <div className="field-card reservation-for-card">
              <span className="field-title">👤 Rezervacija za</span>

              <div className="mini-choice-tabs">
                <button
                  type="button"
                  className={reservationForMode === "self" ? "active" : ""}
                  onClick={() => {
                    setReservationForMode("self");
                    setReservationForName("");
                  }}
                >
                  Zase
                </button>
                <button
                  type="button"
                  className={reservationForMode === "other" ? "active" : ""}
                  onClick={() => setReservationForMode("other")}
                >
                  Drugega delavca
                </button>
              </div>

              {reservationForMode === "other" && (
                <label className="worker-name-field">
                  Ime in priimek delavca
                  <input
                    value={reservationForName}
                    onChange={(e) => setReservationForName(e.target.value)}
                    placeholder="npr. Luka Kovač"
                  />
                </label>
              )}
            </div>

            <div className="name-box">
              <h3>Rezervacijo vnaša</h3>
              <p>{getFullName(user)}</p>
              <small>{user.email} · {getRoleLabel(user)}</small>
            </div>

            <button className="primary-button" type="submit">Rezerviraj vozilo</button>
          </form>
        </section>

        <section className="card calendar-card">
          <div className="calendar-header">
            <div>
              <div className="section-title compact-title">
                <span>02</span>
                <h2>📅 Koledar rezervacij</h2>
              </div>
              <p>Na koledarju se vidi vozilo in oseba, ki ga je rezervirala.</p>
            </div>

            <div className="filter-label custom-filter">
              <span className="field-title">Filter</span>

              <div className="filter-picker">
                <button
                  type="button"
                  className="filter-picker-button"
                  onClick={() => setIsFilterDropdownOpen(!isFilterDropdownOpen)}
                >
                  <span>
                    <strong>{getFilterLabel(filterCarId, cars)}</strong>
                    <small>{getFilterSubLabel(filterCarId, cars)}</small>
                  </span>
                  <em>⌄</em>
                </button>

                {isFilterDropdownOpen && (
                  <div className="filter-picker-menu">
                    <button
                      type="button"
                      className={filterCarId === "all" ? "filter-picker-option selected" : "filter-picker-option"}
                      onClick={() => {
                        setFilterCarId("all");
                        setIsFilterDropdownOpen(false);
                      }}
                    >
                      <span>Vsa vozila</span>
                      <small>Celoten vozni park</small>
                    </button>

                    <button
                      type="button"
                      className={filterCarId === "mine" ? "filter-picker-option selected" : "filter-picker-option"}
                      onClick={() => {
                        setFilterCarId("mine");
                        setIsFilterDropdownOpen(false);
                      }}
                    >
                      <span>Moje rezervacije</span>
                      <small>Samo moje rezervacije</small>
                    </button>

                    {cars.map((car) => (
                      <button
                        type="button"
                        key={car.id}
                        className={filterCarId === car.id ? "filter-picker-option selected" : "filter-picker-option"}
                        onClick={() => {
                          setFilterCarId(car.id);
                          setIsFilterDropdownOpen(false);
                        }}
                      >
                        <span>{getCarDisplayNameById(car.id, car.znamka, cars)}</span>
                        <small>{car.registracija}</small>
                      </button>
                    ))}
                  </div>
                )}
              </div>
            </div>
          </div>

          <FullCalendar
            plugins={[dayGridPlugin, timeGridPlugin, interactionPlugin]}
            initialView="timeGridWeek"
            headerToolbar={{
              left: "prev,next today",
              center: "title",
              right: "dayGridMonth,timeGridWeek,timeGridDay",
            }}
            buttonText={{ today: "danes", month: "mesec", week: "teden", day: "dan" }}
            locale="sl"
            firstDay={1}
            height={isCalendarBig ? "80vh" : "620px"}
            events={calendarEvents}
            dateClick={handleDateClick}
            eventClick={handleEventClick}
            nowIndicator={true}
            eventTimeFormat={{
              hour: "2-digit",
              minute: "2-digit",
              hour12: false,
            }}
            slotLabelFormat={{
              hour: "2-digit",
              minute: "2-digit",
              hour12: false,
            }}
            allDaySlot={false}
            slotMinTime="06:00:00"
            slotMaxTime="22:00:00"
          />
        </section>
      </main>

      <section className="card list-card">
        <h2>Seznam rezervacij</h2>

        {filteredReservations.length === 0 ? (
          <p>Za izbrani filter ni rezervacij.</p>
        ) : (
          <div className="reservation-list">
            {filteredReservations.map((r) => (
              <div className={r.userEmail === user.email ? "reservation-item mine" : "reservation-item"} key={r.id}>
                <strong>{getCarDisplayNameById(r.carId, r.carName, cars)}</strong>
                <span>{r.registracija}</span>
                <span>{r.userName}</span>
                <span>{formatDateTime(r.start)} - {formatDateTime(r.end)}</span>
                <button
                  disabled={!canDeleteReservation(r)}
                  title={canDeleteReservation(r) ? "Izbriši rezervacijo" : "Lahko izbrišeš samo svojo rezervacijo"}
                  onClick={() => deleteReservation(r.id)}
                >
                  Izbriši
                </button>
              </div>
            ))}
          </div>
        )}
      </section>

      {appMessage && (
        <div className="modal-backdrop" onClick={() => setAppMessage(null)}>
          <div className={`message-modal ${appMessage.type}`} onClick={(e) => e.stopPropagation()}>
            <div className="message-icon">
              {appMessage.type === "success" ? "✓" : appMessage.type === "warning" ? "!" : "i"}
            </div>

            <div className="message-content">
              <p className="eyebrow copper">Obvestilo sistema</p>
              <h2>{appMessage.title}</h2>
              <p>{appMessage.text}</p>
            </div>

            <div className="modal-actions message-actions">
              {goToRegisterAfterMessage ? (
                <>
                  <button className="ghost-modal-button" onClick={() => {
                    setGoToRegisterAfterMessage(false);
                    setAppMessage(null);
                  }}>
                    Zapri
                  </button>
                  <button className="primary-message-button" onClick={() => {
                    setAuthMode("register");
                    setGoToRegisterAfterMessage(false);
                    setAppMessage(null);
                  }}>
                    Registracija
                  </button>
                </>
              ) : (
                <button className="primary-message-button" onClick={() => setAppMessage(null)}>
                  V redu
                </button>
              )}
            </div>
          </div>
        </div>
      )}

      {deleteTarget && (
        <div className="modal-backdrop" onClick={() => setDeleteTarget(null)}>
          <div className="message-modal delete-confirm" onClick={(e) => e.stopPropagation()}>
            <div className="message-icon danger-icon">×</div>

            <div className="message-content">
              <p className="eyebrow copper">Potrditev izbrisa</p>
              <h2>Izbrišem rezervacijo?</h2>
              <p>
                Rezervacija za {getCarDisplayNameById(deleteTarget.carId, deleteTarget.carName, cars)} ({deleteTarget.registracija}) bo odstranjena iz koledarja.
              </p>
            </div>

            <div className="modal-actions message-actions">
              <button className="ghost-modal-button" onClick={() => setDeleteTarget(null)}>
                Prekliči
              </button>
              <button className="danger-modal-button" onClick={confirmDeleteReservation}>
                Izbriši
              </button>
            </div>
          </div>
        </div>
      )}

      {selectedReservation && (
        <div className="modal-backdrop" onClick={() => setSelectedReservation(null)}>
          <div className="reservation-modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-top">
              <div>
                <p className="eyebrow copper">Podrobnosti rezervacije</p>
                <h2>{getCarDisplayNameById(selectedReservation.carId, selectedReservation.carName, cars)}</h2>
              </div>
              <button className="modal-close" onClick={() => setSelectedReservation(null)}>
                ×
              </button>
            </div>

            <div className="modal-plate">{selectedReservation.registracija}</div>

            <div className="modal-grid">
              <div className="modal-info">
                <span>Vozilo uporablja</span>
                <strong>{selectedReservation.userName}</strong>
              </div>
              <div className="modal-info">
                <span>E-mail</span>
                <strong>{selectedReservation.userEmail}</strong>
              </div>
              <div className="modal-info">
                <span>Rezervacijo vnesel</span>
                <strong>
                  {selectedReservation.reservedByName || selectedReservation.userName}
                  {" "}
                  ({selectedReservation.reservedByEmail || selectedReservation.userEmail})
                </strong>
              </div>
              <div className="modal-info">
                <span>Od</span>
                <strong>{formatDateTime(selectedReservation.start)}</strong>
              </div>
              <div className="modal-info">
                <span>Do</span>
                <strong>{formatDateTime(selectedReservation.end)}</strong>
              </div>
            </div>

            {!canDeleteReservation(selectedReservation) && (
              <p className="modal-warning">
                Te rezervacije ne moreš izbrisati, ker ni tvoja.
              </p>
            )}

            <div className="modal-actions">
              <button className="ghost-modal-button" onClick={() => setSelectedReservation(null)}>
                Zapri
              </button>

              {canDeleteReservation(selectedReservation) && (
                <button className="danger-modal-button" onClick={() => requestDeleteReservation(selectedReservation)}>
                  Izbriši rezervacijo
                </button>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
