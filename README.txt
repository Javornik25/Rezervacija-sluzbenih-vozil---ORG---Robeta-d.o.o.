ROBETA - FINAL MICROSOFT 365 VERZIJA

VKLJUČENO
- Prava Microsoft 365 prijava prek MSAL.
- Uporabnik geslo vpiše v Microsoft prijavno okno, ne v aplikacijo.
- Tenant ID in Client ID sta vpisana v kodo.
- Lokalna prijava ostane kot rezervna/testna možnost.
- Login stran ima navadno prijavo + ALI + Microsoft 365 gumb.
- Uporabniki z @robeta.si dobijo naziv "pisarna".
- stas.javornik@robeta.si je administrator.
- Pisarna lahko rezervira vozilo za delavca brez službenega e-maila.
- Audi/BMW so vidni vsem, rezervira jih lahko samo administrator.
- Navadni uporabniki imajo statistiko za 6 vozil.
- Admin ima statistiko za vseh 10 vozil.
- Ikone 🚗 in 🔒🚗 so dodane pri vozilih.
- Robeta custom modalna okna in dropdowni.

MICROSOFT ENTRA NASTAVITEV
V Microsoft Entra mora biti dodan Redirect URI:

http://localhost:5173

Platforma:
Single-page application (SPA)

ID-JI
Tenant ID:
b910a59c-fb77-4328-9e53-593dfdba7617

Application / Client ID:
89da1f76-52a6-4278-9349-92ca4cde4a69

ZAGON
npm install
npm run dev

ODPRI
http://localhost:5173

TEST LOKALNA PRIJAVA
stas.javornik@robeta.si / admin123
jan.novak@robeta.si / test123

OPOMBA
Če Microsoft prijava ne deluje, preveri, ali je Redirect URI res dodan kot SPA in ne kot Web.


POPRAVEK
- Microsoft prijava uporablja loginRedirect() namesto loginPopup().


POPRAVEK
- Koledar se začne s ponedeljkom.
- Popravljena odjava, da ne ostane black screen.
- Dodan prehod po Microsoft prijavi.
