STABILNI POPRAVEK

Popravljeno:
- odstranjena napaka cars is not defined
- helper funkcije zdaj dobijo cars kot parameter
- cars useState/useEffect je znotraj function App()
- SQL podatki iz /cars se pretvorijo v obliko, ki jo React pričakuje
- availableCars se osveži, ko se cars naloži iz SQL
- selectedCar ima varnostni fallback

Zagon:
1. npm install
2. node server/server.js
3. nov terminal: npm run dev
