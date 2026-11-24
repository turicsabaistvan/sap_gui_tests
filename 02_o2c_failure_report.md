# Jelentés — `02_o2c_create_sales_order.robot` futáshiba

**Vizsgált futás:** 2026-11-24 07:09:06 – 07:09:39
**Eszköz:** Robot Framework 7.5.1 (Python 3.15.2, win32)
**Forrás:** `reports.xml` (output.xml), `log.html`, `report.html`
**Teszteset:** `Create Sales Order And Verify In VA03 And VBAK Table` (`02_o2c_create_sales_order.robot:48`)
**Eredmény:** 0 PASS / 1 FAIL, összidő 33,13 s

---

## 1. Röviden (TL;DR)

A teszt **el sem indult érdemben**: a `[Setup]`-ban futó `Open Clean SAP Session` kulcsszó
`Connect To Active Session` lépése nem tudott csatlakozni egyetlen SAP GUI sessionhöz sem,
ezért a futás **32,97 s után a setup-ban megállt**.

> **Nem egy SAP-képernyő egy elemét nem találta meg.** A hiba a bejelentkezés/COM-csatlakozás
> szintjén történt, jóval az első képernyő (SAP Easy Access) előtt. Az SAP GUI ablak el sem jött
> létre (vagy legalábbis nem regisztrálta magát a COM Running Object Table-ben), így nincs `wnd[0]`,
> amiben bármit keresni lehetne.

Az első elem, amit a teszt **keresni akart volna** (de már el sem érte):
`wnd[0]/tbar[0]/okcd` — SAP Easy Access (S000) kezdőképernyő, bal felső tranzakciómező
(`02_o2c_create_sales_order.robot:40`).

---

## 2. A pontos hibaüzenet

```
Setup failed:
Could not attach to SAP GUI session within 30.0s:
(-2147221020, 'Invalid syntax', None, None)
```

Helye: `SapComLibrary.py:38` (`connect_to_active_session`), hívva innen:
`02_o2c_create_sales_order.robot:38`.

A `-2147221020` decimális érték hexadecimálisan **`0x800401E4` = `MK_E_SYNTAX`**
(COM moniker: „Invalid syntax”).

### Mit jelent ez a gyakorlatban

`SapComLibrary.connect_to_active_session()` ezt próbálja 0,5 s-onként 30 s-ig:

```python
sap_gui = win32com.client.GetObject("SAPGUI")
engine  = sap_gui.GetScriptingEngine
```

A `win32com.client.GetObject("SAPGUI")` a futó SAP GUI példányt keresi a COM Running Object
Table-ben. Ha **nincs futó, scriptelésre felkészített SAP GUI**, a pywin32 a moniker
feloldásakor `MK_E_SYNTAX` hibát dob — pontosan ezt látjuk. A ciklus ezt a kivételt elnyeli,
ezért a 30 s letelte után csak az utolsó ilyen hibaüzenet jelenik meg. **A hibaüzenet ezért
nem alkalmas arra, hogy megkülönböztesse egymástól a lehetséges okokat** (SAP GUI nem fut /
scripting kikapcsolva / nem indult el a kapcsolat) — ezt a könyvtár nem naplózza részletesen.

---

## 3. Melyik oldalon melyik elemet nem találta?

**Rövid válasz: sehol.** A futás a 0. lépésnél (session csatlakozás) állt meg, az SAP GUI-ban
egyetlen képernyő sem jelent meg. A `wnd[0]` gyökér sem létezett, ezért a `FindById(...)`
hívásokra nem is került sor (a reportban minden ez utáni kulcsszó `NOT RUN`).

Az alábbi táblázat a teszt által **tervezett** lokátorokat mutatja, képernyő szerint, és azt,
hogy a mostani futás meddig jutott el:

| # | Lépés | SAP GUI ID | Képernyő (SAP oldal) | Elérte? |
|---|-------|-----------|----------------------|---------|
| 0 | Csatlakozás | `SAPGUI` COM objektum (nem elem) | — (nincs képernyő) | ❌ **itt állt meg** |
| 1 | Session ellenőrzés | `wnd[0]/tbar[0]/okcd` | SAP Easy Access (S000) | ❌ NOT RUN |
| 2 | VA01 indítás | `wnd[0]/usr/ctxtVBAK-AUART` | Create Sales Order: Initial Screen (VA01) | ❌ NOT RUN |
| 3 | 〃 | `wnd[0]/usr/ctxtVBAK-VKORG` | VA01 Initial Screen | ❌ NOT RUN |
| 4 | 〃 | `wnd[0]/usr/ctxtVBAK-VTWEG` | VA01 Initial Screen | ❌ NOT RUN |
| 5 | 〃 | `wnd[0]/usr/ctxtVBAK-SPART` | VA01 Initial Screen | ❌ NOT RUN |
| 6 | Fejléc – Sold-To | `.../subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR` | Create Sales Order: Overview (VA01) | ❌ NOT RUN |
| 7 | Fejléc – Ship-To | `.../subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR` | VA01 Overview | ❌ NOT RUN |
| 8 | Fejléc – PO szám | `.../txtVBKD-BSTKD` | VA01 Overview | ❌ NOT RUN |
| 9 | Tétel – anyag | `.../tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]` | VA01 Overview, első tételsor | ❌ NOT RUN |
| 10 | Tétel – mennyiség | `.../tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,0]` | VA01 Overview, első tételsor | ❌ NOT RUN |
| 11 | Mentés | `wnd[0]/tbar[0]/btn[11]` | VA01 Overview (Save) | ❌ NOT RUN |
| 12 | VA03 megnyitás | `wnd[0]/usr/ctxtVBAK-VBELN` | Display Sales Order: Initial (VA03) | ❌ NOT RUN |
| 13 | VA03 ellenőrzés | `.../txtVBKD-BSTKD` | Display Sales Order: Overview (VA03) | ❌ NOT RUN |
| 14 | SE16N indítás | `wnd[0]/usr/ctxtGD-TAB` | SE16N kezdőképernyő (Table Name) | ❌ NOT RUN |
| 15 | VBAK szűrés | `wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,1]` | SE16N szelekciós képernyő | ❌ NOT RUN |

Tehát a konkrét futásban **nincs „nem talált elem”** — a hiba a csatlakozás előtt keletkezett.
A fenti elemek létezését sem megerősíteni, sem cáfolni nem lehet ebből a logból, mert egyik
képernyő sem töltődött be.

> Megjegyzés: a projektben lévő `sap_dump.json` **nem** segít ebben — az 2026-11-23 21:52-kor
> készült, és az **SAP Easy Access (S000)** kezdőképernyő elemeit tartalmazza (168 elem,
> `mbar`, `tbar[0]`, `IMAGE_CONTAINER` stb.), nem a VA01/VA03/SE16N képernyőket.

---

## 4. Idővonal — bizonyítékok a reportból

| Időpont | Kulcsszó | Eredmény | Megjegyzés |
|---------|----------|----------|------------|
| 07:09:07.10 | `Kill Existing Sap Processes` | PASS (1,35 s) | `taskkill /F` a `sapshcut/sapgui/saplogon` processzekre + 1 s sleep |
| 07:09:08.45 | `Start Process` | PASS (0,02 s) | `sapshcut.exe -guiparm=/H/207.175.121.75/S/3200 -system=ABA -client=800 -user=IDADMIN -pw=*** -language=EN` |
| 07:09:08.47 | `Connect To Active Session` | **FAIL (30,14 s)** | `MK_E_SYNTAX`, 30 s timeout |
| 07:09:38.62 | `Run Transaction /nex` (teardown) | FAIL (elnyelve) | `AttributeError: 'NoneType' object has no attribute 'FindById'` |
| 07:09:38.62 | `Terminate Process` | PASS | a `sapshcut` processz kilövése |
| 07:09:38.63 | `Kill Existing Sap Processes` | PASS | takarítás |

Fontos: a `Start Process` **sikeres** — az `sapshcut.exe` létezik és elindult. A hiba nem az,
hogy az exe nem található, hanem hogy **az SAP GUI nem regisztrálta magát COM-ban 30 másodpercen
belül**.

---

## 5. Lehetséges gyökér-okok (valószínűségi sorrendben)

### ① A célrendszer / host nem elérhető vagy megváltozott — **legvalószínűbb**

- A mostani futás a **`207.175.121.75`** hostra próbál csatlakozni
  (`02_o2c_create_sales_order.robot:7`, `.env:4`).
- A korábbi, működő futások a memóriában rögzített környezet szerint a **`34.76.24.4`**
  (`vhcalabaci`, ABA, client 800) hostra mentek — ott a teszt eljutott a VA01-ig (és csak
  később, a naptár/fiskális év miatt hasalt el).
- A `.env` és a `.robot` fájl **2026-11-24 07:03-kor, illetve 07:09-kor** módosult (a futás
  előtt pár perccel), a szerver IP-je ekkor változott `207.175.121.75`-re.
- Ha a host nem elérhető (tűzfal, VPN, leállt szerver) vagy rossz IP, az SAP GUI nem tud
  sessiont nyitni, így a `SAPGUI` COM objektum sem jön létre → pontosan ez a hiba.

**Ellenőrzés (futtatás nélkül is megtehető):** `ping 207.175.121.75`, illetve a 3200-as port
elérhetősége; a korábbi, működő `34.76.24.4` visszaállításával próbafutás.

### ② Kliensoldali SAP GUI Scripting kikapcsolva

SAP GUI → *Options → Accessibility & Scripting → Scripting → „Enable scripting”*.
Ha ez nincs bekapcsolva, a `GetObject("SAPGUI")` **ugyanezt** az `MK_E_SYNTAX` hibát adja.
Mivel a project a `SapComLibrary.py`-vel (nyers `win32com`) dolgozik, a scripting beállítás
kritikus. (Korábban működött, ezért ez csak akkor valószínű, ha a gépen/beállításban történt
változás.)

### ③ `sapshcut.exe` biztonsági megkötés / SAP Logon hiánya

A setup előbb kilövi a `saplogon.exe`-t, majd `sapshcut.exe`-vel indít. Egyes telepítéseken a
shortcut-indítás tiltott (SAP GUI security policy), ilyenkor a `sapshcut` azonnal kilép vagy
csak a Logon padet nyitja meg, session nélkül. Ezt támasztja alá, hogy a `Start Process`
sikeres, de 30 s alatt semmi nem regisztrálódik.

### ④ Az SAP GUI indítása 30 s-nál tovább tart

Első indítás / lassú gép / vírusirtó esetén előfordulhat, de 30 s általában elegendő; a
memóriában rögzített korábbi futásoknál a csatlakozás gyors volt. Kevésbé valószínű.

### ⑤ 32/64 bit vagy COM-regisztrációs probléma

Ha az SAP GUI vagy a Python bitessége változott, a `SAPGUI` ProgID nem oldható fel. Korábban
működött, ezért csak környezetváltozás (SAP GUI frissítés/újratelepítés) esetén jön szóba.

---

## 6. Másodlagos hiba a teardown-ban (tünet, nem ok)

```
Run Transaction    /nex
→ AttributeError: 'NoneType' object has no attribute 'FindById'
```

Ez **következmény**: mivel a csatlakozás nem sikerült, `SapComLibrary.session` értéke `None`
maradt (`SapComLibrary.py:11`, illetve a `connect_to_active_session` nem ért el a `return`-ig).
A `Run Keyword And Ignore Error` elnyeli, ezért a teszt végeredményét nem ez, hanem a setup
hibája határozza meg. A könyvtárban érdemes lenne a `session is None` esetre külön, beszédes
hibát adni (pl. „no SAP GUI session — check scripting/host”).

---

## 7. Statikus kockázatok a későbbi lépésekben (nem ez a futás)

Ezek **nem okozták** a mostani hibát (a kód el sem ért odáig), de a csatlakozás helyreállítása
után valószínűleg ezek közül kerül ki a következő hiba:

| Elem | Képernyő | Kockázat |
|------|----------|----------|
| `wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,1]` | SE16N szelekciós | A sorindex (`[2,1]`) a megjelenített mezőktől / felhasználói beállításoktól függ; a VBELN nem biztos, hogy a 2. sorban van. A legtörékenyebb lokátor a tesztben. |
| `.../ctxtRV45A-MABNR[1,0]` és `.../txtRV45A-KWMENG[2,0]` | VA01 Overview tételtábla | A táblázat-vezérlő (`GuiTableControl`) oszlop/sor indexei GUI-verzió és testreszabás függvényében eltérhetnek; a `subSUBSCREEN_TC:SAPMV45A:4900` útvonal is verziófüggő. |
| `wnd[0]/usr/ctxtVBAK-AUART`, `-VKORG`, `-VTWEG`, `-SPART` | VA01 Initial | Standard mezők, de ha a user default (SU3) paraméterezése miatt a VA01 közvetlenül az Overview-ra ugrik, az Initial képernyő kimarad. |
| `wnd[0]/usr/ctxtVBAK-VBELN` | VA03 Initial | Ha a VA03-at üresen hívják és a user default más tranzakciót indít, eltérhet; alapból rendben. |
| `txtVBKD-BSTKD` | VA01/VA03 Overview fejléc | A fejléc `SUBSCREEN_HEADER:SAPMV45A:4021` útvonal standard; megjelenítés (VA03) módban is létezik, de csak akkor, ha a PO-mező az adott nézetben látszik. |

Továbbá érdemes tudni: a `sap_dump.json` alapján készült ellenőrzés **nem használható** a fenti
lokátorok validálására, mert az a kezdőképernyő dumpja.

---

## 8. Biztonsági megjegyzés (a report tartalma)

A `Start Process` argumentumai a **jelszót nyílt szövegben** tartalmazzák, és a Robot Framework
ezt a parancssort beírja az `output.xml`-be / `log.html`-be / `report.html`-be:

```
... -user=IDADMIN -pw=ides123 -language=EN
```

Ez a jelszó jelenleg a riportokban és a `.robot` fájl default értékében is szerepel
(`02_o2c_create_sales_order.robot:12`). Javasolt: környezeti változóból átadni (már most is
`%{SAP_PASS=...}`), a reportból kitakarni, és a jelszót cserélni, ha a riport megosztásra kerül.

---

## 9. Architektúra / CLAUDE.md megfelelés

- **Megsértett szabály:** a `CLAUDE.md` szerint tilos SAP GUI ID-kat (`wnd[0]/usr/...`)
  közvetlenül `.robot` fájlba írni — minden lokátornak `resources/pages/*.resource` fájlban
  kell lennie. A `02_o2c_create_sales_order.robot` ezt jelenleg nem követi, és a
  `resources/`, `tests/`, `variables/` könyvtárak sem léteznek.
- **`.env` nem érvényesül:** a `env_loader.py` `get_variables()`-t definiál, de sem a
  `.robot` fájl, sem a suite nem hivatkozik rá (`Variables env_loader.py` hiányzik). A Robot
  csak az operációs rendszer környezeti változóit látja, tehát a `.env` tartalma nem töltődik
  be automatikusan; a tényleges értékek a `.robot`-ban lévő defaultok. Emiatt a
  `sap_gui_test.robot` (default `34.76.24.4`) és a `02_o2c_create_sales_order.robot` /
  `.env` (`207.175.121.75`) **eltérő szerverre mutat** — ez önmagában is zavarforrás.
- `sap_dump.py` a `sap_dump2.json`-ba ír, de a konzolra `sap_dump.json`-t ír ki — elavult,
  félrevezető dump.
- A suite-nak nincs `[Documentation]` / tag / `Suite Setup`-szintű konfigurációja; a
  teardown hibája elnyelve van, így a logban könnyen elvész.

---

## 10. Javasolt javítási sorrend

1. **Host ellenőrzése:** melyik a helyes szerver — `207.175.121.75` vagy a korábbi
   `34.76.24.4`? Ping + 3200-as port teszt; ha a régi host él, állítsuk vissza, és nézzük meg,
   hogy a csatlakozás működik-e.
2. **SAP GUI Scripting ellenőrzése** a kliensen (Options → Accessibility & Scripting →
   Scripting → Enable scripting), majd SAP GUI újraindítás.
3. **Kézi indítás próbája:** `sapshcut.exe` ugyanezekkel a paraméterekkel — megjelenik-e a
   bejelentkező ablak, vagy hibát dob?
4. A `SapComLibrary.connect_to_active_session`-ben a hibák megkülönböztetése és a
   `session is None` eset kezelése (beszédes hibaüzenet, ne `NoneType.FindById`).
5. A csatlakozás helyreállítása után a 7. pontban felsorolt lokátorok validálása friss
   `sap_dump`-pal (a tényleges VA01 / VA03 / SE16N képernyőkről).
6. A lokátorok kiemelése `resources/pages/*.resource` fájlokba a `CLAUDE.md` szerint; a
   `.env` bekötése `Variables` fájlként; a jelszó kivétele a riportokból.

---

### Összefoglaló egy mondatban

A teszt nem azért hasalt el, mert egy SAP-képernyőn nem talált egy elemet, hanem mert a
`sapshcut`-tal indított SAP GUI **nem regisztrált COM sessiont 30 másodpercen belül**
(`MK_E_SYNTAX`), és a legvalószínűbb ok a frissen átírt, nem elérhető vagy hibás
`207.175.121.75` célrendszer (illetve kliensoldalon kikapcsolt SAP GUI Scripting).
