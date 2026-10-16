# Securitatea aplicațiilor web și a serverelor – Laboratoare

Pe parcursul semestrului lucrezi pe **o singură aplicație**: „Notițe” (NestJS + React + PostgreSQL, în Docker). Primești aplicația gata făcută, cu vulnerabilități ascunse intenționat. Le găsești, le repari, apoi o publici în siguranță pe propriul server.

| Lab | Tema | Fișa |
|---|---|---|
| 1 | Pornește aplicația și găsește riscurile | [lab1-docker](lab1-docker/README.md) |
| 2 | SQL Injection și XSS | [lab2-injectii](lab2-injectii/README.md) |
| 3 | Securizează login-ul și accesul la date | [lab3-autentificare](lab3-autentificare/README.md) |
| 4 | Securizează serverul Linux (DigitalOcean) | [lab4-server-linux](lab4-server-linux/README.md) |
| 5 | Deploy și securizarea serverului web (Nginx, HTTPS, GitHub Actions) | [lab5-server-web](lab5-server-web/README.md) |

## Ce îți trebuie

- Laptop cu **Linux** (Ubuntu recomandat), minimum 8 GB RAM
- **Docker** și Docker Compose, **Node.js 22**, **Git**, **VS Code**
- Cont **GitHub** și cont **DigitalOcean** (din Lab 4)

Tot software-ul este gratuit. Serverul DigitalOcean (aproximativ 12 $/lună) se plătește din creditul pentru studenți.

## Reguli

- Atacurile se fac **doar** pe aplicația ta de pe laptop (`localhost`) și pe propriul server.
- Pe serverul public ajunge doar codul **reparat**.
- Fișierul `.env` (parole, chei) **nu** se urcă niciodată pe GitHub.
