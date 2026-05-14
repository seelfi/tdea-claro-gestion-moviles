# 📱 TDEA - Claro Colombia | Gestión Dispositivos Móviles

App desarrollada en **Google Apps Script** como proyecto académico del **Tecnológico de Antioquia (TDEA)**, que simula la operación del área de dispositivos móviles de **Claro Colombia**.

🔗 **[Abrir aplicación](https://script.google.com/macros/s/AKfycbybzWf7p0MbB0Vo89TKhPZxS758cFgWqzQFfkwKWirindikWYiyWSBOYVspPogPsvQR-g/exec)**

---

## 🚀 Módulos

| Módulo | Descripción |
|---|---|
| 🔐 **Login** | Autenticación de usuarios por rol |
| 🏠 **Home** | Panel principal de navegación |
| 🛒 **Ventas** | Registro y seguimiento de ventas de dispositivos |
| 📦 **Inventario** | Control de stock de equipos móviles |
| 👥 **Empleados** | Gestión de personal y roles |
| 🖥️ **PDV** | Punto de venta |
| 🧾 **Facturación** | Generación y consulta de facturas |

---

## 🛠️ Tecnologías

- **Frontend/Backend:** Google Apps Script
- **Interfaz:** Web App (HTML, CSS, JavaScript)
- **Base de datos:** SQL Server — Azure SQL Database
- **Control de versiones:** Git + GitHub
- **Sincronización:** Clasp (CLI de Apps Script)

---

## 🗄️ Base de Datos

- **Motor:** SQL Server
- **Servidor:** srv-clarotdea.database.windows.net
- **Base de datos:** db-clarotdea
- **Usuario:** adminclaro
- **Contraseña:** solicitar al autor

📁 Exportación disponible en `/database/db-clarotdea-2026-5-13-20-27.bacpac`

---

## 📁 Estructura del proyecto

```
tdea-claro-gestion-moviles/
├── database/
│   └── db-clarotdea-2026-5-13-20-27.bacpac   ← Exportación de la BD
├── Code.js                ← Lógica principal
├── Code_Azure.js          ← Conexión con Azure SQL
├── Code_Ventas.js         ← Módulo de ventas
├── Code_Inventario.js     ← Módulo de inventario
├── Code_Empleados.js      ← Módulo de empleados
├── Code_PDV.js            ← Punto de venta
├── Code_Factura.js        ← Módulo de facturación
├── Code_Home.js           ← Panel principal
├── Index.html             ← Entrada de la app
├── Login.html             ← Pantalla de login
├── Home.html              ← Vista principal
├── Estilos.html           ← Estilos globales
├── Estilos_Home.html      ← Estilos del home
├── Scripts.html           ← Scripts globales
├── Scripts_Home.html      ← Scripts del home
├── Scripts_Facturas.html  ← Scripts de facturación
└── appsscript.json        ← Configuración del proyecto
```

---

## ▶️ Cómo ejecutar

1. Abrir el [link de la aplicación](https://script.google.com/macros/s/AKfycbybzWf7p0MbB0Vo89TKhPZxS758cFgWqzQFfkwKWirindikWYiyWSBOYVspPogPsvQR-g/exec)
2. Iniciar sesión con las credenciales proporcionadas
3. Navegar por los módulos desde el panel principal

---

## 👨‍💻 Autor

- **Institución:** Tecnológico de Antioquia — TDEA
- **Proyecto:** Simulación operativa Claro Colombia — Dispositivos Móviles
