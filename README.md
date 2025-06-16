Instalador de Microsoft Office (2016 - 2024) - Digitalcode Spa

Este script permite instalar diferentes versiones de Microsoft Office, Visio y Project (2016 a 2024) tanto en versión Retail como LTSC (Volume).

⚠️ Advertencia: Esta herramienta está orientada a usuarios con conocimientos técnicos avanzados en sistemas Windows, instalación de software y administración de equipos. El mal uso puede provocar conflictos con instalaciones existentes de Office.

📦 Características

Instalación por configuración XML para versiones 2019, 2021 y 2024 (Office, Visio, Project).
- Instalación directa vía descarga oficial para Office 2016 (Retail).
- Soporta arquitectura 32-bit y 64-bit.
- Idiomas disponibles: Español (es-ES), Inglés (en-US), Portugués (pt-BR).
- Validación de permisos de administrador.
- Prevención de instalación si Office ya está instalado.

🛠️ Requisitos

Windows 10 o superior
PowerShell habilitado
Conexión a internet (para instalación directa o descarga de componentes)

📥 Cómo usar

Descarga el repositorio o clona con:

git clone https://github.com/tuusuario/instalador-office-digitalcode.git

Coloca el archivo setup.exe del Office Deployment Tool en la misma carpeta (solo requerido para versiones 2019, 2021, 2024).

Ejecuta el archivo instalador_office_2016_2024.bat como administrador:

Clic derecho → Ejecutar como administrador.

🧭 Opciones del instalador

Selecciona si deseas instalar:
- Versiones RETAIL
- Versiones LTSC (Volume)

Elige el producto:
- Office (Pro Plus / Standard)
- Visio (Pro / Standard)
- Project (Pro / Standard)

Desde 2016 a 2024

Para Office 2016 Retail se descargará automáticamente el instalador desde Microsoft.

Elige la arquitectura y el idioma.

📌 Notas importantes

- Office 2016 Retail no funciona con ODT, por eso se usa un ejecutable directo desde Microsoft.
- Asegúrate de no tener otra versión de Office instalada antes de ejecutar el script.
- El script eliminará el archivo de configuración temporal una vez finalizada la instalación.

🧑‍💻 Autor

- Digitalcode Spa
- Desarrollado por: Freddy Beardesley
- Correo: ventas@digitalcode.cl


📄 Licencia
- Este proyecto está licenciado bajo los términos de la licencia MIT.
