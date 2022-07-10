DMNUtils:

DMNUtils es un conjunto de herramientas para interactuar con archivos DMN (Decision Model and Notation), y poder realizar acciones con los mismos que habitualmente no se pueden llevar a cabo.
Actualmente cuenta con tres funcionalidades:

* Exportar Tablas de Decisión (decisionTable) a formato XLSX (Excel).

* Generar explicación en lenguaje natural de las Tablas de Decisión (decisionTable) que conforman el DMN.

* Reemplazar una Tabla de Decisión (decisionTable) por el contenido de un archivo XLSX (Excel), el cual debe contar con un formato standard para que el reemplazo funcione correctamente.

Cuenta con una interfaz de usuario simple creada en PyQt5, y controles de formato simples para asegurar que el archivo DMN y/o XLSX elegido por el usuario reúna los requisitos mínimos para ser procesado.

Requiere para su funcionamiento de las siguientes librerías de Python:

- lxml.etree
- openpyxl
- PyQt5
- uuid
- docx


Gervasio Frugoni

10/07/2022
