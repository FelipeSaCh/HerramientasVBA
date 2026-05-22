# Herramientas VBA

Este repositorio contiene una colección de herramientas y automatizaciones desarrolladas en Visual Basic for Applications (VBA) para Microsoft Excel.

## 📂 Estructura del Proyecto

El proyecto se encuentra organizado en directorios según cada herramienta, lo cual facilita el mantenimiento y la separación entre el código fuente y los archivos ejecutables:

- **`Estado_SolP_Pedidos/`**: 
  Herramienta diseñada para automatizar, consultar o realizar seguimiento del estado de las Solicitudes de Pedido (SolP) y Pedidos de compra.
  - `Programa Estado SolP + Pedidos.xlsm`: Archivo de Excel que contiene la interfaz y macros integradas listo para usarse.
  - `Programa Estado SolP + Pedidos.bas`: Módulo de código fuente en VBA exportado (útil para control de versiones).

- **`Stock_Codigos_Material/`**: 
  Herramienta orientada a la gestión, consulta de stock y administración de códigos de materiales.
  - `Programa Stock + Codigos Material - Rev03.xlsm`: Archivo de Excel con macros habilitadas.
  - `Programa Stock + Codigos Material - Rev03.bas`: Módulo de código fuente en VBA exportado.

## ⚙️ Instrucciones de Uso

1. Descarga o realiza un `git clone` de este repositorio en tu equipo local.
2. Abre el archivo `.xlsm` de la herramienta que desees utilizar desde Microsoft Excel.
3. Asegúrate de **Habilitar el contenido** (las macros) cuando Excel muestre la advertencia de seguridad, para que los programas puedan ejecutarse correctamente.

*Nota técnica: Los archivos `.bas` están incluidos principalmente con propósitos de desarrollo y control de cambios. Si eres desarrollador, puedes abrir el Editor de Visual Basic (Alt + F11) en Excel para explorar, modificar o importar estos módulos de código.*