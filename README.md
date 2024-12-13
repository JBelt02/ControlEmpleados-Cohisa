Control de Operarios

Este repositorio contiene un sistema para la gestión y control de operarios en un entorno de fabricación. La aplicación permite visualizar información relevante sobre los trabajadores, las órdenes en fabricación y las órdenes finalizadas, con opciones adicionales para generar etiquetas al completar las órdenes. Todos los datos se extraen directamente de Sage200 y se presentan en una interfaz intuitiva accesible desde múltiples pantallas dentro de la empresa.

Características Principales

Gestión de Operarios

Visualización de Trabajadores: Lista de operarios activos y su estado actual.

Gestión de Órdenes:

Visualización de órdenes en fabricación.

Información detallada de cada orden en pestañas dedicadas.

Generación de Etiquetas:

Creación automática de etiquetas al finalizar las órdenes.

Compatible con múltiples formatos de etiquetas para impresión.

Interfaz Gráfica

Diseño Intuitivo:

Organización en pestañas para facilitar la navegación.

Visualización clara de información relevante.

Actualización en Tiempo Real:

Datos sincronizados directamente desde Sage200.

Actualización automática de estados y detalles de órdenes.

Integración con Sage200

Extracción de Datos:

Información de trabajadores, órdenes y materiales directamente desde Sage200.

Sincronización Bidireccional:

Registro de finalización de órdenes y otros eventos en Sage200.

Requisitos del Sistema

Sage200 instalado y configurado.

Python 3.11.7 o superior.

Dependencias adicionales (ver sección de instalación).

Instalación

Clona el repositorio:

git clone https://github.com/tuusuario/control-operarios.git

Instala las dependencias necesarias utilizando pip:

pip install -r requirements.txt

Configura los parámetros de conexión en el archivo config.json, incluyendo las credenciales de la base de datos.
Licencia

Este proyecto está licenciado bajo la Licencia MIT.

Contribuciones

¡Las contribuciones son bienvenidas! Por favor, crea un fork del repositorio, realiza tus cambios y envía un pull request.

Contacto

Para preguntas o soporte, contacta a granjuan02@gmail.com .
