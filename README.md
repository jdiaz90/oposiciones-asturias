# OposicionesAsturias

**OposicionesAsturias** es una aplicación en Python que automatiza la consulta de oposiciones y concursos publicados en el **Boletín Oficial del Estado (BOE)** relacionados con el Principado de Asturias.

El programa descarga los sumarios oficiales de los últimos días, filtra las disposiciones relevantes y muestra los resultados en una tabla clara en consola.  
Además, permite exportar la información a un archivo Excel con formato profesional, incluyendo notas específicas como *Turno libre* o *Promoción interna* extraídas directamente del XML oficial del BOE.

---

## ✨ Funcionalidades

- 📥 Descarga automática de los **sumarios del BOE** de los últimos días.
- 🔎 Filtrado de convocatorias de la sección *Oposiciones y concursos* que mencionan Asturias.
- 📑 Lectura del **XML completo de cada disposición** para extraer información adicional del bloque `<notas>`.
- 📝 Identificación de notas relevantes como:
  - *Turno libre*
  - *Promoción interna*
- 📊 Visualización en consola en formato tabla.
- 📂 Exportación a Excel con:
  - Fecha de publicación
  - Ayuntamiento convocante
  - Título completo de la disposición
  - Nota de turno (*Turno libre* / *Promoción interna*)
  - Enlace directo al BOE (hipervínculo clicable)

---

## 📊 Ejemplo de salida en consola

```
📊 RESULTADOS: Oposiciones y concursos en Asturias (últimos 15 días)
Fecha       | Ayuntamiento              | Nota turno                           | Enlace
12/11/2025  | Ayuntamiento de Oviedo    | Turno libre: Encargado/a de Obras   | https://www.boe.es/diario_boe/txt.php?id=BOE-A-2025-12345
12/11/2025  | Ayuntamiento de Gijón     | Promoción interna: Técnico/a        | https://www.boe.es/diario_boe/txt.php?id=BOE-A-2025-12346
```

---

## 📂 Ejemplo de Excel generado

- Encabezados destacados con fondo azul y texto blanco.  
- Colores alternos en las filas para facilitar la lectura.  
- Columnas ajustadas automáticamente al contenido.  
- Hipervínculos clicables en la columna de enlace.  

---

## 🚀 Instalación y uso

1. Clona este repositorio:
   ```bash
   git clone https://github.com/tuusuario/OposicionesAsturias.git
   cd OposicionesAsturias
   ```

2. Instala las dependencias necesarias:
   ```bash
   pip install requests openpyxl
   ```

3. Ejecuta el programa:
   ```bash
   python oposiciones_asturias.py
   ```

4. El programa mostrará los resultados en consola y te preguntará si deseas exportarlos a Excel.

---

## 🎯 Público objetivo

- **Opositores**: localizar rápidamente convocatorias en Asturias.
- **Administraciones públicas**: seguimiento de procesos selectivos.
- **Profesionales del sector jurídico y educativo**: disponer de información organizada y exportable.

---

## 📌 Próximas mejoras

- Soporte para más comunidades autónomas.
- Descarga automática de las bases completas desde el BOPA.
- Filtros avanzados por tipo de plaza o cuerpo.

---

## 📄 Licencia

Este proyecto se distribuye bajo la licencia MIT. Consulta el archivo `LICENSE` para más detalles.
