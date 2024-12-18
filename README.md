# VisualBasic-Mega# Curso de Visual Basic 6.0 - Mega Semillero


## Objetivos del Curso

- Conocer el entorno de desarrollo de Visual Basic 6.0, sus controles y propiedades más importantes.
- Aprender las estructuras de control esenciales (If, Select Case, For, Do While, etc.).
- Implementar el manejo básico de errores en VB6 (On Error).
- Aprender a conectarse a una base de datos (por ejemplo, Access) a través de ADO, ejecutar consultas SQL y mostrar resultados en formularios.
- Trabajar con archivos externos (lectura/escritura de texto, carga de imágenes).
- Desarrollar un proyecto práctico integrador: un "Vision Board" interactivo que permita al usuario elegir imágenes, guardar frases motivacionales y almacenarlas en una base de datos y/o archivo, mostrando todo en un formulario VB6.

## Estructura del Módulo de Aprendizaje

- **Duración:** 2 semanas
- **Número de clases:** 4 sesiones de 2 horas cada una (total 8 horas)
- **Formato:** Clases teórico-prácticas con ejercicios en clase y tareas entre sesiones.
- **Proyecto Final:** "Vision Board" entregable al finalizar el módulo.

![Proyecto final](images\proyecto.png)



---

### Clase 1 (2 horas): Introducción al Entorno y Controles Básicos

**Contenidos:**
- Introducción a Visual Basic 6.0 y su IDE: ventanas, explorador de proyectos, propiedades, formulario, caja de herramientas.
- Creación de un proyecto estándar (EXE).
- Uso de controles básicos: `Label`, `TextBox`, `CommandButton`, `PictureBox`, `Frame`, `CheckBox`, `OptionButton`, `ListBox`, `ComboBox`.
- Propiedades y eventos más comunes (`Caption`, `Text`, `Enabled`, `Visible`, `Click`, `Change`).

**Ejemplo Práctico:**
1. Crear un nuevo proyecto EXE.
2. Colocar un `Label` con el texto “Bienvenido a mi aplicación”.
3. Agregar un `TextBox` para que el usuario ingrese su nombre.
4. Agregar un `CommandButton` que al hacer clic muestre un `MsgBox` con el mensaje “Hola, [Nombre]”.
5. Agregar un `PictureBox` y cargar una imagen estática desde las propiedades (`Picture`).

**Tarea Entre Sesiones:**
- Modificar el formulario para que el usuario pueda seleccionar entre dos imágenes (por ejemplo, mediante `OptionButtons`) y se muestre la elegida en el `PictureBox` al presionar un botón.

---

### Clase 2 (2 horas): Estructuras de Control, Manejo de Errores y Controles Avanzados

**Contenidos:**
- Repaso de variables y tipos de datos.
- Estructuras de control: `If...Then...Else`, `Select Case`, bucles `For...Next`, `Do...Loop`.
- Manejo de errores: `On Error GoTo`, `On Error Resume Next`, uso de `Err.Number` y `Err.Description`.
- Controles adicionales: `Common Dialog` (para abrir/guardar archivos, elegir colores, fuentes), `MSFlexGrid` (para mostrar datos en forma tabular, opcional).

**Ejemplo Práctico:**
1. Crear una función que valide el contenido de un `TextBox` (por ejemplo, verificar que no esté vacío) y que use `If...Then` para mostrar un mensaje de error personalizado.
2. Implementar un bloque `On Error GoTo` para controlar una posible división por cero. Ejemplo:
   ```vb
   On Error GoTo ManejoDeError
   Dim x As Integer, y As Integer
   x = 10
   y = 0 ' Probar con cero
   Debug.Print x / y
   Exit Sub
    ManejoDeError:
        MsgBox "Ocurrió un error: " & Err.Description

3. Mostrar con un `Select Case` diferentes mensajes según el día de la semana que ingrese el usuario en un `TextBox`.

**Tarea Entre Sesiones:**
- Implementar un pequeño formulario que solicite un número y muestre su tabla de multiplicar usando un bucle `For…Next`. 
- Incluir manejo de errores si el usuario ingresa datos no numéricos.

---

### Clase 3 (2 horas): Manejo de Archivos y Conexión a Base de Datos (SQL)

**Contenidos:**
- Lectura y escritura de archivos de texto: uso de `Open`, `Input #`, `Print #`, `Close`.
- Conexión a base de datos Access o SQL Server mediante ADO:
  - Referenciar Microsoft ActiveX Data Objects.
  - Crear un objeto `Connection` y un objeto `Recordset`.
  - Ejecutar consultas `SELECT`, `INSERT`, `UPDATE`.
- Mostrar datos obtenidos de la base en un `ListBox` o en un `MSFlexGrid`.

**Ejemplo Práctico con Archivos:**
1. Crear un botón “Guardar Frase” que tome el texto de un `TextBox` y lo guarde en un archivo de texto (ej: `frases.txt`):
   ```vb
   Dim MiFichero As Integer
   MiFichero = FreeFile
   Open "C:\MiProyecto\frases.txt" For Append As #MiFichero
   Print #MiFichero, TextBox1.Text
   Close #MiFichero
   MsgBox "Frase guardada."

**Ejemplo Práctico con Base de Datos:**

1. Conectar a una base Access llamada `VisionBoard.mdb`:

```vb
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MiProyecto\VisionBoard.mdb;"
cn.Open

Set rs = New ADODB.Recordset
rs.Open "SELECT * FROM Imagenes", cn, adOpenStatic, adLockOptimistic
While Not rs.EOF
    ListBox1.AddItem rs!NombreDeLaImagen
    rs.MoveNext
Wend
rs.Close
cn.Close
```

2. Mostrar los registros en un `ListBox` y permitir seleccionar uno para cargar una imagen en un `PictureBox`.

**Tarea Entre Sesiones:**

- Añadir una función que cargue frases desde el archivo `frases.txt` y las muestre en un `ListBox`.
- Crear una tabla en la base de datos para almacenar las rutas de imágenes motivacionales y sus títulos.

## Clase 4 (2 horas): Integración de Todo en el Proyecto Final: Vision Board

**Contenidos:**
 - Integración de controles, manejo de archivos y conexión a la base de datos.
 - Diseñar el formulario final: un “Vision Board” donde el usuario pueda:
   - Seleccionar una imagen motivacional desde una lista cargada desde la base de datos.
   - Agregar sus propias frases motivacionales y guardarlas en un archivo.
   - Visualizar en el formulario imágenes y frases.
   - Implementar una pequeña interfaz que permita refrescar el contenido.
- Finalización del proyecto, empaquetado (opcional) y documentación básica.

