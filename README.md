# Censo Automatico 6.0

Guia escrita para medicos que no programan.

Este programa toma informacion desde la ficha clinica / sistema de gestion clinica del hospital y actualiza automaticamente un Google Sheet usado como censo de Neurocirugia. La idea es evitar copiar paciente por paciente a mano.

## Que hace, en castellano simple

`CENSO AUTOMATICO 6.0.py` hace esto:

1. Abre sesion en el sistema clinico del hospital con usuario y clave.
2. Entra a varios sectores, por ejemplo Neurocirugia, UCI, UTI, Pediatria, Urgencia y TMT.
3. Descarga desde cada sector una planilla Excel con los pacientes actuales.
4. Lee el Google Sheet del censo.
5. Busca las primeras 4 pestanas del Google Sheet:
   - `S. DE NEUROCIRUGIA`
   - `UCI/UTI`
   - `PEDIATRICOS`
   - `OTROS SERVICIOS`
6. Compara las camas y pacientes de la ficha clinica contra las camas del Google Sheet.
7. Actualiza automaticamente:
   - nombre del paciente
   - diagnostico
   - RUT
8. Intenta conservar datos escritos a mano por el equipo:
   - `STAFF`
   - `BECADO`
   - `PLAN`
9. Si un paciente se mueve de cama o de unidad, intenta reconocerlo por RUT o por nombre para no perder esos datos manuales.

## Que NO hace

El programa no reemplaza el juicio medico.

No decide conducta, no interpreta diagnosticos, no confirma indicaciones y no valida si la informacion de la ficha esta clinicamente correcta. Solo copia y ordena datos administrativos desde una fuente hacia otra.

Siempre revise el censo despues de correrlo, especialmente si hubo cambios de cama, altas, ingresos multiples o pacientes con nombres parecidos.

## Archivos importantes

`CENSO AUTOMATICO 6.0.py`

El programa principal. Es el archivo que se ejecuta.

`.env`

Archivo privado de configuracion local. Aqui van usuario, clave y ruta al archivo de credenciales de Google. Este archivo no debe subirse a GitHub.

`.env.example`

Ejemplo de como debe verse el `.env`. Sirve para copiarlo en otro computador.

`censo-automatico-430916-f23cd66e6932.json`

Credencial de Google para que el programa pueda escribir en el Google Sheet. Es como una llave. No debe compartirse publicamente ni subirse a GitHub.

`requirements.txt`

Lista de librerias que Python necesita para correr el programa.

`.gitignore`

Lista de archivos que Git debe ignorar para no subir claves, archivos temporales, cache, logs ni la carpeta `venv`.

## Como se conecta con la ficha clinica

Dentro del codigo hay direcciones internas del hospital:

- pagina de login
- pagina de listado de pacientes por sector
- pagina que exporta el listado a Excel

El programa abre una sesion como si fuera un usuario entrando al sistema por navegador. Luego visita las paginas de los sectores configurados y descarga el Excel de cada uno.

Los sectores configurados por defecto son:

- Neurocirugia
- UCI segundo piso
- UTI segundo piso
- UCI/UTI tercer piso
- UCI/UTI pediatricos
- Urgencia, filtrando solo Neurocirugia
- TMT, filtrando solo Neurocirugia

La informacion descargada se transforma a un formato comun: cama, paciente, diagnostico y RUT. Despues se escribe en el Google Sheet.

## Como decide donde escribir cada paciente

El programa necesita que el Google Sheet tenga columnas reconocibles. Busca encabezados parecidos a:

- cama: `CAMA`, `SALA`, `HABITACION`, `BOX`
- paciente: `NOMBRE`, `PACIENTE`, `NOMBRE PACIENTE`
- diagnostico: `DIAGNOSTICO`, `DIAG`, `DX`
- RUT: `RUT`, `RUN`, `ID`, `IDENTIFICACION`
- campos manuales: `STAFF`, `BECADO`, `PLAN`

El programa busca la fila de encabezados dentro de las primeras filas de cada pestana. Luego crea un mapa de camas. Por ejemplo: si la ficha dice cama 15 y el Google Sheet tiene una fila para cama 15, actualiza esa fila.

## Campos que el programa cuida

El programa considera automaticos estos campos:

- paciente
- diagnostico
- RUT

Y considera manuales estos campos:

- staff
- becado
- plan

Eso significa que, si reconoce al mismo paciente, intenta mantener el staff, becado y plan aunque el paciente se haya movido.

## Como correr el programa paso a paso

### 1. Abra Visual Studio Code

Visual Studio Code, o VS Code, es un editor de texto para trabajar con codigo. Piense en el como Word, pero para programas.

No es la ficha clinica. No es Python. No es GitHub. Es la ventana donde usted puede ver y editar los archivos del proyecto.

### 2. Abra la carpeta correcta

En VS Code:

1. Vaya a `File`.
2. Elija `Open Folder`.
3. Seleccione la carpeta `CENSO AUTOMATICO`.

Debe ver el archivo `CENSO AUTOMATICO 6.0.py` en la barra lateral.

### 3. Revise que exista el archivo `.env`

Debe existir un archivo llamado `.env`. Si no existe:

1. Copie `.env.example`.
2. Pegue una copia en la misma carpeta.
3. Cambie el nombre de la copia a `.env`.

El `.env` debe tener valores como estos:

```text
CENSO_SERVICE_ACCOUNT_JSON=./censo-automatico-430916-f23cd66e6932.json
CENSO_USERNAME=su_usuario
CENSO_PASSWORD=su_clave
```

No suba el `.env` a GitHub.

### 4. Abra la terminal

En VS Code:

1. Vaya a `Terminal`.
2. Elija `New Terminal`.

Abajo aparecera una ventana negra o blanca donde se escriben comandos.

### 5. Ejecute el programa

Escriba:

```bash
python3 "CENSO AUTOMATICO 6.0.py"
```

Si su computador usa `python` en vez de `python3`, pruebe:

```bash
python "CENSO AUTOMATICO 6.0.py"
```

### 6. Espere los mensajes

El programa mostrara mensajes de avance. Si termina bien, dira algo parecido a:

```text
Census update finished. No sheets outside the first 4 tabs were written.
```

### 7. Revise el Google Sheet

Despues de correr el programa, revise manualmente:

- pacientes nuevos
- pacientes trasladados
- camas vacias
- pacientes con nombres parecidos
- campos `STAFF`, `BECADO` y `PLAN`

## Errores frecuentes

### Faltan librerias

El programa intenta instalar librerias automaticamente usando `requirements.txt`. Si falla, pida a Codex o Claude:

```text
Estoy intentando correr CENSO AUTOMATICO 6.0.py y falla al instalar librerias.
Lee el error de la terminal y arregla el entorno sin cambiar la logica clinica del programa.
```

### No encuentra credenciales de Google

Revise:

- que exista el archivo `.env`
- que exista el archivo JSON de Google
- que la linea `CENSO_SERVICE_ACCOUNT_JSON` apunte al archivo correcto

Prompt sugerido:

```text
El programa no encuentra las credenciales de Google.
Revisa .env, .env.example y la funcion resolve_credentials_path.
Explicame el problema en lenguaje simple y arreglalo sin subir secretos a GitHub.
```

### Usuario o clave incorrectos

Revise el `.env`:

```text
CENSO_USERNAME=...
CENSO_PASSWORD=...
```

Prompt sugerido:

```text
El login a la ficha clinica esta fallando.
Revisa la funcion login_to_hospital_system y dime si el problema parece ser usuario/clave, red, token CSRF o cambio de pagina.
No muestres mi clave en la respuesta.
```

### Cambio la estructura del Google Sheet

Si alguien cambia nombres de columnas, elimina columnas o mueve las primeras pestanas, el programa puede fallar.

Prompt sugerido:

```text
El Google Sheet del censo cambio.
Quiero que adaptes CENSO AUTOMATICO 6.0.py para reconocer las nuevas columnas.
Primero explicame que columnas espera el codigo actualmente y luego modifica solo lo necesario.
```

### Cambio la pagina de la ficha clinica

Si el hospital cambia la pagina interna, puede fallar la descarga.

Prompt sugerido:

```text
La ficha clinica cambio y CENSO AUTOMATICO 6.0.py ya no descarga los Excel.
Revisa las funciones login_to_hospital_system y download_sector_dataframe.
Explicame que parte probablemente se rompio y propon una correccion minima.
```

## Que es GitHub

GitHub es una pagina web donde se guarda una copia del codigo.

Sirve para:

- tener respaldo
- ver versiones anteriores
- compartir el programa con otros computadores
- pedir ayuda a IA o a programadores sin mandar archivos sueltos por WhatsApp

Pero GitHub no debe recibir claves, archivos `.env`, credenciales JSON ni datos de pacientes.

## Que es Git

Git es el sistema que guarda versiones del codigo.

Piense en Git como un historial de cambios. Permite decir:

- antes funcionaba asi
- ahora cambio esto
- quien hizo el cambio
- cuando se hizo

No necesita dominar Git para usar el programa, pero es util saber que existe.

## Que es VS Code o VSC

Mucha gente le dice VSC, pero el nombre correcto suele ser VS Code.

VS Code es el programa donde abre la carpeta, ve los archivos y conversa con herramientas como Codex o Claude Code.

En VS Code usted puede:

- abrir `CENSO AUTOMATICO 6.0.py`
- editar el codigo
- abrir una terminal
- correr el programa
- ver errores
- trabajar con IA dentro del proyecto

## Que es Codex o Claude Code

Codex y Claude Code son asistentes de IA que pueden leer el codigo, entender errores y proponer modificaciones.

Lo importante: no hay que hablarles como programador. Hay que hablarles como medico que describe un flujo de trabajo.

Ejemplo malo:

```text
Arregla el codigo.
```

Ejemplo bueno:

```text
El programa deberia mantener el PLAN cuando un paciente se traslada desde UCI a Neurocirugia, pero hoy lo pierde.
Lee el codigo, dime donde ocurre eso y modifica lo minimo para conservar PLAN.
Despues prueba que el archivo siga compilando.
```

## Que es vibe coding

Vibe coding significa trabajar con una IA describiendo lo que usted quiere lograr, sin escribir todo el codigo a mano.

La parte importante no es decir "haz magia". La parte importante es guiar a la IA como si fuera un interno inteligente pero nuevo en el servicio:

- explicarle el objetivo clinico
- mostrarle el error exacto
- pedir cambios pequenos
- pedir que no toque lo que ya funciona
- pedir que pruebe antes de terminar
- revisar el resultado

Usted pone el criterio clinico. La IA ayuda con el codigo.

## Regla de oro para trabajar con IA

Pida cambios chicos.

No diga:

```text
Rehaz todo el censo automatico.
```

Diga:

```text
Quiero agregar una nueva unidad llamada X.
Debe descargar pacientes desde el sector Y.
Debe escribirlos en la pestana Z.
No cambies la logica de preservacion de STAFF, BECADO y PLAN.
```

## Como pedirle a Codex o Claude que modifique algo

Use esta estructura:

```text
Contexto:
Estoy trabajando en CENSO AUTOMATICO 6.0.py.
El programa actualiza un Google Sheet de censo desde la ficha clinica.

Objetivo:
Quiero cambiar [explique el cambio].

Restricciones:
No subas claves ni datos de pacientes a GitHub.
No cambies la logica que conserva STAFF, BECADO y PLAN, salvo que sea necesario.
Haz el cambio mas pequeno posible.

Verificacion:
Despues de cambiarlo, corre python3 -m py_compile "CENSO AUTOMATICO 6.0.py".
Explicame en palabras simples que cambiaste.
```

## Ejemplos de prompts utiles

### Agregar una nueva unidad

```text
Quiero agregar una nueva unidad al censo automatico.
La unidad se llama [nombre].
En la ficha clinica corresponde al sector_id [numero].
Debe escribirse en la pestana [nombre exacto de la pestana].
Lee como se define DEFAULT_SOURCE_SECTORS y agrega esta unidad siguiendo el estilo actual.
No cambies otras unidades.
Prueba que el archivo compile.
```

### Cambiar el nombre de una pestana

```text
Cambio el nombre de una pestana del Google Sheet.
Antes se llamaba [nombre antiguo] y ahora se llama [nombre nuevo].
Actualiza CENSO AUTOMATICO 6.0.py para usar el nuevo nombre.
Revisa si tambien hay filtros o sectores que dependan de ese nombre.
```

### Cambiar columnas reconocidas

```text
El Google Sheet ahora usa la columna [nuevo nombre] para [cama/paciente/diagnostico/RUT].
Actualiza HEADER_ALIASES para reconocer ese encabezado.
No modifiques la logica de descarga ni login.
```

### Evitar borrar camas vacias

```text
Quiero confirmar si el programa borra pacientes de camas que ya no aparecen en la ficha.
Lee la configuracion CLEAR_MISSING_BEDS y explicamela.
Si esta activada, dejala desactivada por defecto.
```

### Revisar un error de terminal

```text
Te pego el error que aparece al correr el programa.
No hagas cambios todavia.
Primero explicame en lenguaje medico/no programador que significa, cual es la causa mas probable y que archivo habria que tocar.

[pegue aqui el error completo]
```

### Pedir una reparacion con cuidado

```text
El programa falla con este error:

[pegue aqui el error]

Arreglalo con el cambio mas pequeno posible.
No cambies nombres de columnas ni unidades si no es necesario.
No subas .env ni credenciales.
Al final dime exactamente que archivo cambiaste y como probarlo.
```

## Como revisar cambios antes de aceptarlos

Pida siempre:

```text
Antes de terminar, resumeme:
1. Que cambiaste.
2. Por que lo cambiaste.
3. Que riesgo queda.
4. Como lo pruebo.
```

Si algo no le gusta, diga:

```text
No quiero ese enfoque.
Revierte solo tu ultimo cambio y propon una alternativa mas pequena.
```

## Donde tocar el codigo segun lo que quiera cambiar

### Cambiar usuario, clave o credenciales

No toque el codigo. Cambie `.env`.

### Cambiar el Google Sheet

Busque:

```python
DEFAULT_SPREADSHEET_URL
```

Mejor opcion: poner `SPREADSHEET_URL` en `.env`.

### Cambiar sectores descargados desde la ficha

Busque:

```python
DEFAULT_SOURCE_SECTORS
```

Cada sector tiene:

- `sector_id`: numero interno del sector en la ficha
- `target_sheet`: pestana del Google Sheet donde escribira
- `label`: nombre humano para los logs
- `filter_column`: columna usada para filtrar, si corresponde
- `filter_value`: valor esperado para filtrar, si corresponde

### Cambiar nombres de columnas reconocidas

Busque:

```python
HEADER_ALIASES
```

Aqui se agregan sinonimos. Por ejemplo, si una columna ahora se llama `RUN PACIENTE`, se puede agregar a la lista de RUT.

### Cambiar que campos se preservan

Busque:

```python
PRESERVED_FIELD_KEYS
```

Actualmente conserva:

```python
["staff", "becado", "plan"]
```

No cambie esto sin probar muy bien, porque afecta datos escritos manualmente.

## Seguridad y privacidad

Este proyecto puede conectarse a sistemas clinicos y manejar informacion sensible.

Nunca suba a GitHub:

- `.env`
- claves
- credenciales JSON de Google
- planillas con datos de pacientes
- logs con datos identificables
- archivos Excel o CSV exportados desde la ficha

Antes de pedir ayuda externa, borre o anonimice:

- nombres
- RUT
- fechas de nacimiento
- numeros de ficha
- diagnosticos identificables raros

## Checklist antes de correrlo en serio

1. Estoy dentro de la red/VPN necesaria para acceder a la ficha.
2. El `.env` tiene usuario y clave correctos.
3. El JSON de Google existe.
4. El Google Sheet tiene las primeras 4 pestanas esperadas.
5. Las columnas principales existen.
6. Nadie cambio manualmente encabezados importantes.
7. Tengo claro que debo revisar el censo despues de actualizar.

## Checklist despues de modificar codigo con IA

1. La IA explico que cambio.
2. El cambio fue pequeno y entendible.
3. No se subieron claves ni datos de pacientes.
4. Se corrio:

```bash
python3 -m py_compile "CENSO AUTOMATICO 6.0.py"
```

5. Se hizo una prueba controlada antes de confiar en el resultado.
6. Se reviso el Google Sheet despues de correrlo.

## Frase util para empezar una sesion con Codex o Claude

```text
Soy medico y no programador. Este repositorio contiene CENSO AUTOMATICO 6.0.py, que actualiza un Google Sheet de censo desde la ficha clinica.
Quiero que me expliques los cambios en lenguaje simple, que no subas secretos ni datos de pacientes, y que hagas modificaciones pequenas y verificables.
Antes de editar, lee el archivo y dime que parte vas a tocar.
```

## Frase util cuando algo falla

```text
El programa fallo. Te voy a pegar el error completo.
Quiero que lo leas, me expliques la causa probable en lenguaje simple y luego lo arregles con el cambio minimo.
No cambies el flujo clinico salvo que sea estrictamente necesario.
```

## Resumen final

El Censo Automatico 6.0 es un puente entre la ficha clinica y el Google Sheet del equipo. Descarga pacientes desde sectores definidos, identifica camas y pacientes, actualiza datos automaticos y trata de proteger los campos manuales importantes.

VS Code es el lugar donde se abre y edita el proyecto. GitHub es el respaldo del codigo. Codex o Claude Code son asistentes para modificarlo, siempre que usted les de instrucciones clinicas claras, cambios pequenos y reglas de seguridad.
