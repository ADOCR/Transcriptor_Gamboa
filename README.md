# Gamboa Transcriptor

**Gamboa Desarrollos - Transcriptor Local y Minutador** es una aplicacion de escritorio para Windows que transcribe audios y videos en español usando modelos locales con `faster-whisper` y, opcionalmente, genera minutas/resumenes con Ollama local.

La transcripcion y la generacion de minutas se realizan localmente en el equipo del usuario.

## Funciones principales

- Transcripcion local de audio y video.
- Uso de GPU CUDA cuando esta disponible.
- Respaldo automatico en CPU/int8 si CUDA falla.
- Generacion opcional de minuta con Ollama local.
- Procesamiento de archivo individual o carpeta completa.
- Modo separado para generar minuta desde TXT existente.
- Exportacion a TXT con tiempos, TXT limpio, SRT, DOCX de transcripcion, TXT de minuta y DOCX de minuta.
- Interfaz grafica en PySide6 con progreso, registro y reporte final.
- Preparado para generar `.exe` con PyInstaller e instalador con Inno Setup.

## Requisitos

- Windows 10/11.
- Python 3.11 recomendado.
- 32 GB RAM recomendado para lotes grandes.
- GPU NVIDIA compatible con CUDA para usar `large-v3` en GPU.
- Ollama instalado si se quieren generar minutas.

Hardware objetivo probado/diseñado:

- AMD Ryzen 5 4600G.
- 32 GB RAM.
- NVIDIA RTX 3070 8 GB VRAM.

## Instalacion recomendada con Anaconda

No instale PySide6 en el entorno `base` de Anaconda. En algunos equipos se mezclan DLLs de Qt/ICU y aparece un error como:

```text
ImportError: DLL load failed while importing QtCore
```

Cree un entorno separado:

```bat
conda env create -f environment.yml
conda activate gamboa-transcriptor
python main.py
```

Tambien puede usar:

```bat
build_scripts\setup_conda_env.bat
```

## Instalacion manual para desarrollo

Crear entorno limpio:

```bat
conda create -n gamboa-transcriptor python=3.11 -y
conda activate gamboa-transcriptor
```

Instalar PySide6 con conda-forge para evitar conflictos de DLLs:

```bat
conda install -c conda-forge "pyside6>=6.7,<6.10" -y
```

Instalar el resto:

```bat
python -m pip install faster-whisper python-docx requests pyinstaller
```

Ejecutar:

```bat
python main.py
```

## Ollama

Para generar minutas:

1. Instale Ollama desde https://ollama.com
2. Descargue el modelo:

```bat
ollama pull qwen3:8b
```

La aplicacion verifica:

- Si Ollama esta activo.
- Si `qwen3:8b` aparece instalado.
- Si Ollama no esta disponible, la app permite transcribir de todos modos.

## FFmpeg

`faster-whisper` usa PyAV y normalmente puede abrir muchos formatos. Si algun archivo multimedia no abre correctamente, instale FFmpeg:

```bat
conda install -c conda-forge ffmpeg
```

## Uso

1. Abra la aplicacion.
2. Elija un modo:
   - Solo transcripcion.
   - Transcripcion + resumen/minuta.
   - Generar resumen/minuta desde TXT existente.
3. Seleccione archivo, carpeta o TXT.
4. Elija carpeta de salida automatica o una carpeta personalizada.
5. Ajuste modelo Whisper, dispositivo, salidas y opciones de Ollama.
6. Presione **INICIAR PROCESO**.

## Estructura de salida

Por cada archivo:

```text
transcripciones/
  NombreArchivo/
    NombreArchivo_con_tiempos.txt
    NombreArchivo_texto_limpio.txt
    NombreArchivo.srt
    NombreArchivo.docx
    NombreArchivo_MINUTA.txt
    NombreArchivo_MINUTA.docx
    NombreArchivo_log.txt
```

Si la opcion de sobrescritura esta desactivada, se crean nombres alternativos:

```text
NombreArchivo_MINUTA_2.docx
NombreArchivo_MINUTA_3.docx
```

## Recomendaciones RTX 3070 8 GB

- Use `large-v3` para maxima calidad.
- Use `medium` para menor consumo de VRAM y mayor velocidad.
- Use `num_ctx=4096` para mayor estabilidad con Ollama.
- Para muchos archivos o audios largos, use primero **Solo transcripcion** y despues **Minuta desde TXT**.
- La opcion **Liberar modelo Whisper antes de generar minutas** puede ayudar con VRAM, pero tambien puede ser mas delicada segun drivers, CUDA y librerias. Si la app se vuelve inestable, deje esa opcion desactivada.

## Compilar .exe con PyInstaller

Desde la carpeta del proyecto:

```bat
build_scripts\build_exe.bat
```

El ejecutable queda en:

```text
dist\GamboaTranscriptor\GamboaTranscriptor.exe
```

El build usa `--onedir` y `--noconsole` para mayor estabilidad con dependencias grandes.

## Crear instalador con Inno Setup

1. Instale Inno Setup.
2. Primero genere el `.exe` con:

```bat
build_scripts\build_exe.bat
```

3. Abra `installer.iss` o `build_scripts\build_installer.iss` con Inno Setup.
4. Compile el instalador.

El instalador resultante se llama:

```text
GamboaTranscriptor_Setup.exe
```

El instalador:

- Instala en Program Files.
- Crea acceso directo en menu inicio.
- Puede crear acceso directo en escritorio.
- Permite desinstalacion.
- No incluye modelos de Whisper ni modelos de Ollama.
- Muestra aviso final: "Para generar minutas, instale Ollama y descargue qwen3:8b."

## Modelos

Los modelos de Whisper y Ollama no se incluyen dentro del instalador.

- `faster-whisper` descargara/cacheara el modelo Whisper la primera vez que se use.
- Ollama requiere descargar el modelo con `ollama pull qwen3:8b`.

Luego de descargados, pueden funcionar localmente/offline.

## Solucion de problemas

### DLL load failed while importing QtCore

Este error suele aparecer cuando PySide6 fue instalado con `pip` dentro del entorno `base` de Anaconda y se mezclan DLLs de Qt/ICU.

Solucion recomendada:

```bat
conda deactivate
conda env create -f environment.yml
conda activate gamboa-transcriptor
python main.py
```

Si ya instalo PySide6 en `base`, no es necesario arreglar `base` para usar la app. Use el entorno dedicado. Luego, con calma, puede limpiar `base` desinstalando PySide6 desde ahi:

```bat
conda activate base
python -m pip uninstall -y PySide6 PySide6_Addons PySide6_Essentials shiboken6
```

### CUDA no disponible

La app mostrara:

```text
CUDA no esta disponible. Se usara CPU.
```

Revise:

- Driver NVIDIA actualizado.
- Paquetes CUDA/CTranslate2 compatibles.
- Que no haya otro proceso usando demasiada VRAM.

### Ollama no activo

La app mostrara:

```text
Ollama no esta activo. Puede transcribir, pero no generar minutas.
```

Abra Ollama o ejecute:

```bat
ollama serve
```

### Modelo qwen3:8b no instalado

Ejecute:

```bat
ollama pull qwen3:8b
```

### Un archivo multimedia no abre

Instale FFmpeg:

```bat
conda install -c conda-forge ffmpeg
```

### La app tarda mucho al abrir el .exe

El proyecto usa `--onedir` porque dependencias como `faster-whisper`, `ctranslate2`, `av` y PySide6 son grandes. Evite `--onefile` al inicio.

## Estructura del proyecto

```text
.
├─ main.py
├─ app_core.py
├─ requirements.txt
├─ README.md
├─ installer.iss
├─ assets/
│  ├─ icon.ico
│  └─ logo.png
└─ build_scripts/
   ├─ build_exe.bat
   └─ build_installer.iss
```

## Nota de privacidad

La transcripcion y la generacion de minutas se realizan localmente en el equipo del usuario. No se envia audio, video ni texto a servicios en la nube desde esta aplicacion.
