# Actualizador Automático de archivos Excel en OneDrive

Script en Python que abre, refresca y guarda automáticamente archivos de Excel con Power Query que están sincronizados en OneDrive/SharePoint.  
Evita los típicos problemas de procesos colgados de Excel, archivos no guardados y sincronización fallida con OneDrive.

Ideal para mantener dashboards y reportes siempre actualizados sin intervención manual todos los días.

## Archivos que actualiza automáticamente (cambialos por los nombres exacto de tus archivos)

1. **Archivo 1.xlsx**
2. **Archivo 2.xlsx** 
3. **Archivo 3.xlsx**

(Puedes agregar o quitar archivos fácilmente modificando la lista)

## Características principales
- Mata procesos antiguos de Excel antes de empezar (nada de EXCEL.EXE colgados)
- Espera a que OneDrive termine de sincronizar al abrir el archivo
- RefreshAll + espera real a que terminen todas las consultas Power Query
- Sistema de guardado con reintentos hasta que Excel confirme que realmente guardó
- Espera extra para que OneDrive suba los cambios antes de cerrar
- 100% automático y muy robusto (probado en entornos reales con OneDrive empresarial)

## Requisitos
- Windows 10/11
- Python 3.7 o superior
- OneDrive sincronizado y funcionando
- Los archivos Excel deben estar en una carpeta sincronizada de OneDrive/SharePoint

### Instalación de la única dependencia
```bash
pip install pywin32
