import psutil
import platform
import socket
import os
import wmi
from datetime import datetime
from screeninfo import get_monitors
import winreg

def get_system_info():
    info = {
        "Número de núcleos": psutil.cpu_count(logical=False),
        "Número de subprocesos": psutil.cpu_count(logical=True),
        "Frecuencia básica del procesador (GHz)": psutil.cpu_freq().max / 1000,
    }
    print("Información del sistema obtenida correctamente.")  # Mensaje de depuración
    return info

def get_cpu_cache_info():
    c = wmi.WMI()
    cache_info = {}
    
    # Obtener L2 y L3 cache usando WMI
    for cpu in c.Win32_Processor():
        cache_info['L1 Cache Size'] = "N/A" 
        cache_info['L2 Cache Size'] = round(cpu.L2CacheSize / 1024, 2)  # Convertir de KB a MB
        cache_info['L3 Cache Size'] = round(cpu.L3CacheSize / 1024, 2)  # Convertir de KB a MB
    print("Información de la caché del CPU obtenida correctamente.")  # Mensaje de depuración

    return cache_info

def format_bios_date(bios_date_str): 
    try:
        bios_date_str = bios_date_str.split('.')[0][:8]
        return datetime.strptime(bios_date_str, '%Y%m%d').strftime('%d/%m/%Y')
    except ValueError:
        return "Fecha desconocida"

def get_bios_info():
    c = wmi.WMI()
    bios_info = {}
    
    for bios in c.Win32_BIOS():
        bios_info['Versión BIOS'] = bios.SMBIOSBIOSVersion
        bios_info['Fecha BIOS'] = format_bios_date(bios.ReleaseDate)
        bios_info['Fabricante BIOS'] = bios.Manufacturer
    print("Información del BIOS obtenida correctamente.")  # Mensaje de depuración

    return bios_info

def get_physical_memory():
    c = wmi.WMI()
    total_memory = 0
    for memory in c.Win32_PhysicalMemory():
        total_memory += int(memory.Capacity)
    total_physical_memory_gb = round(total_memory / (1024 ** 3), 2)  # Convertir de bytes a GB
    return total_physical_memory_gb

def get_disk_info():
    c = wmi.WMI()
    disk_info = []
    
    for disk in c.Win32_DiskDrive():
        disk_size_gb = round(float(disk.Size) / (1024 ** 3), 2)  # Convertir de bytes a GB
        disk_type = "SSD" if "SSD" in disk.Model else "HDD"  # Determinar si es SSD o HDD basado en el modelo
        disk_info.append({
            "Size (GB)": disk_size_gb,
            "Type": disk_type
        })
    return disk_info

def get_screen_info():
    screen_info = []
    estimated_dpi = 96  # Estimación del DPI si no está disponible
    
    for monitor in get_monitors():
        if monitor.width_mm and monitor.height_mm:  # Si los valores en mm están disponibles
            width_inches = monitor.width_mm / 25.4
            height_inches = monitor.height_mm / 25.4
            diagonal_size = round((width_inches**2 + height_inches**2) ** 0.5, 2)
        else:  # Si los valores en mm no están disponibles, usar DPI estimado
            diagonal_size = round(((monitor.width ** 2 + monitor.height ** 2) ** 0.5) / estimated_dpi, 2)
        
        screen_info.append({
            "Width": monitor.width,
            "Height": monitor.height,
            "Name": monitor.name,
            "Size (inches)": diagonal_size
        })
        
    return screen_info

def get_antivirus_info():
    c = wmi.WMI(namespace="root\\SecurityCenter2")
    antivirus_info = []
    
    for antivirus in c.AntivirusProduct():
        antivirus_info.append({
            "Name": antivirus.displayName,
            "Enabled": "Sí" if antivirus.productState & 0x10000 else "No"
        })
    
    return antivirus_info

import winreg
import wmi

def get_office_version_and_activation():
    office_info = []
    try:
        # Intentar obtener la versión de Office desde el registro
        access_registry = winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE)
        access_key = winreg.OpenKey(access_registry, r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration")
        product_version, reg_type = winreg.QueryValueEx(access_key, "VersionToReport")
        office_info.append({
            "Name": "Microsoft Office",
            "Version": product_version,
            "Activated": "Desconocido"
        })
        winreg.CloseKey(access_key)
    except FileNotFoundError:
        office_info.append({"Name": "Microsoft Office", "Version": "No encontrado", "Activated": "No encontrado"})
    except Exception as e:
        office_info.append({"Name": "Microsoft Office", "Version": f"Error: {e}", "Activated": "Desconocido"})

    try:
        # Usar WMI para verificar el estado de activación
        c = wmi.WMI(namespace="root\\cimv2")
        for product in c.SoftwareLicensingProduct(Description="Office 16, Office16ProPlusVL_KMS_Client edition"):
            if product.LicenseStatus == 1:
                office_info[-1]["Activated"] = "Sí"
            else:
                office_info[-1]["Activated"] = "No"
    except Exception as e:
        office_info[-1]["Activated"] = f"Error: {e}"

    return office_info

def save_system_info_to_file(info, cache_info, bios_info, physical_memory, disk_info, screen_info, antivirus_info, office_info, filename):
    with open(filename, 'w') as f:
        f.write("INFORMACION DE TDR ATSI 2024\n\n")
        f.write(f"Número de núcleos: {info['Número de núcleos']}\n")
        f.write(f"Número de subprocesos: {info['Número de subprocesos']}\n")
        f.write(f"Frecuencia básica del procesador (GHz): {info['Frecuencia básica del procesador (GHz)']:.2f}\n")
        f.write(f"L1 Cache : {cache_info['L1 Cache Size']}\n")
        f.write(f"L2 Cache : {cache_info['L2 Cache Size']} MB\n")
        f.write(f"L3 Cache : {cache_info['L3 Cache Size']} MB\n\n")
        f.write(f"Versión del BIOS: {bios_info['Versión BIOS']}\n")
        f.write(f"Fecha del BIOS: {bios_info['Fecha BIOS']}\n")
        f.write(f"Fabricante del BIOS: {bios_info['Fabricante BIOS']}\n\n")
        f.write(f"\nMemoria Física Instalada (RAM): {physical_memory} GB\n\n")
        f.write("Información del Disco Duro:\n")
        for disk in disk_info:
            f.write(f"Capacidad Total: {disk['Size (GB)']} GB\n")
            f.write(f"Tipo de Disco: {disk['Type']}\n\n")
        for screen in screen_info:
            f.write(f"Resolución: {screen['Width']}x{screen['Height']}\n")
            f.write(f"Tamaño (pulgadas): {screen['Size (inches)']}\n\n")
        for antivirus in antivirus_info:
            f.write(f"Nombre del Antivirus: {antivirus['Name']}\n")
            f.write(f"Antivirus Activado: {antivirus['Enabled']}\n\n")
        for office in office_info:
            f.write(f"Nombre de Office: {office['Name']}\n")
            f.write(f"Versión de Office: {office['Version']}\n\n\n\n")
        f.write("CREADO POR RAOD\n")
    print(f"Información guardada en {filename}.")  

def main():
    print("Inicio del programa.")  
    try:
        info = get_system_info()
        print("Información del sistema obtenida.")  
        cache_info = get_cpu_cache_info()
        print("Información del caché obtenida.")  
        bios_info = get_bios_info()
        print("Información del BIOS obtenida.")  
        physical_memory = get_physical_memory()
        print("Información de la memoria física obtenida.")  
        disk_info = get_disk_info()
        print("Información del disco obtenida.")  
        screen_info = get_screen_info()
        print("Información de la pantalla obtenida.")  
        antivirus_info = get_antivirus_info()
        print("Información del antivirus obtenida.")  
        office_info = get_office_version_and_activation()
        print("Información de Office obtenida.")  

        filename = "system_info.txt"
        save_system_info_to_file(info, cache_info, bios_info, physical_memory, disk_info, screen_info, antivirus_info, office_info, filename)
        print("Información guardada en el archivo.")  

        # Verificar si el archivo existe y luego abrirlo
        if os.path.exists(filename):
            print("El archivo existe. Intentando abrir el Bloc de Notas...")
            os.system(f'notepad.exe {filename}')
        else:
            print("Error: el archivo no fue creado.")
    except Exception as e:
        print(f"Se produjo un error: {e}")

if __name__ == "__main__":
    main()

