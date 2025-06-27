import psutil

def get_cpu_usage():
    """Return CPU usage percentage."""
    return psutil.cpu_percent(interval=1)

def get_memory_usage():
    """Return Memory usage percentage."""
    mem = psutil.virtual_memory()
    return mem.percent

def get_disk_usage():
    """Return Disk usage percentage."""
    disk = psutil.disk_usage('/')
    return disk.percent

def get_battery_status():
    """Return battery percentage and charging status if available."""
    battery = psutil.sensors_battery()
    if battery is None:
        return "Battery info not available."
    return f"{battery.percent}% {'(Charging)' if battery.power_plugged else '(Not Charging)'}"

def get_system_info():
    """Combine all system info into one formatted string."""
    cpu = get_cpu_usage()
    memory = get_memory_usage()
    disk = get_disk_usage()
    battery = get_battery_status()
    info = (
        f"CPU Usage: {cpu}%\n"
        f"Memory Usage: {memory}%\n"
        f"Disk Usage: {disk}%\n"
        f"Battery: {battery}"
    )
    return info
