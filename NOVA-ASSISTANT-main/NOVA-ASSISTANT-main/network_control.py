import os
import subprocess
import speedtest

def open_network_settings():
    """
    Open Windows network settings.
    """
    os.system("start ms-settings:network")
    return "Opened Network Settings."

def run_speed_test():
    """
    Run a speed test and return download and upload speeds.
    """
    try:
        st = speedtest.Speedtest()
        st.get_best_server()
        download_speed = st.download() / 1_000_000  # Mbps
        upload_speed = st.upload() / 1_000_000      # Mbps
        return f"Download: {download_speed:.2f} Mbps, Upload: {upload_speed:.2f} Mbps"
    except Exception as e:
        return f"Speed test error: {e}"

def ping_test(host="8.8.8.8", count=4):
    """
    Ping a host and return the output.
    """
    try:
        output = subprocess.check_output(["ping", host, "-n", str(count)], universal_newlines=True)
        return output
    except Exception as e:
        return f"Ping error: {e}"
