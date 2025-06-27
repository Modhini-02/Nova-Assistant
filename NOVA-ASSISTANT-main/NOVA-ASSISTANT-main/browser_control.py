import pyautogui
import time
import webbrowser

def search_google(command):
    try:
        # Extract the query from the command (e.g., "search cats")
        query = command.replace("search", "").strip()
        # Coordinates for the browser's address/search bar.
        SEARCH_BOX_COORDINATES = (735, 565)  # Update these coordinates as needed.
        pyautogui.click(SEARCH_BOX_COORDINATES[0], SEARCH_BOX_COORDINATES[1])
        time.sleep(0.5)
        pyautogui.write(query, interval=0.05)
        pyautogui.press("enter")
        return f"Searched Google for: {query}."
    except Exception as e:
        return f"Error: {str(e)}"

def click_nth_link(n):
    # Coordinates for the first link on the search results page.
    BASE_LINK_COORDINATES = (307, 392)  # Update these coordinates as needed.
    VERTICAL_OFFSET = 50  # Distance between links; adjust as needed.
    pyautogui.click(BASE_LINK_COORDINATES[0], BASE_LINK_COORDINATES[1] + (n - 1) * VERTICAL_OFFSET)
    return f"Clicked link number {n}."
