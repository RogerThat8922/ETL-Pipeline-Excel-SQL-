import pyautogui
import time

# Lowered fail-safe sensitivity for standard users
pyautogui.FAILSAFE = True

print("Anti-AFK (No-Admin version) started.")
print("Mode: Toggle Scroll Lock every 60s.")
print("To stop: Move mouse to any corner or press Ctrl+C.")

try:
    while True:
        # Press Scroll Lock twice (ON then OFF)
        # This triggers a hardware-level event without moving your mouse
        pyautogui.press('scrolllock')
        time.sleep(0.2)
        pyautogui.press('scrolllock')
        
        # Log the activity
        current_time = time.strftime("%H:%M:%S")
        print(f"[{current_time}] System pinged (Scroll Lock toggled).")
        
        # Wait 60 seconds
        time.sleep(60)

except pyautogui.FailSafeException:
    print("\n[Stopped] Mouse moved to corner.")
except KeyboardInterrupt:
    print("\n[Stopped] Manual exit.")
