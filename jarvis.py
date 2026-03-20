import os, sys, time, threading, subprocess, webbrowser
import datetime, platform, shutil, json, socket, re
import winreg, ctypes, glob

try:
    import pyttsx3
    import speech_recognition as sr
    import psutil
    import pygetwindow as gw
    import pyautogui
    import pyperclip
except ImportError as e:
    print(f"[ERROR] Missing library: {e}")
    print("Run: pip install pyttsx3 SpeechRecognition pyaudio psutil pygetwindow pyautogui pyperclip")
    sys.exit(1)

pyautogui.FAILSAFE = True   # move mouse to top-left corner to abort

# ═══════════════════════════════════════════════════════════════
#  HELPER — run as admin check
# ═══════════════════════════════════════════════════════════════
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

# ═══════════════════════════════════════════════════════════════
#  JARVIS CLASS
# ═══════════════════════════════════════════════════════════════
class Jarvis:
    def __init__(self):
        self.wake_word = "jarvis"
        self.running   = True
        self.name      = "Jarvis"

        # TTS
        self.engine = pyttsx3.init()
        voices = self.engine.getProperty("voices")
        self.engine.setProperty("voice",  voices[0].id)
        self.engine.setProperty("rate",   172)
        self.engine.setProperty("volume", 1.0)

        # STT
        self.rec = sr.Recognizer()
        self.rec.pause_threshold  = 1.0
        self.rec.energy_threshold = 300

        admin_str = "with Administrator privileges" if is_admin() else "— NOTE: run as Administrator for full access"
        self.speak(f"J.A.R.V.I.S online {admin_str}. All systems ready, sir.")

    # ─────────────────────────────────────────────────────────
    #  VOICE I/O
    # ─────────────────────────────────────────────────────────
    def speak(self, text: str):
        print(f"\n[JARVIS] {text}")
        self.engine.say(text)
        self.engine.runAndWait()

    def _check_microphone(self) -> bool:
        """Check if any microphone is available."""
        try:
            mics = sr.Microphone.list_microphone_names()
            if not mics:
                print("[ERROR] No microphone found. Connect a mic and restart.")
                return False
            return True
        except Exception as e:
            print(f"[ERROR] Microphone check failed: {e}")
            return False

    def listen(self, timeout=7, phrase_limit=12) -> str:
        # ── Keyboard fallback if no mic ──────────────────────
        if not self._check_microphone():
            try:
                user_input = input("[TYPE CMD] > ").strip()
                return user_input.lower()
            except (EOFError, KeyboardInterrupt):
                return ""

        # ── Microphone mode ──────────────────────────────────
        try:
            mic = sr.Microphone()
            with mic as src:
                print("\n[MIC] Listening...")
                try:
                    self.rec.adjust_for_ambient_noise(src, duration=0.4)
                except AssertionError:
                    pass   # stream issue — skip noise adjustment
                try:
                    audio = self.rec.listen(src, timeout=timeout, phrase_time_limit=phrase_limit)
                    text  = self.rec.recognize_google(audio, language="en-IN")
                    print(f"[YOU] {text}")
                    return text.lower()
                except sr.WaitTimeoutError:   return ""
                except sr.UnknownValueError:  return ""
                except sr.RequestError:
                    self.speak("Speech service unavailable.")
                    return ""
        except (OSError, AttributeError) as e:
            print(f"[MIC ERROR] {e}")
            print("[FALLBACK] Microphone unavailable — switching to keyboard input.")
            print("[HINT] Fix: pip install pipwin && pipwin install pyaudio")
            try:
                user_input = input("[TYPE CMD] > ").strip()
                return user_input.lower()
            except (EOFError, KeyboardInterrupt):
                return ""

    # ─────────────────────────────────────────────────────────
    #  MASTER COMMAND ROUTER
    # ─────────────────────────────────────────────────────────
    def process(self, cmd: str):
        c = cmd.strip().lower()

        # ── Greetings ────────────────────────────────────────
        if any(w in c for w in ["hello","hi","hey","namaste","helo"]):
            self.speak("Hello sir! All modules operational.")

        # ── Time / Date ──────────────────────────────────────
        elif any(w in c for w in ["time","kitna baja"]):
            self.speak("Time is " + datetime.datetime.now().strftime("%I:%M %p"))
        elif any(w in c for w in ["date","aaj ka din","today"]):
            self.speak("Today is " + datetime.datetime.now().strftime("%A, %d %B %Y"))

        # ════════════════════════════════════════════════════
        #  MOUSE & KEYBOARD CONTROL
        # ════════════════════════════════════════════════════
        elif "move mouse" in c:
            # "move mouse to 500 300"
            nums = re.findall(r'\d+', c)
            if len(nums) >= 2:
                pyautogui.moveTo(int(nums[0]), int(nums[1]), duration=0.4)
                self.speak(f"Mouse moved to {nums[0]}, {nums[1]}.")
            else:
                self.speak("Please say coordinates, like: move mouse to 500 300.")

        elif "click" in c and "right" in c:
            pyautogui.rightClick()
            self.speak("Right click done.")
        elif "double click" in c:
            pyautogui.doubleClick()
            self.speak("Double click done.")
        elif "click" in c:
            pyautogui.click()
            self.speak("Click done.")

        elif "scroll up" in c:
            pyautogui.scroll(5)
            self.speak("Scrolled up.")
        elif "scroll down" in c:
            pyautogui.scroll(-5)
            self.speak("Scrolled down.")

        elif "type" in c:
            text_to_type = c.replace("type","").replace("jarvis","").strip()
            if text_to_type:
                pyautogui.typewrite(text_to_type, interval=0.05)
                self.speak(f"Typed: {text_to_type}")

        elif "press" in c:
            # "press ctrl c", "press enter", "press alt f4"
            keys = c.replace("press","").strip().split()
            if keys:
                try:
                    if len(keys) == 1:
                        pyautogui.press(keys[0])
                    else:
                        pyautogui.hotkey(*keys)
                    self.speak(f"Pressed {' + '.join(keys)}.")
                except Exception as e:
                    self.speak(f"Could not press those keys: {e}")

        elif "screenshot" in c:
            path = os.path.join(os.path.expanduser("~"), "Desktop",
                                f"jarvis_{int(time.time())}.png")
            pyautogui.screenshot(path)
            self.speak(f"Screenshot saved to Desktop.")

        # ════════════════════════════════════════════════════
        #  VOLUME CONTROL
        # ════════════════════════════════════════════════════
        elif any(w in c for w in ["volume up","awaaz badho","sound up"]):
            self._volume("up")
        elif any(w in c for w in ["volume down","awaaz kam","sound down"]):
            self._volume("down")
        elif any(w in c for w in ["mute","chup kar","mute karo"]):
            self._volume("mute")
        elif "set volume" in c:
            nums = re.findall(r'\d+', c)
            if nums:
                self._volume("set", int(nums[0]))

        # ════════════════════════════════════════════════════
        #  SCREEN BRIGHTNESS
        # ════════════════════════════════════════════════════
        elif any(w in c for w in ["brightness up","roshan karo","screen bright"]):
            self._brightness("up")
        elif any(w in c for w in ["brightness down","andhera karo","screen dim"]):
            self._brightness("down")
        elif "set brightness" in c:
            nums = re.findall(r'\d+', c)
            if nums: self._brightness("set", int(nums[0]))

        # ════════════════════════════════════════════════════
        #  CLIPBOARD MANAGER
        # ════════════════════════════════════════════════════
        elif any(w in c for w in ["clipboard","copy kya hai","what did i copy"]):
            content = pyperclip.paste()
            if content.strip():
                self.speak(f"Clipboard contains: {content[:120]}")
            else:
                self.speak("Clipboard is empty.")

        elif "copy" in c and "text" in c:
            text = c.replace("copy text","").replace("jarvis","").strip()
            pyperclip.copy(text)
            self.speak(f"Copied to clipboard: {text}")

        elif "clear clipboard" in c:
            pyperclip.copy("")
            self.speak("Clipboard cleared.")

        # ════════════════════════════════════════════════════
        #  PROCESS / APP MANAGER
        # ════════════════════════════════════════════════════
        elif any(w in c for w in ["open","start","launch","chalu karo"]):
            self._open_app(c)

        elif any(w in c for w in ["kill","close app","band karo app","terminate"]):
            app = c.replace("kill","").replace("close app","").replace("band karo app","").replace("terminate","").replace("jarvis","").strip()
            self._kill_process(app)

        elif any(w in c for w in ["list processes","running apps","kya chal raha","process list"]):
            self._list_processes()

        elif any(w in c for w in ["close all","sab band karo","close windows"]):
            self._close_all_windows()

        elif any(w in c for w in ["cpu","ram","memory","system info","system status"]):
            self._system_info()

        # ════════════════════════════════════════════════════
        #  FILE MANAGER
        # ════════════════════════════════════════════════════
        elif "create folder" in c or "new folder" in c:
            name = c.replace("create folder","").replace("new folder","").replace("jarvis","").strip()
            path = os.path.join(os.path.expanduser("~"), "Desktop", name or "NewFolder")
            os.makedirs(path, exist_ok=True)
            self.speak(f"Folder '{name or 'NewFolder'}' created on Desktop.")

        elif "delete file" in c or "file delete" in c:
            name = c.replace("delete file","").replace("file delete","").replace("jarvis","").strip()
            matches = glob.glob(os.path.join(os.path.expanduser("~"), "Desktop", name + "*"))
            if matches:
                os.remove(matches[0])
                self.speak(f"Deleted {os.path.basename(matches[0])}.")
            else:
                self.speak(f"File '{name}' not found on Desktop.")

        elif "copy file" in c:
            # "copy file report.txt to Documents"
            parts = c.replace("copy file","").strip().split(" to ")
            if len(parts) == 2:
                src  = os.path.join(os.path.expanduser("~"), "Desktop", parts[0].strip())
                dest = os.path.join(os.path.expanduser("~"), parts[1].strip())
                os.makedirs(dest, exist_ok=True)
                shutil.copy2(src, dest)
                self.speak(f"File copied to {parts[1].strip()}.")
            else:
                self.speak("Say: copy file filename to foldername.")

        elif "move file" in c:
            parts = c.replace("move file","").strip().split(" to ")
            if len(parts) == 2:
                src  = os.path.join(os.path.expanduser("~"), "Desktop", parts[0].strip())
                dest = os.path.join(os.path.expanduser("~"), parts[1].strip())
                os.makedirs(dest, exist_ok=True)
                shutil.move(src, dest)
                self.speak(f"File moved to {parts[1].strip()}.")

        elif "list files" in c or "desktop files" in c:
            files = os.listdir(os.path.join(os.path.expanduser("~"), "Desktop"))
            self.speak(f"Desktop has {len(files)} items. First few: {', '.join(files[:5])}.")

        elif "disk space" in c or "disk" in c or "storage" in c:
            t, u, f = shutil.disk_usage("/")
            self.speak(f"Disk: {t//2**30}GB total, {u//2**30}GB used, {f//2**30}GB free.")

        # ════════════════════════════════════════════════════
        #  NETWORK CONTROL
        # ════════════════════════════════════════════════════
        elif any(w in c for w in ["wifi off","wifi band","disable wifi","turn off wifi"]):
            os.system("netsh interface set interface Wi-Fi admin=disabled")
            self.speak("Wi-Fi disabled.")

        elif any(w in c for w in ["wifi on","wifi chalu","enable wifi","turn on wifi"]):
            os.system("netsh interface set interface Wi-Fi admin=enabled")
            self.speak("Wi-Fi enabled.")

        elif any(w in c for w in ["wifi list","available networks","networks nearby"]):
            result = subprocess.run(["netsh","wlan","show","networks"],
                                    capture_output=True, text=True)
            lines  = [l for l in result.stdout.split('\n') if 'SSID' in l and 'BSSID' not in l]
            self.speak(f"Found {len(lines)} networks: " + ", ".join(l.split(':')[1].strip() for l in lines[:5]))

        elif any(w in c for w in ["ip address","mera ip","my ip","what is my ip"]):
            ip = socket.gethostbyname(socket.gethostname())
            self.speak(f"Your local IP address is {ip}.")

        elif any(w in c for w in ["ping","internet check","internet speed","connected"]):
            result = subprocess.run(["ping","-n","2","8.8.8.8"], capture_output=True, text=True)
            if "Reply from" in result.stdout:
                self.speak("Internet connection is active. Google DNS is reachable.")
            else:
                self.speak("Cannot reach internet. You may be offline.")

        elif "flush dns" in c:
            os.system("ipconfig /flushdns")
            self.speak("DNS cache flushed.")

        elif "network info" in c or "ipconfig" in c:
            result = subprocess.run(["ipconfig"], capture_output=True, text=True)
            lines  = [l.strip() for l in result.stdout.split('\n') if 'IPv4' in l or 'Default Gateway' in l]
            self.speak("Network info: " + ". ".join(lines[:4]))

        # ════════════════════════════════════════════════════
        #  STARTUP APPS CONTROL
        # ════════════════════════════════════════════════════
        elif any(w in c for w in ["startup apps","startup list","boot programs"]):
            self._list_startup()

        elif "add startup" in c:
            app = c.replace("add startup","").strip()
            self._add_startup(app)

        elif "remove startup" in c:
            app = c.replace("remove startup","").strip()
            self._remove_startup(app)

        # ════════════════════════════════════════════════════
        #  REGISTRY (Advanced — Admin required)
        # ════════════════════════════════════════════════════
        elif "registry read" in c:
            # "registry read HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run"
            path = c.replace("registry read","").strip()
            self._registry_read(path)

        elif "registry set" in c:
            self.speak("Registry write commands require careful syntax. Please use the Python API directly for safety.")

        # ════════════════════════════════════════════════════
        #  POWER CONTROL
        # ════════════════════════════════════════════════════
        elif any(w in c for w in ["shutdown","band karo computer","pc band"]):
            self.speak("Shutting down in 15 seconds. Say cancel to abort.")
            time.sleep(3)
            os.system("shutdown /s /t 15")

        elif any(w in c for w in ["restart","reboot","dobara chalu"]):
            self.speak("Restarting in 15 seconds.")
            os.system("shutdown /r /t 15")

        elif any(w in c for w in ["sleep","hibernate","so jao","so ja"]):
            self.speak("Putting system to sleep.")
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")

        elif any(w in c for w in ["lock","screen lock","lock screen","computer lock"]):
            self.speak("Locking screen.")
            ctypes.windll.user32.LockWorkStation()

        elif any(w in c for w in ["cancel shutdown","abort","shutdown cancel"]):
            os.system("shutdown /a")
            self.speak("Shutdown cancelled.")

        elif any(w in c for w in ["log off","sign out"]):
            self.speak("Logging off.")
            os.system("shutdown /l")

        # ════════════════════════════════════════════════════
        #  SEARCH & WEB
        # ════════════════════════════════════════════════════
        elif "search" in c or "google" in c:
            q = re.sub(r'(search|google|jarvis)', '', c).strip()
            webbrowser.open(f"https://google.com/search?q={q}")
            self.speak(f"Searching for {q}.")

        elif "youtube" in c:
            q = re.sub(r'(youtube|play|jarvis)', '', c).strip()
            webbrowser.open(f"https://youtube.com/results?search_query={q}" if q else "https://youtube.com")
            self.speak(f"Opening YouTube{' for ' + q if q else ''}.")

        elif "wikipedia" in c:
            q = re.sub(r'(wikipedia|jarvis)', '', c).strip()
            webbrowser.open(f"https://en.wikipedia.org/wiki/{q.replace(' ','_')}")
            self.speak(f"Opening Wikipedia for {q}.")

        # ════════════════════════════════════════════════════
        #  BATTERY & SYSTEM INFO
        # ════════════════════════════════════════════════════
        elif "battery" in c:
            bat = psutil.sensors_battery()
            if bat:
                self.speak(f"Battery at {int(bat.percent)}%, {'charging' if bat.power_plugged else 'on battery'}.")
            else:
                self.speak("No battery detected.")

        # ════════════════════════════════════════════════════
        #  EXIT
        # ════════════════════════════════════════════════════
        elif any(w in c for w in ["exit","quit","bye","alvida","band ho","shutdown jarvis"]):
            self.speak("Goodbye sir. J.A.R.V.I.S going offline.")
            self.running = False

        else:
            self.speak("Command not recognised. Could you rephrase, sir?")


    # ─────────────────────────────────────────────────────────
    #  VOLUME
    # ─────────────────────────────────────────────────────────
    def _volume(self, action, level=None):
        try:
            from ctypes import cast, POINTER
            from comtypes import CLSCTX_ALL
            from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
            dev = AudioUtilities.GetSpeakers()
            iface = dev.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            vol = cast(iface, POINTER(IAudioEndpointVolume))
            cur = vol.GetMasterVolumeLevelScalar()
            if action == "up":
                vol.SetMasterVolumeLevelScalar(min(1.0, cur + 0.1), None)
                self.speak(f"Volume increased to {int(min(1.0,cur+0.1)*100)}%.")
            elif action == "down":
                vol.SetMasterVolumeLevelScalar(max(0.0, cur - 0.1), None)
                self.speak(f"Volume decreased to {int(max(0,cur-0.1)*100)}%.")
            elif action == "mute":
                muted = vol.GetMute()
                vol.SetMute(not muted, None)
                self.speak("Muted." if not muted else "Unmuted.")
            elif action == "set" and level is not None:
                vol.SetMasterVolumeLevelScalar(max(0, min(1, level/100)), None)
                self.speak(f"Volume set to {level}%.")
        except Exception:
            # keyboard fallback
            if action == "up":
                for _ in range(3): pyautogui.press("volumeup")
            elif action == "down":
                for _ in range(3): pyautogui.press("volumedown")
            elif action == "mute":
                pyautogui.press("volumemute")
            self.speak(f"Volume {action}.")

    # ─────────────────────────────────────────────────────────
    #  BRIGHTNESS
    # ─────────────────────────────────────────────────────────
    def _brightness(self, action, level=None):
        try:
            import screen_brightness_control as sbc
            cur = sbc.get_brightness(display=0)[0]
            if action == "up":
                sbc.set_brightness(min(100, cur + 15))
                self.speak(f"Brightness increased to {min(100,cur+15)}%.")
            elif action == "down":
                sbc.set_brightness(max(5, cur - 15))
                self.speak(f"Brightness decreased to {max(5,cur-15)}%.")
            elif action == "set" and level is not None:
                sbc.set_brightness(level)
                self.speak(f"Brightness set to {level}%.")
        except Exception as e:
            self.speak(f"Could not change brightness: {e}. Install screen-brightness-control.")

    # ─────────────────────────────────────────────────────────
    #  OPEN APPLICATION
    # ─────────────────────────────────────────────────────────
    def _open_app(self, cmd):
        apps = {
            "chrome":          "chrome",
            "firefox":         "firefox",
            "edge":            "msedge",
            "word":            "winword",
            "excel":           "excel",
            "powerpoint":      "powerpnt",
            "paint":           "mspaint",
            "vs code":         "code",
            "vscode":          "code",
            "notepad":         "notepad",
            "notepad++":       "notepad++",
            "calculator":      "calc",
            "cmd":             "cmd",
            "command prompt":  "cmd",
            "powershell":      "powershell",
            "task manager":    "taskmgr",
            "file explorer":   "explorer",
            "explorer":        "explorer",
            "control panel":   "control",
            "settings":        "ms-settings:",
            "device manager":  "devmgmt.msc",
            "disk management": "diskmgmt.msc",
            "registry editor": "regedit",
            "event viewer":    "eventvwr",
            "services":        "services.msc",
            "spotify":         "spotify",
            "discord":         "discord",
            "telegram":        "telegram",
            "whatsapp":        "whatsapp",
            "vlc":             "vlc",
            "snipping tool":   "snippingtool",
            "sticky notes":    "stikynot",
            "wordpad":         "wordpad",
        }
        for name, exe in apps.items():
            if name in cmd:
                try:
                    if exe.startswith("ms-"):
                        os.startfile(exe)
                    else:
                        subprocess.Popen(exe, shell=True)
                    self.speak(f"Opening {name}.")
                    return
                except:
                    self.speak(f"Could not find {name}.")
                    return
        self.speak("Application not recognised. Please specify the app name.")

    # ─────────────────────────────────────────────────────────
    #  KILL PROCESS
    # ─────────────────────────────────────────────────────────
    def _kill_process(self, name):
        if not name:
            self.speak("Please specify process name.")
            return
        killed = 0
        for proc in psutil.process_iter(['name','pid']):
            if name.lower() in proc.info['name'].lower():
                proc.kill()
                killed += 1
        if killed:
            self.speak(f"Killed {killed} instance(s) of {name}.")
        else:
            self.speak(f"No running process found named {name}.")

    # ─────────────────────────────────────────────────────────
    #  LIST PROCESSES
    # ─────────────────────────────────────────────────────────
    def _list_processes(self):
        procs = sorted(psutil.process_iter(['name','cpu_percent','memory_percent']),
                       key=lambda p: p.info['memory_percent'] or 0, reverse=True)
        top5  = [f"{p.info['name']} ({p.info['memory_percent']:.1f}% RAM)" for p in procs[:5]]
        self.speak(f"Top processes by memory: {', '.join(top5)}.")

    # ─────────────────────────────────────────────────────────
    #  CLOSE ALL WINDOWS
    # ─────────────────────────────────────────────────────────
    def _close_all_windows(self):
        self.speak("Closing all windows.")
        count = 0
        for w in gw.getAllWindows():
            try:
                if w.title.strip():
                    w.close(); count += 1; time.sleep(0.08)
            except: pass
        self.speak(f"Closed {count} windows.")

    # ─────────────────────────────────────────────────────────
    #  SYSTEM INFO
    # ─────────────────────────────────────────────────────────
    def _system_info(self):
        cpu  = psutil.cpu_percent(interval=1)
        ram  = psutil.virtual_memory()
        disk = psutil.disk_usage('/')
        self.speak(
            f"CPU at {cpu}%. "
            f"RAM: {ram.used//2**20} MB used of {ram.total//2**20} MB. "
            f"Disk: {disk.free//2**30} GB free."
        )

    # ─────────────────────────────────────────────────────────
    #  STARTUP APPS
    # ─────────────────────────────────────────────────────────
    def _list_startup(self):
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        try:
            key   = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path)
            items = []
            i = 0
            while True:
                try:
                    name, _, _ = winreg.EnumValue(key, i)
                    items.append(name); i += 1
                except OSError: break
            winreg.CloseKey(key)
            self.speak(f"Startup apps: {', '.join(items[:8]) if items else 'None found'}.")
        except Exception as e:
            self.speak(f"Could not read startup registry: {e}")

    def _add_startup(self, app_name):
        try:
            exe_path = shutil.which(app_name)
            if not exe_path:
                self.speak(f"Cannot find executable for {app_name}.")
                return
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                 r"Software\Microsoft\Windows\CurrentVersion\Run",
                                 0, winreg.KEY_SET_VALUE)
            winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, exe_path)
            winreg.CloseKey(key)
            self.speak(f"{app_name} added to startup.")
        except Exception as e:
            self.speak(f"Failed to add startup entry: {e}")

    def _remove_startup(self, app_name):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                 r"Software\Microsoft\Windows\CurrentVersion\Run",
                                 0, winreg.KEY_SET_VALUE)
            winreg.DeleteValue(key, app_name)
            winreg.CloseKey(key)
            self.speak(f"{app_name} removed from startup.")
        except FileNotFoundError:
            self.speak(f"{app_name} was not in startup list.")
        except Exception as e:
            self.speak(f"Failed to remove startup entry: {e}")

    # ─────────────────────────────────────────────────────────
    #  REGISTRY READ
    # ─────────────────────────────────────────────────────────
    def _registry_read(self, path):
        hive_map = {
            "HKCU": winreg.HKEY_CURRENT_USER,
            "HKLM": winreg.HKEY_LOCAL_MACHINE,
            "HKCR": winreg.HKEY_CLASSES_ROOT,
        }
        try:
            parts = path.strip().split("\\", 1)
            hive  = hive_map.get(parts[0].upper(), winreg.HKEY_CURRENT_USER)
            sub   = parts[1] if len(parts) > 1 else ""
            key   = winreg.OpenKey(hive, sub)
            items = []
            i = 0
            while True:
                try:
                    name, data, _ = winreg.EnumValue(key, i)
                    items.append(f"{name}={data}"); i += 1
                except OSError: break
            winreg.CloseKey(key)
            self.speak(f"Registry has {len(items)} values. First few: {', '.join(items[:3])}.")
        except Exception as e:
            self.speak(f"Registry read failed: {e}")

    # ─────────────────────────────────────────────────────────
    #  MAIN LOOP
    # ─────────────────────────────────────────────────────────
    def run(self):
        print("\n" + "═"*60)
        print("   J.A.R.V.I.S  FULL SYSTEM CONTROL  —  ONLINE")
        print(f"   Say '{self.wake_word.upper()}' to activate, then speak your command.")
        print("   Or just speak — commands processed continuously.")
        print("═"*60 + "\n")

        while self.running:
            heard = self.listen(timeout=10, phrase_limit=5)

            if self.wake_word in heard:
                self.speak("Yes sir?")
                command = self.listen(timeout=8, phrase_limit=18)
                if command:
                    self.process(command)
            elif heard:
                self.process(heard)

        print("\n[JARVIS] Offline.")


# ═══════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    if not is_admin():
        print("⚠️  TIP: Run as Administrator for full system access (registry, network, etc.)")
    try:
        j = Jarvis()
        j.run()
    except KeyboardInterrupt:
        print("\n[JARVIS] Interrupted. Goodbye!")