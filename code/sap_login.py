import subprocess
import win32com.client
import win32gui
import time
import json
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
SAP_PATH = r'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe'
SAP_SYSTEM = "PRD - ECC Production"
CLIENT = "900"
LANGUAGE = "EN"


def load_credentials(filepath: str | Path) -> dict:
    try:
        with open(filepath, "r") as f:
            data = json.load(f)
    except FileNotFoundError:
        print(f"Credentials file not found: {filepath}", file=sys.stderr)
        raise
    except json.JSONDecodeError as e:
        print(f"Invalid JSON in credentials file: {e}", file=sys.stderr)
        raise
    for key in ("username", "password"):
        if key not in data:
            raise KeyError(f"Missing required key '{key}' in credentials file: {filepath}")
    return data


class SapAutomation:
    def __init__(self, credentials_path: str | Path):
        self.credentials_path = Path(credentials_path)
        creds = load_credentials(credentials_path)
        self.username: str = creds["username"].strip()
        self.password: str = creds["password"]
        self.session     = None  # session1 — VL06F (lectura + verificaciones + BOL)
        self.session2    = None  # session2 — Operaciones (VF11, VI05, VT02N, VL09)
        self.session3    = None  # session3 — Órdenes (ZCMR, VA02, ME22N)
        self.app         = None
        self._connection = None  # referencia a la conexión que abrimos nosotros

    def launch_sap(self) -> bool:
        try:
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            if isinstance(sap_gui_auto, win32com.client.CDispatch):
                engine = sap_gui_auto.GetScriptingEngine
                _ = engine.Children.Count
                self.app = engine
                print("SAP GUI already running, reusing existing instance.")
                return True
        except Exception:
            pass
        subprocess.Popen(SAP_PATH)
        return self._wait_for_sap()

    def _wait_for_sap(self, retries: int = 10, delay: float = 2.0) -> bool:
        for attempt in range(retries):
            try:
                sap_gui_auto = win32com.client.GetObject("SAPGUI")
                if isinstance(sap_gui_auto, win32com.client.CDispatch):
                    self.app = sap_gui_auto.GetScriptingEngine
                    return True
            except Exception:
                print(f"Waiting for SAP GUI... attempt {attempt + 1}/{retries}")
                time.sleep(delay)
        return False

    def open_session(self, retries: int = 10, delay: float = 2.0) -> bool:
        try:
            # Reutilizar sesión existente si el usuario ya está logueado
            for i in range(self.app.Children.Count):
                conn = self.app.Children(i)
                if conn.Children.Count > 0:
                    candidate = conn.Children(0)
                    # Verificar que no sea una pantalla de login
                    try:
                        candidate.findById("wnd[0]/usr/txtRSYST-BNAME")
                        continue  # Es pantalla de login, ignorar
                    except Exception:
                        pass
                    self._connection = conn
                    self.session = candidate
                    self.session.findById("wnd[0]").maximize()
                    print("Sesión SAP existente reutilizada.")
                    return True

            # No hay sesión activa — abrir nueva conexión y hacer login
            print("No se encontró sesión activa. Abriendo nueva conexión...")
            self._connection = self.app.OpenConnection(SAP_SYSTEM, True)
            for attempt in range(retries):
                try:
                    self.session = self._connection.Children[0]
                    self.session.findById("wnd[0]").maximize()
                    return True
                except Exception:
                    print(f"Waiting for session... attempt {attempt + 1}/{retries}")
                    time.sleep(delay)
            return False
        except Exception as e:
            print(f"Failed to open session: {e}", file=sys.stderr)
            return False

    def login(self) -> bool:
        try:
            creds = load_credentials(self.credentials_path)
            self.username = creds["username"].strip()
            self.password = creds["password"]

            self._print_login_diagnostics()

            self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = CLIENT
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = self.username
            subprocess.run(["clip"], input=self.password.encode("utf-16-le"), check=True)
            pwd_field = self.session.findById("wnd[0]/usr/pwdRSYST-BCODE")
            pwd_field.SetFocus()
            win32gui.SetForegroundWindow(self.session.findById("wnd[0]").Handle)
            time.sleep(0.3)
            win32com.client.Dispatch("WScript.Shell").SendKeys("^v")
            subprocess.run(["clip"], input="".encode("utf-16-le"), check=True)
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = LANGUAGE
            self._print_filled_fields()
            self.session.findById("wnd[0]").sendVKey(0)
            self._wait_for_session_ready()

            if self._is_still_on_login_screen():
                status_message, _ = self._get_status_message()
                if (
                    self._is_new_password_screen()
                    or "campos obligatorios" in status_message.lower()
                    or "new password" in status_message.lower()
                ):
                    print(
                        "SAP is asking you to set a new password. Complete the password change "
                        "manually in SAP, then update credentials.json and run the script again.",
                        file=sys.stderr,
                    )
                else:
                    print(f"Login failed: {status_message}", file=sys.stderr)
                return False

            self.session.findById("wnd[0]").sendVKey(0)
            print("Login successful")
            return True
        except Exception as e:
            print(f"Login failed: {e}", file=sys.stderr)
            return False

    def _wait_for_session_ready(self, timeout: float = 10.0, poll: float = 0.5) -> None:
        end = time.time() + timeout
        while time.time() < end:
            try:
                if not self.session.Busy:
                    return
            except Exception:
                pass
            time.sleep(poll)

    def _print_login_diagnostics(self) -> None:
        print("SAP login attempt")
        print(f"  System: {SAP_SYSTEM}")
        print(f"  Client: {CLIENT}")
        print(f"  Username: {self.username}")
        print(f"  Language: {LANGUAGE}")
        print(f"  Credentials file: {self.credentials_path}")
        print(f"  Password length: {len(self.password)} characters")

        if self.password != self.password.strip():
            print("  Warning: password starts or ends with whitespace.")

    def _print_filled_fields(self) -> None:
        try:
            client = self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text
            username = self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text
            language = self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text
            print("Fields filled in SAP")
            print(f"  Client field: {client}")
            print(f"  Username field: {username}")
            print(f"  Language field: {language}")
        except Exception as e:
            print(f"Could not read filled SAP fields: {e}", file=sys.stderr)

    def _is_still_on_login_screen(self) -> bool:
        try:
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME")
            return True
        except Exception:
            return False

    def _is_new_password_screen(self) -> bool:
        try:
            for field_id in ("wnd[0]/usr/pwdRSYST-NCODE", "wnd[0]/usr/pwdRSYST-NCOD2"):
                try:
                    self.session.findById(field_id)
                    return True
                except Exception:
                    pass
            window_text = self.session.findById("wnd[0]").Text.lower()
            return "clave acceso nueva" in window_text or "new password" in window_text
        except Exception:
            return False

    def _get_status_message(self) -> tuple[str, str]:
        try:
            sbar = self.session.findById("wnd[0]/sbar")
            return sbar.text.strip(), sbar.messageType.strip()
        except Exception:
            return "", ""

    def _wait_for_session_count(self, connection, expected: int, timeout: float = 10.0) -> None:
        """Espera hasta que la conexión tenga al menos `expected` sesiones abiertas."""
        end = time.time() + timeout
        while time.time() < end:
            if connection.Children.Count >= expected:
                return
            time.sleep(0.5)
        raise TimeoutError(
            f"Se esperaban {expected} sesiones SAP pero solo hay "
            f"{connection.Children.Count} después de {timeout}s."
        )

    def open_additional_sessions(self) -> bool:
        """
        Abre las sesiones 2 y 3 después del login.
        session1 (self.session) ya existe; este método agrega session2 y session3.
        Usa self._connection (guardada en open_session) para evitar confundir
        nuestra conexión con otras que el usuario pueda tener abiertas.
        """
        try:
            connection = self._connection
            sessions_before = connection.Children.Count
            print(f"  [Login] Sesiones SAP antes de createSession: {sessions_before}")

            # Abrir session2
            self.session.createSession()
            self._wait_for_session_count(connection, 2)
            self.session2 = connection.Children(1)
            self.session2.findById("wnd[0]").maximize()

            # Abrir session3
            self.session.createSession()
            self._wait_for_session_count(connection, 3)
            self.session3 = connection.Children(2)
            self.session3.findById("wnd[0]").maximize()

            print("3 sesiones SAP abiertas correctamente.")
            print("  session1 -> VL06F (lectura + verificaciones + BOL)")
            print("  session2 -> Operaciones (VF11, VI05, VT02N, VL09)")
            print("  session3 -> Órdenes (ZCMR, VA02, ME22N)")
            return True

        except Exception as e:
            sessions_now = 0
            try:
                if self._connection is not None:
                    sessions_now = self._connection.Children.Count
            except Exception:
                pass
            print("", file=sys.stderr)
            print(f"  ADVERTENCIA: No se pudieron abrir las 3 sesiones SAP.", file=sys.stderr)
            print(f"  Causa: {e!r}", file=sys.stderr)
            print(f"  Sesiones SAP actualmente abiertas: {sessions_now}", file=sys.stderr)
            print(f"  Cierre TODAS las ventanas SAP innecesarias y vuelva a ejecutar.", file=sys.stderr)
            print("", file=sys.stderr)
            # Fallback: asignar session2/3 a session1 para que el módulo siga siendo
            # usable en tests/diagnóstico. El fail-fast lo aplica main.py.
            self.session2 = self.session
            self.session3 = self.session
            return False

    def run(self):
        if not self.launch_sap():
            print("Could not connect to SAP GUI after retries.", file=sys.stderr)
            return
        if not self.open_session():
            print("Could not open SAP session.", file=sys.stderr)
            return
        # Si open_session reutilizó una sesión activa, el login ya fue hecho manualmente
        needs_login = self._is_still_on_login_screen()
        if needs_login:
            if not self.login():
                return
        self.open_additional_sessions()


if __name__ == "__main__":
    sap = SapAutomation(CREDENTIALS_FILE)
    sap.run()
