import requests
import urllib3

from config import API_URL, USER_TOKEN, APP_TOKEN, VERIFY_SSL


# Evita o aviso do certificado quando VERIFY_SSL = False
if not VERIFY_SSL:
    urllib3.disable_warnings(
        urllib3.exceptions.InsecureRequestWarning
    )


class GLPI:

    def __init__(self):

        self.session_token = None

        # Cabeçalhos utilizados para iniciar a sessão
        self.headers_auth = {
            "Authorization": f"user_token {USER_TOKEN}",
            "App-Token": APP_TOKEN
        }

        # Cabeçalhos utilizados depois que a sessão é criada
        self.headers_api = {
            "App-Token": APP_TOKEN
        }


    def conectar(self):

        response = requests.get(
            f"{API_URL}/initSession",
            headers=self.headers_auth,
            verify=VERIFY_SSL
        )

        response.raise_for_status()

        self.session_token = response.json()["session_token"]

        self.headers_api["Session-Token"] = self.session_token

        print("Conectado à API do GLPI.")


    def buscar_chamados(self, params):

        response = requests.get(
            f"{API_URL}/search/Ticket",
            headers=self.headers_api,
            params=params,
            verify=VERIFY_SSL
        )

        response.raise_for_status()

        return response.json()


    def obter_chamado(self, chamado_id):

        response = requests.get(
            f"{API_URL}/Ticket/{chamado_id}",
            headers=self.headers_api,
            verify=VERIFY_SSL
        )

        response.raise_for_status()

        return response.json()


    def desconectar(self):

        if not self.session_token:
            return

        try:

            requests.get(
                f"{API_URL}/killSession",
                headers=self.headers_api,
                verify=VERIFY_SSL
            )

            print("Sessão encerrada.")

        except Exception:
            pass