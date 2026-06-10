import requests, json

REFRESH_TOKEN = "1000.9b2190e906bcb52c5944970e44497ac4.4cdc09459045e0865441fdceed03cf3e"
CLIENT_ID     = "1000.C37R6R47ZDGF9Y635H1J6YPTERUEQN"
CLIENT_SECRET = "d187d8556f52883e9624b1f55a5495f84d31d383c3"
PORTAL_ID     = "757079135"

# 1. Obtener token
r = requests.post("https://accounts.zoho.com/oauth/v2/token", params={
    "refresh_token": REFRESH_TOKEN,
    "client_id":     CLIENT_ID,
    "client_secret": CLIENT_SECRET,
    "grant_type":    "refresh_token",
})
token = r.json().get("access_token")
print("Token OK:", bool(token))

# 2. Un proyecto de muestra
headers = {"Authorization": f"Zoho-oauthtoken {token}"}
r2 = requests.get(f"https://projectsapi.zoho.com/restapi/portal/{PORTAL_ID}/projects/", 
                  headers=headers, params={"status": "active", "range": 1})
projects = r2.json().get("projects", [])
if projects:
    p = projects[0]
    print("\n── KEYS del proyecto ──")
    print(list(p.keys()))
    print("\n── Campos custom relevantes ──")
    for k in ["razon_social","rut_empresa","plan_contratado","vendedor","modulo_vendido","nombre_del_contacto","cantidad_de_empleados","status"]:
        print(f"  {k}: {p.get(k, 'NO EXISTE')}")
    
    # 3. Tareas del primer proyecto
    pid = p.get("id")
    r3 = requests.get(f"https://projectsapi.zoho.com/restapi/portal/{PORTAL_ID}/projects/{pid}/tasks/",
                      headers=headers, params={"range": 3})
    print(f"\n── Respuesta tareas (status {r3.status_code}) ──")
    print(r3.text[:500])
