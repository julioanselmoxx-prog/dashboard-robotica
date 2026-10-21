import os

# Configuración de credenciales SharePoint / Microsoft Graph API
SHAREPOINT_TENANT_ID = os.getenv("SP_TENANT_ID", "tu-tenant-id-here")
SHAREPOINT_CLIENT_ID = os.getenv("SP_CLIENT_ID", "tu-client-id-here")
SHAREPOINT_CLIENT_SECRET = os.getenv("SP_CLIENT_SECRET", "tu-client-secret-here")
SHAREPOINT_SITE_ID = os.getenv("SP_SITE_ID", "tu-site-id-here")
SHAREPOINT_FILE_PATH = "General/Robotica/Bitacora_Robots_Campo.xlsx"

# Configuración de APIs de Drones / Robots
DRONE_API_URL = os.getenv("DRONE_API_URL", "https://api.equans-drones.com/v1")
DRONE_API_KEY = os.getenv("DRONE_API_KEY", "demo_drone_key_12345")

ROBOT_API_URL = os.getenv("ROBOT_API_URL", "https://api.equans-robots.com/v1")
ROBOT_API_KEY = os.getenv("ROBOT_API_KEY", "demo_robot_key_67890")