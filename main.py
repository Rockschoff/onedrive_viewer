import streamlit as st
import requests
import msal
from io import BytesIO

# --- Page Configuration ---
st.set_page_config(
    page_title="Efficient OneDrive Explorer",
    page_icon="📂",
    layout="wide",
)

# --- MSAL Authentication & Graph API Functions ---

def get_access_token(tenant_id, client_id, client_secret):
    """Authenticates and returns an access token."""
    if not all([tenant_id, client_id, client_secret]):
        st.error("One or more credential values are missing.")
        return None

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=authority
    )

    scopes = ["https://graph.microsoft.com/.default"]
    result = app.acquire_token_for_client(scopes=scopes)

    if "access_token" in result:
        return result['access_token']
    else:
        st.error(f"Authentication Failed: {result.get('error_description', 'No error description.')}")
        return None

# --- Cached API Functions for Efficiency ---
@st.cache_data(ttl=600, show_spinner=False)
def get_drive_children_cached(_drive_id, item_id, headers):
    """
    Fetches and caches ALL children of a specific folder (item_id) in a drive,
    following @odata.nextLink pagination.
    """
    url = f"https://graph.microsoft.com/v1.0/drives/{_drive_id}/items/{item_id}/children"
    items = []

    try:
        while url:
            resp = requests.get(url, headers=headers, timeout=30)
            resp.raise_for_status()
            data = resp.json()

            items.extend(data.get("value", []))
            url = data.get("@odata.nextLink") # follow pagination until exhausted

        return items

    except requests.exceptions.RequestException as e:
        st.error(f"API Error fetching children for item {item_id}: {e}")
        return items
    except Exception as e:
        st.error(f"An unexpected error occurred while fetching children: {e}")
        return items

@st.cache_data(ttl=600, show_spinner=False)
def get_file_content_from_url_cached(download_url):
    """
    Fetches and caches the content of a specific file using its pre-authenticated download URL.
    """
    if not download_url:
        st.error("Download URL is missing.")
        return None
    try:
        resp = requests.get(download_url)
        resp.raise_for_status()
        return resp.content
    except requests.exceptions.RequestException as e:
        st.error(f"API Error fetching content from download URL: {e}")
        return None
    except Exception as e:
        st.error(f"An unexpected error occurred while fetching file content: {e}")
        return None

@st.cache_data(ttl=600, show_spinner=False)
def get_sharepoint_fields_cached(drive_id, item_id, headers):
    """
    Fetches and caches SharePoint listItem.fields for a drive item.
    """
    url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}"
        f"/items/{item_id}"
        f"?$expand=listItem($expand=fields)"
    )
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    return resp.json().get("listItem", {}).get("fields", {})

# --- Streamlit UI Components ---

def display_breadcrumbs():
    """Displays the navigation breadcrumbs and handles navigation."""
    path_items = st.session_state.get('path', [])
    cols = st.columns(len(path_items))
    for i, (name, item_id) in enumerate(path_items):
        with cols[i]:
            if st.button(f"▸ {name}", key=f"crumb_{item_id}"):
                st.session_state.path = st.session_state.path[:i + 1]
                clear_download_state()
                st.rerun()
    st.markdown("---")


def clear_download_state():
    """Callback function to clear the download state."""
    st.session_state.download_target_id = None

HIDE_PREFIXES = ("MEX-", "NIA-")

def _is_hidden_at_root(name: str, item_id: str) -> bool:
    # Only hide in the root; case-insensitive on the name
    # print(name.upper() , name.upper().startswith(HIDE_PREFIXES))
    return item_id == "root" and name.upper().startswith(HIDE_PREFIXES)

def display_folder_contents(drive_id, headers, item_id, folder_name):
    """Displays folder contents, including QC Document Number metadata."""
    with st.spinner(f"Loading contents of '{folder_name}'..."):
        children = get_drive_children_cached(drive_id, item_id, headers)

    if not children:
        st.info("_(This folder is empty)_")
        return

    folders = [
        c for c in children
        if "folder" in c and not _is_hidden_at_root(c.get("name", ""), item_id)
    ]
    folders = sorted([c for c in folders if 'folder' in c], key=lambda x: x['name'])
    files = sorted([c for c in children if 'file' in c], key=lambda x: x['name'])

    # Folders first
    for folder in folders:
        col1, col2 = st.columns([0.8, 0.2])
        with col1:
            st.write(f"📁 {folder['name']}")
        with col2:
            if st.button("Open", key=f"open_{folder['id']}", help=f"Open {folder['name']}"):
                st.session_state.path.append((folder['name'], folder['id']))
                clear_download_state()
                st.rerun()

    if folders and files:
        st.markdown("---")

    # Files and metadata
    for file_item in files:
        file_id = file_item['id']
        fields = get_sharepoint_fields_cached(drive_id, file_id, headers)
        qc_number = fields.get('QCDocumentNumber', "-")

        col1, col2, col3, col4 = st.columns([0.6, 0.2, 0.15, 0.05])
        with col1:
            st.write(f"📄 {file_item['name']}")
        with col2:
            size = file_item.get('size', 0)
            if size < 1024:
                size_str = f"{size} B"
            elif size < 1024**2:
                size_str = f"{size/1024:.1f} KB"
            else:
                size_str = f"{size/1024**2:.1f} MB"
            st.write(f"_{size_str}_")
        with col3:
            st.write(qc_number)
        with col4:
            if st.session_state.get('download_target_id') == file_id:
                with st.spinner("Preparing..."):
                    download_url = file_item.get('@microsoft.graph.downloadUrl')
                    if download_url:
                        content = get_file_content_from_url_cached(download_url)
                        if content:
                            st.download_button(
                                label="✅ Save",
                                data=content,
                                file_name=file_item['name'],
                                mime='application/octet-stream',
                                key=f"final_dl_{file_id}",
                                on_click=clear_download_state,
                                help=f"Click to save {file_item['name']}"
                            )
                        else:
                            st.error("Download failed.")
                            if st.button("Retry", key=f"retry_{file_id}", on_click=clear_download_state):
                                st.rerun()
                    else:
                        st.error("URL not found.")
                        if st.button("OK", key=f"ok_err_{file_id}", on_click=clear_download_state):
                            st.rerun()
            else:
                if st.button("Download", key=f"download_{file_id}", help=f"Download {file_item['name']}"):
                    st.session_state.download_target_id = file_id
                    st.rerun()


def main():
    st.title("☁️ Efficient OneDrive Explorer")

    # Init session state
    if 'path' not in st.session_state:
        st.session_state.path = [("Root", "root")]
    if 'download_target_id' not in st.session_state:
        st.session_state.download_target_id = None

    # Load secrets
    try:
        tenant_id = st.secrets["TENANT_ID"]
        client_id = st.secrets["APPLICATION_ID"]
        client_secret = st.secrets["CLIENT_SECRET"]
        drive_id = st.secrets["DRIVE_ID"]
        credentials_found = True
    except KeyError:
        credentials_found = False

    if not credentials_found:
        st.error("Required credentials not found in st.secrets.")
        st.info("Please configure your secrets for deployment.")
        st.markdown("""
        **Example `secrets.toml`:**
        ```toml
        TENANT_ID = "your-tenant-id"
        APPLICATION_ID = "your-application-id"
        CLIENT_SECRET = "your-client-secret"
        DRIVE_ID = "your-drive-id"
        ```
        """)
        st.stop()

    with st.spinner("Authenticating with Microsoft Graph API..."):
        access_token = get_access_token(tenant_id, client_id, client_secret)

    if access_token:
        st.sidebar.success("✅ Authentication Successful")
        st.sidebar.markdown("**Drive ID:**")
        st.sidebar.code(drive_id)

        headers = {'Authorization': f"Bearer {access_token}"}

        st.header("File Explorer")
        display_breadcrumbs()

        current_folder_name, current_item_id = st.session_state.path[-1]
        display_folder_contents(drive_id, headers, current_item_id, current_folder_name)
    else:
        st.error("Could not authenticate. Check credentials and API permissions.")
        st.stop()

if __name__ == "__main__":
    main()