# run_streamlit_cli.py
import os
import sys

def resource_path(rel):
    try:
        base = sys._MEIPASS  # PyInstaller temp folder
    except AttributeError:
        base = os.path.abspath(".")
    return os.path.join(base, rel)

if __name__ == "__main__":
    app_script = resource_path("streamlit_ui.py")

    # Disable telemetry & dev mode
    os.environ.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")
    os.environ.setdefault("STREAMLIT_SERVER_RUN_ON_SAVE", "false")
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"

    # Equivalent to: streamlit run streamlit_ui.py --server.address localhost
    # Let Streamlit pick a default port in prod mode
    sys.argv = [
        "streamlit", "run", app_script,
        "--server.address", "localhost"
    ]

    from streamlit.web.cli import main as stcli
    stcli()
