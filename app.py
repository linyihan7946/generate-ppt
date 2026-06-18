import os

from generate_ppt.web import create_app, find_free_port, open_browser_later


app = create_app()


if __name__ == "__main__":
    port = int(os.environ.get("PORT") or find_free_port(8765))
    url = f"http://127.0.0.1:{port}"
    print(f"generate-ppt running at {url}")
    open_browser_later(url)
    app.run(host="127.0.0.1", port=port, debug=False)
