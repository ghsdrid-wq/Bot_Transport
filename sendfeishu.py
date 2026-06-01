import json
import os
import time

import requests


def request_with_retry(func, retries=3, delay=2, log=None, name=""):
    for attempt in range(retries):
        try:
            response = func()
            response.raise_for_status()
            return response.json()
        except Exception as error:
            if log:
                log(f"{name} failed ({attempt + 1}/{retries}): {error}")
            if attempt < retries - 1:
                time.sleep(delay * (attempt + 1))
            else:
                raise


def get_token(app_id, app_secret, log=None):
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal/"
    result = request_with_retry(
        lambda: requests.post(url, json={"app_id": app_id, "app_secret": app_secret}, timeout=10),
        log=log,
        name="get_token",
    )
    if "tenant_access_token" not in result:
        raise RuntimeError(f"Get token failed: {result}")
    if log:
        log("Feishu token acquired")
    return result["tenant_access_token"]


def upload_image(token, path, log=None):
    url = "https://open.feishu.cn/open-apis/im/v1/images"
    headers = {"Authorization": f"Bearer {token}"}

    def upload():
        with open(path, "rb") as image_file:
            return requests.post(
                url,
                headers=headers,
                files={"image": image_file},
                data={"image_type": "message"},
                timeout=30,
            )

    result = request_with_retry(upload, log=log, name="upload_image")
    if result.get("code") != 0:
        raise RuntimeError(f"Upload image failed: {result}")
    return result["data"]["image_key"]


def send_image(token, chat_id, image_key, log=None):
    url = "https://open.feishu.cn/open-apis/im/v1/messages"
    headers = {"Authorization": f"Bearer {token}"}
    params = {"receive_id_type": "chat_id"}
    payload = {"receive_id": chat_id, "msg_type": "image", "content": json.dumps({"image_key": image_key})}
    result = request_with_retry(
        lambda: requests.post(url, headers=headers, params=params, json=payload, timeout=10),
        log=log,
        name="send_image",
    )
    if result.get("code") != 0:
        raise RuntimeError(f"Send image failed: {result}")


def run_send(image_path, chat_id, app_id, app_secret, log=None):
    write = log or print
    if not chat_id:
        raise ValueError("Missing Feishu Chat ID")
    if not app_id or not app_secret:
        raise ValueError("Missing Feishu App ID / App secret")
    if not os.path.exists(image_path):
        raise ValueError(f"File not found: {image_path}")
    token = get_token(app_id, app_secret, log=write)
    write(f"Uploading: {image_path}")
    image_key = upload_image(token, image_path, log=write)
    write(f"Sending image to Chat ID: {chat_id}")
    send_image(token, chat_id, image_key, log=write)
    write("Feishu image sent")
