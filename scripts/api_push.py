# -*- coding: utf-8 -*-
"""通过 GitHub Git Data API 推送提交（github.com:443 不可达时的备用通道）。
用法：python scripts/api_push.py <github_token> <owner/repo> <branch> [message]
"""
import base64
import json
import subprocess
import sys
import urllib.request


def api(token, endpoint, payload=None, method=None):
    url = f"https://api.github.com{endpoint}"
    data = json.dumps(payload).encode() if payload is not None else None
    req = urllib.request.Request(url, data=data, method=method or ("POST" if data else "GET"))
    req.add_header("Authorization", f"token {token}")
    req.add_header("Accept", "application/vnd.github+json")
    req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            body = resp.read()
            return json.loads(body) if body else {}
    except urllib.error.HTTPError as e:
        print(f"API 失败 {endpoint} -> {e.code}: {e.read().decode()[:300]}")
        raise


def main():
    token, repo, branch = sys.argv[1], sys.argv[2], sys.argv[3]
    message = sys.argv[4] if len(sys.argv) > 4 else "update"
    files = subprocess.run(["git", "ls-files"], capture_output=True, text=True,
                           encoding="utf-8").stdout.splitlines()
    files = [f for f in files if f.strip()]
    print(f"待上传文件 {len(files)} 个")

    tree = []
    for i, path in enumerate(files, 1):
        with open(path, "rb") as fh:
            content = base64.b64encode(fh.read()).decode()
        blob = api(token, f"/repos/{repo}/git/blobs",
                   {"content": content, "encoding": "base64"})
        tree.append({"path": path.replace("\\", "/"), "mode": "100644",
                     "type": "blob", "sha": blob["sha"]})
        print(f"[{i}/{len(files)}] {path}")

    new_tree = api(token, f"/repos/{repo}/git/trees", {"tree": tree})
    # 已有分支则在其上提交，否则作为首个提交创建分支
    try:
        ref = api(token, f"/repos/{repo}/git/ref/heads/{branch}")
        parents = [ref["object"]["sha"]]
        commit = api(token, f"/repos/{repo}/git/commits",
                     {"message": message, "tree": new_tree["sha"], "parents": parents})
        api(token, f"/repos/{repo}/git/refs/heads/{branch}",
            {"sha": commit["sha"], "force": True}, method="PATCH")
    except urllib.error.HTTPError:
        commit = api(token, f"/repos/{repo}/git/commits",
                     {"message": message, "tree": new_tree["sha"], "parents": []})
        api(token, f"/repos/{repo}/git/refs",
            {"ref": f"refs/heads/{branch}", "sha": commit["sha"]})
    print("已创建提交", commit["sha"], "并更新分支", branch)


if __name__ == "__main__":
    main()
