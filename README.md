# QDII 海外额度宝 · 额度看板

QDII 基金单日申购额度排行看板（741 只份额），数据源为天天基金移动端公开接口。

- `docs/index.html` —— 静态快照（数据内嵌，可直接双击打开或托管到 GitHub Pages）
- `data/quota_data.json` —— 看板结构化数据
- `app.py` —— 本地看板服务（http://localhost:5001，每日 08:30 自动同步）
- `sync_data.py` —— 全量抓取并生成快照 / Excel
- `publish.py` / `publish.bat` —— 把当天快照同步到 `docs/` 并 commit + push

## 本地使用

双击 `start_server.bat`，浏览器打开 http://localhost:5001

## 发布到 GitHub Pages

1. 在 GitHub 新建仓库（例如 `qdii-quota-dashboard`），**Public**（免费版 Pages 需要公开仓库）
2. 关联远端并推送：

```bash
git remote add origin https://github.com/<用户名>/<仓库>.git
git push -u origin main
```

3. 仓库 Settings → Pages → Source 选 `main` 分支，目录选 `/docs`
4. 访问 `https://<用户名>.github.io/<仓库>/`

之后每天同步完成后执行 `publish.bat`（或 `python publish.py`）即可更新线上看板。

## 注意

- 本仓库已通过 `.gitignore` 排除含内部渠道标注的 Excel、抓取缓存与日志，推送前请确认无敏感内容。
- `docs/index.html` 每次发布约 430KB，一年约 100MB 历史；如仓库过大，可定期清理历史或改为只保留最新数据。
