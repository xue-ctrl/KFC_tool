# KFC 素材核查助手：Docker 部署说明

本项目使用 Docker Compose 部署，默认开放 `8790` 端口。统一入口会同时提供美团和抖音核对功能。

## 1. 准备服务器

服务器建议使用 Linux（Ubuntu 22.04 / 24.04 或 CentOS 系列均可），并确保已安装 Docker Engine 与 Docker Compose 插件。

Ubuntu / Debian 可执行：

```bash
sudo apt update
sudo apt install -y docker.io docker-compose-plugin
sudo systemctl enable --now docker
```

确认安装成功：

```bash
docker --version
docker compose version
```

如服务器开启了防火墙，先放行端口：

```bash
sudo ufw allow 8790/tcp
```

同时请在云服务器控制台的安全组 / 防火墙中放行 TCP 8790；仅公司内网使用时，建议只允许公司出口 IP 或 VPN 网段访问。

## 2. 上传项目

将项目压缩包上传到服务器，例如放到 `/opt`，再执行：

```bash
cd /opt
unzip KFC素材核查店小助_最终版.zip
cd KFC素材核查店小助_最终版
```

若服务器没有 `unzip`，先安装：

```bash
sudo apt install -y unzip
```

确认当前目录中能看到以下文件：

```bash
ls Dockerfile compose.yaml unified_tool douyin_tool meituan_tool
```

## 3. 启动服务

首次部署或代码更新后，构建并后台启动：

```bash
docker compose up -d --build
```

如果构建阶段出现 `Read timed out` 或 `files.pythonhosted.org` 超时，说明服务器访问 Python 官方依赖源不稳定。当前项目已默认使用腾讯云镜像；对已上传的旧项目，可执行以下命令后再重试：

```bash
sed -i 's|pip install --no-cache-dir -r /tmp/requirements.txt|pip install --no-cache-dir --timeout 180 --retries 8 -i https://mirrors.cloud.tencent.com/pypi/simple -r /tmp/requirements.txt|' Dockerfile
docker compose up -d --build
```

查看是否启动成功：

```bash
docker compose ps
docker compose logs -f
```

看到 `Uvicorn running on http://0.0.0.0:8788` 后，按 `Ctrl+C` 退出日志查看（服务仍会在后台运行）。

## 4. 访问系统

浏览器打开：

```text
http://服务器公网IP:8790
```

例如：`http://10.10.10.20:8790`。部署在公司内网时，请使用服务器的内网 IP 或域名。

核对后生成的 Excel 会持久保存到项目目录下的 `runtime/outputs`，重启或升级容器不会丢失。

## 5. 常用运维命令

```bash
# 查看运行状态
docker compose ps

# 查看实时日志
docker compose logs -f

# 重启服务
docker compose restart

# 停止服务（不会删除核对结果）
docker compose down

# 更新代码后重新构建并启动
docker compose up -d --build
```

## 6. 配置域名和 HTTPS（推荐）

如果需要用域名访问并启用 HTTPS，请在服务器前配置 Nginx 或由公司运维使用现有网关反向代理到 `127.0.0.1:8790`。同时将 `compose.yaml` 中的端口改为：

```yaml
ports:
  - "127.0.0.1:8790:8788"
```

这样应用不会直接暴露在公网，只通过公司的 HTTPS 域名访问。
