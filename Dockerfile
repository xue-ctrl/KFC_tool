# 统一核查工具的生产镜像。构建上下文必须是项目根目录，
# 因为 unified_tool 会加载同级的 douyin_tool 和 meituan_tool。
FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

# 先安装依赖，便于利用 Docker 的构建缓存。
COPY unified_tool/requirements.txt /tmp/requirements.txt
# 使用腾讯云镜像并增加超时/重试，避免国内服务器连接官方 PyPI 时超时。
# 如公司网络有内部镜像，可在构建时用 --build-arg PIP_INDEX_URL=... 覆盖。
ARG PIP_INDEX_URL=https://mirrors.cloud.tencent.com/pypi/simple
RUN pip install --no-cache-dir --timeout 180 --retries 8 \
    --index-url "${PIP_INDEX_URL}" \
    -r /tmp/requirements.txt

# 只会复制 .dockerignore 未排除的运行所需代码。
COPY unified_tool /app/unified_tool
COPY douyin_tool /app/douyin_tool
COPY meituan_tool /app/meituan_tool

# 下载生成的 Excel 报告会保存到这里；生产环境由 compose 挂载为持久目录。
RUN mkdir -p /app/unified_tool/data/outputs \
    && useradd --create-home --uid 10001 appuser \
    && chown -R appuser:appuser /app

USER appuser

EXPOSE 8788

CMD ["uvicorn", "unified_tool.app:app", "--host", "0.0.0.0", "--port", "8788", "--workers", "1"]
