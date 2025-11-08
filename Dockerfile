# 使用 Python 官方轻量版镜像
FROM python:3.11-slim

# 设置工作目录
WORKDIR /app

# 复制依赖文件并安装
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 复制代码
COPY . .

# 暴露端口（可选）
EXPOSE 8080

# 启动命令（重点）
CMD ["python", "telegram_checkin_pro_v3.py"]
